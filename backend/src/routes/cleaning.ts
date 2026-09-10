import { Router } from "express";
import { z } from "zod";
import { query, withTransaction } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { enqueueCleaningScanJob, enqueueCleanupExecutionJob, enqueueCloudSyncJob } from "../jobs/queue.js";
import { computeSyncStatus } from "../services/syncStatus.js";
import { ApiError } from "../types/index.js";
import type { OperatorRole } from "../types/index.js";
import type { CloudType } from "../types/connections.js";
import type {
  CleaningChannelRow,
  CleaningChatRow,
  CleaningResourceRow,
  CleaningScanRow,
  CleaningSyncOperation,
  CleaningSyncResourceStatus,
  CleaningTeamsSummary,
  CleanupItemFileRow,
  CleanupManifest,
  CleanupOperationItemRow,
  CleanupOperationRow,
  CleanupOperationStatus,
  CleanupProgress,
  CleanupRecentFile,
  CleanupResourceType,
  CleanupValidationResult,
} from "../types/cleaning.js";

/**
 * Cleaning module (discovery phase) — read-only, reuses the Add Clouds/Manage Clouds `connections`
 * architecture entirely; never touches connections/sync_jobs/connection_users beyond reading them.
 * Mounted outside /api/v1 at a new top-level /api/cleaning prefix, same precedent as /api/clouds.
 */
export const cleaningRouter = Router();
cleaningRouter.use(requireSession);

/**
 * The single security gate for every connection-scoped route in this file — verifies, in one query,
 * that (1) the connection exists, (2) the operator has a tenant_roles row for its tenant, and, when
 * `expectedCloudType` is passed, (3) the connection is actually that cloud type. All three failure
 * modes return the exact same 404 CONNECTION_NOT_FOUND — a SharePoint connection requested through
 * an Outlook route is indistinguishable from a connection that doesn't exist at all, never a
 * distinguishing 400 that would leak "this id is valid, just the wrong type." This is what gives
 * every OneDrive/SharePoint/Teams/Outlook workload its own hard connection-type boundary: a
 * connectionId that resolves to the wrong cloud_type (or a tenant this operator has no role in)
 * never reaches a workload's discovery/service code at all, regardless of which route it was
 * pointed at. Copied in spirit from routes/cloudConnections.ts's own (still separate, still
 * cloud-type-agnostic on purpose — see its docstring) version rather than importing, to avoid
 * touching that file at all.
 */
async function requireConnectionAccess(
  connectionId: string,
  operatorId: string,
  expectedCloudType?: CloudType
): Promise<{ tenantId: string; cloudType: CloudType }> {
  const result = await query<{ tenant_id: string; cloud_type: CloudType; has_access: boolean }>(
    `SELECT c.tenant_id, c.cloud_type, (tr.operator_id IS NOT NULL) AS has_access
     FROM connections c
     LEFT JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $2
     WHERE c.id = $1`,
    [connectionId, operatorId]
  );
  const row = result.rows[0];
  if (!row || !row.has_access) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");
  if (expectedCloudType && row.cloud_type !== expectedCloudType) {
    throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");
  }
  return { tenantId: row.tenant_id, cloudType: row.cloud_type };
}

function toScanRow(r: {
  id: string;
  scan_type: "teams_structure" | "message_counts";
  status: string;
  total_items: number;
  processed_items: number;
  created_at: string;
  started_at: string | null;
  finished_at: string | null;
}): CleaningScanRow {
  return {
    id: r.id,
    scanType: r.scan_type,
    status: r.status as CleaningScanRow["status"],
    totalItems: r.total_items,
    processedItems: r.processed_items,
    createdAt: r.created_at,
    startedAt: r.started_at,
    finishedAt: r.finished_at,
  };
}

/** GET /api/cleaning/connections — connected clouds available for Cleaning, same RBAC scoping as GET /clouds/manage. */
cleaningRouter.get(
  "/connections",
  asyncHandler(async (req, res) => {
    const result = await query<{
      id: string;
      cloud_type: CloudType;
      display_name: string;
      admin_upn: string;
      admin_display_name: string | null;
      status: string;
      last_synced_at: string | null;
    }>(
      `SELECT c.id, c.cloud_type, c.display_name, c.admin_upn, c.admin_display_name, c.status, c.last_synced_at
       FROM connections c
       JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $1
       WHERE c.status != 'disconnected'
       ORDER BY c.display_name, c.cloud_type`,
      [req.session!.operatorId]
    );

    res.json({
      connections: result.rows.map((r) => ({
        id: r.id,
        cloudType: r.cloud_type,
        displayName: r.display_name,
        adminEmail: r.admin_upn,
        adminDisplayName: r.admin_display_name,
        status: r.status,
        lastSyncedAt: r.last_synced_at,
      })),
    });
  })
);

/**
 * Page-number pagination (OFFSET/LIMIT + a total count) rather than keyset/cursor — the UI needs
 * to jump directly to an arbitrary page ("Go to: [3]"), which a forward-only cursor can't do.
 * These tables top out at a few thousand rows (connection_users, cleaning_channels/chats), so the
 * O(offset) cost of OFFSET is not a real concern at this scale.
 */
interface PageResult<T> {
  rows: T[];
  total: number;
  page: number;
  pageSize: number;
}

function parsePageQuery(req: { query: Record<string, unknown> }) {
  const search = typeof req.query.search === "string" && req.query.search.trim() ? req.query.search.trim() : null;
  const sort: "name" | "storage" = req.query.sort === "name" ? "name" : "storage";
  const page = Math.max(1, Number(req.query.page) || 1);
  const pageSize = Math.min(Math.max(1, Number(req.query.pageSize) || 20), 200);
  return { search, sort, page, pageSize };
}

/**
 * Fragment shared by every listCleaningResources-family query below: for one resource row, finds
 * when a permanent-deletion cleanup last completed against it. Joined by resource_id + resource_type
 * (resource_id alone, a gen_random_uuid(), is already effectively unique — resource_type is added
 * defensively, and connection_id purely so the planner can use cleanup_operation_items_connection_idx).
 * resourceType is always a hardcoded call-site literal from the CleanupResourceType union, never
 * request input, so inlining it here (rather than as a bound parameter, which would need every
 * caller's own `$n` numbering to line up) carries no injection risk.
 * This is a per-row LATERAL subquery over a paginated (page-size-bounded) result set, not a
 * full-table scan, so it doesn't introduce the kind of cost that would call for a
 * materialized/precomputed column instead.
 */
function lastPermanentDeleteJoin(resourceType: CleanupResourceType): string {
  return `
     LEFT JOIN LATERAL (
       SELECT MAX(coi.completed_at) AS at
       FROM cleanup_operation_items coi
       JOIN cleanup_operations co ON co.id = coi.cleanup_operation_id
       WHERE coi.resource_id = r.id
         AND coi.connection_id = r.connection_id
         AND coi.resource_type = '${resourceType}'
         AND coi.status = 'completed'
         AND co.deletion_mode = 'permanent'
     ) last_permanent_delete ON true`;
}

/**
 * How long we keep flagging a resource's storage/item figures as possibly not caught up, after its
 * last permanent deletion. Deliberately NOT "has this resource been synced since the delete" —
 * confirmed against real data that a sync run moments *after* a permanent delete completed can still
 * return Graph's not-yet-recalculated quota.used, so "synced later" doesn't prove the number is
 * right. Both windows are heuristics, not guarantees — Microsoft publishes no SLA for this
 * recalculation:
 * - Real-world reports of this lag mostly resolve within a day (STORAGE_RECALC_GRACE_PERIOD_MS) —
 *   the common case, worth a direct "still catching up" note.
 * - Past that but within a week (STORAGE_RECALC_EXTENDED_WINDOW_MS), rather than either silently
 *   dropping the hint (making a still-wrong number look normal) or keeping the same urgency
 *   indefinitely, we taper to a softer "verify independently" suggestion for the rare long tail.
 * - Past a week, we stop flagging entirely, to avoid shadowing old, since-corrected numbers forever.
 */
const STORAGE_RECALC_GRACE_PERIOD_MS = 24 * 60 * 60 * 1000;
const STORAGE_RECALC_EXTENDED_WINDOW_MS = 7 * 24 * 60 * 60 * 1000;

function toDeletionRecalcHint(lastPermanentDeleteAt: string | null): "recent" | "verify" | null {
  if (!lastPermanentDeleteAt) return null;
  const age = Date.now() - new Date(lastPermanentDeleteAt).getTime();
  if (age < STORAGE_RECALC_GRACE_PERIOD_MS) return "recent";
  if (age < STORAGE_RECALC_EXTENDED_WINDOW_MS) return "verify";
  return null;
}

/** Shared by OneDrive, SharePoint, and Outlook Mailboxes — all read connection_users directly, no Graph calls, no job. */
async function listCleaningResources(
  connectionId: string,
  resourceType: CleanupResourceType,
  opts: { search: string | null; sort: "storage" | "name"; page: number; pageSize: number }
): Promise<PageResult<CleaningResourceRow>> {
  const searchClause = `($2::text IS NULL OR r.upn ILIKE '%' || $2 || '%' OR r.display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_users WHERE connection_id = $1 AND ($2::text IS NULL OR upn ILIKE '%' || $2 || '%' OR display_name ILIKE '%' || $2 || '%')`,
    [connectionId, opts.search]
  );
  const total = Number(countResult.rows[0]!.count);

  const result = await query<{
    id: string;
    display_name: string | null;
    upn: string;
    storage_used_bytes: string;
    item_count: number;
    sync_status: string;
    last_synced_at: string | null;
    last_permanent_delete_at: string | null;
  }>(
    `SELECT r.id, r.display_name, r.upn, r.storage_used_bytes, r.item_count, r.sync_status, r.last_synced_at,
            last_permanent_delete.at AS last_permanent_delete_at
     FROM connection_users r
     ${lastPermanentDeleteJoin(resourceType)}
     WHERE r.connection_id = $1 AND ${searchClause}
     ORDER BY ${opts.sort === "storage" ? "r.storage_used_bytes DESC" : "COALESCE(r.display_name, r.upn)"}, r.id
     LIMIT $3 OFFSET $4`,
    [connectionId, opts.search, opts.pageSize, (opts.page - 1) * opts.pageSize]
  );

  const rows: CleaningResourceRow[] = result.rows.map((r) => ({
    id: r.id,
    name: r.display_name ?? r.upn,
    detail: r.upn,
    storageUsedBytes: Number(r.storage_used_bytes),
    itemCount: r.item_count,
    status: r.sync_status as CleaningResourceRow["status"],
    lastSyncedAt: r.last_synced_at,
    deletionRecalcHint: toDeletionRecalcHint(r.last_permanent_delete_at),
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

/** GET /api/cleaning/connections/:id/onedrive — OneDrive Accounts table. */
cleaningRouter.get(
  "/connections/:id/onedrive",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "onedrive");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "onedrive_account", parsePageQuery(req));
    res.json({ accounts: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/sharepoint — SharePoint Sites table. */
cleaningRouter.get(
  "/connections/:id/sharepoint",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "sharepoint");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "sharepoint_site", parsePageQuery(req));
    res.json({ sites: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/outlook — Outlook Mailboxes table. */
cleaningRouter.get(
  "/connections/:id/outlook",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "outlook");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "outlook_mailbox", parsePageQuery(req));
    res.json({ mailboxes: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/google-my-drive — Google My Drive Accounts table. Identical shape to /onedrive above — listCleaningResources already reads generic connection_users columns, no Google-specific query needed. */
cleaningRouter.get(
  "/connections/:id/google-my-drive",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "google_my_drive");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "google_my_drive_account", parsePageQuery(req));
    res.json({ accounts: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/shared-drives — Shared Drives table. Same reuse as google-my-drive above. */
cleaningRouter.get(
  "/connections/:id/shared-drives",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "shared_drive");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "shared_drive", parsePageQuery(req));
    res.json({ drives: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/gmail — Gmail Mailboxes table. Same reuse as google-my-drive above. */
cleaningRouter.get(
  "/connections/:id/gmail",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "gmail");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, "gmail_mailbox", parsePageQuery(req));
    res.json({ mailboxes: rows, total, page, pageSize });
  })
);

/**
 * Google Chat spaces read connection_google_spaces (migrations/015), not connection_users — its
 * own dedicated query, same "no shared-Outlook-helper"-style reasoning as
 * listOutlookCalendarSummaries below. storageUsedBytes is always 0 (not meaningful for Chat);
 * itemCount is member_count here, not message_count — messageCount isn't computed at sync time
 * (see jobs/googleChatSync.ts's comment), so surfacing member_count is the one number this table
 * can show accurately without an expensive per-space enumeration pass.
 */
async function listChatSpaces(
  connectionId: string,
  opts: { search: string | null; sort: "storage" | "name"; page: number; pageSize: number }
): Promise<PageResult<CleaningResourceRow>> {
  const searchClause = `($2::text IS NULL OR display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_google_spaces WHERE connection_id = $1 AND ${searchClause}`,
    [connectionId, opts.search]
  );
  const total = Number(countResult.rows[0]!.count);

  const result = await query<{
    id: string;
    display_name: string | null;
    space_id: string;
    member_count: number;
    sync_status: string;
  }>(
    `SELECT id, display_name, space_id, member_count, sync_status
     FROM connection_google_spaces
     WHERE connection_id = $1 AND ${searchClause}
     ORDER BY ${opts.sort === "storage" ? "member_count DESC" : "COALESCE(display_name, space_id)"}, id
     LIMIT $3 OFFSET $4`,
    [connectionId, opts.search, opts.pageSize, (opts.page - 1) * opts.pageSize]
  );

  const rows: CleaningResourceRow[] = result.rows.map((r) => ({
    id: r.id,
    name: r.display_name ?? r.space_id,
    detail: r.space_id,
    storageUsedBytes: 0,
    itemCount: r.member_count,
    status: r.sync_status as CleaningResourceRow["status"],
    lastSyncedAt: null,
    // Chat message deletion is always effectively permanent (no trash/recycle alternative — see
    // graph/googleChatDeletion.ts), so the deletion-recalc-lag hint doesn't apply here.
    deletionRecalcHint: null,
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

/** GET /api/cleaning/connections/:id/google-chat — Google Chat Spaces table. */
cleaningRouter.get(
  "/connections/:id/google-chat",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "google_chat");
    const { rows, total, page, pageSize } = await listChatSpaces(req.params.id!, parsePageQuery(req));
    res.json({ spaces: rows, total, page, pageSize });
  })
);

/**
 * Reads connection_outlook_calendars — its own dedicated query, not a reuse/generalization of
 * listCleaningResources (which is hardcoded to connection_users), per the no-shared-Outlook-helper
 * rule. Same connection-type gate as the Mail route above: this is a resource view within the same
 * Outlook connection, not a different cloud type.
 */
async function listOutlookCalendarSummaries(
  connectionId: string,
  opts: { search: string | null; sort: "storage" | "name"; page: number; pageSize: number }
): Promise<PageResult<CleaningResourceRow>> {
  const searchClause = `($2::text IS NULL OR upn ILIKE '%' || $2 || '%' OR display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_outlook_calendars WHERE connection_id = $1 AND ${searchClause}`,
    [connectionId, opts.search]
  );
  const total = Number(countResult.rows[0]!.count);

  const result = await query<{
    id: string;
    display_name: string | null;
    upn: string;
    storage_used_bytes: string;
    item_count: number;
    sync_status: string;
    last_synced_at: string | null;
  }>(
    `SELECT id, display_name, upn, storage_used_bytes, item_count, sync_status, last_synced_at
     FROM connection_outlook_calendars
     WHERE connection_id = $1 AND ${searchClause}
     ORDER BY ${opts.sort === "storage" ? "storage_used_bytes DESC" : "COALESCE(display_name, upn)"}, id
     LIMIT $3 OFFSET $4`,
    [connectionId, opts.search, opts.pageSize, (opts.page - 1) * opts.pageSize]
  );

  const rows: CleaningResourceRow[] = result.rows.map((r) => ({
    id: r.id,
    name: r.display_name ?? r.upn,
    detail: r.upn,
    storageUsedBytes: Number(r.storage_used_bytes),
    itemCount: r.item_count,
    status: r.sync_status as CleaningResourceRow["status"],
    lastSyncedAt: r.last_synced_at,
    // Calendar events always go through a plain soft DELETE regardless of an operation's
    // deletion_mode (see cleanupExecutionWorker.ts's executeCalendarItem) — never permanently
    // deleted, so this never applies here.
    deletionRecalcHint: null,
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

cleaningRouter.get(
  "/connections/:id/outlook/calendar",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "outlook");
    const { rows, total, page, pageSize } = await listOutlookCalendarSummaries(req.params.id!, parsePageQuery(req));
    res.json({ calendars: rows, total, page, pageSize });
  })
);

/** Same shape as listOutlookCalendarSummaries above, targeting connection_outlook_contacts — its own dedicated query, not shared with it. */
async function listOutlookContactSummaries(
  connectionId: string,
  opts: { search: string | null; sort: "storage" | "name"; page: number; pageSize: number }
): Promise<PageResult<CleaningResourceRow>> {
  const searchClause = `($2::text IS NULL OR upn ILIKE '%' || $2 || '%' OR display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_outlook_contacts WHERE connection_id = $1 AND ${searchClause}`,
    [connectionId, opts.search]
  );
  const total = Number(countResult.rows[0]!.count);

  const result = await query<{
    id: string;
    display_name: string | null;
    upn: string;
    storage_used_bytes: string;
    item_count: number;
    sync_status: string;
    last_synced_at: string | null;
  }>(
    `SELECT id, display_name, upn, storage_used_bytes, item_count, sync_status, last_synced_at
     FROM connection_outlook_contacts
     WHERE connection_id = $1 AND ${searchClause}
     ORDER BY ${opts.sort === "storage" ? "storage_used_bytes DESC" : "COALESCE(display_name, upn)"}, id
     LIMIT $3 OFFSET $4`,
    [connectionId, opts.search, opts.pageSize, (opts.page - 1) * opts.pageSize]
  );

  const rows: CleaningResourceRow[] = result.rows.map((r) => ({
    id: r.id,
    name: r.display_name ?? r.upn,
    detail: r.upn,
    storageUsedBytes: Number(r.storage_used_bytes),
    itemCount: r.item_count,
    status: r.sync_status as CleaningResourceRow["status"],
    lastSyncedAt: r.last_synced_at,
    // Contacts always go through a plain soft DELETE regardless of an operation's deletion_mode
    // (see cleanupExecutionWorker.ts's executeContactItem) — never permanently deleted, so this
    // never applies here.
    deletionRecalcHint: null,
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

cleaningRouter.get(
  "/connections/:id/outlook/contacts",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "outlook");
    const { rows, total, page, pageSize } = await listOutlookContactSummaries(req.params.id!, parsePageQuery(req));
    res.json({ contacts: rows, total, page, pageSize });
  })
);

/**
 * GET /connections/:id/outlook/overview — one row per mailbox with all three resources' counts
 * together, purely for the selection table (letting a user tick Mail/Calendar/Contacts
 * independently per mailbox from one list instead of three separate tables). This is a read-only
 * SQL join for display only — it never calls Graph and is not part of the discovery, manifest, or
 * execution path, all of which stay three fully separate resource types (see the outlook/mail,
 * outlook/calendar, outlook/contacts routes and resolveManifestItems below, all unchanged). Do not
 * extend this into a generic cross-resource helper for anything beyond rendering this one table.
 */
interface OutlookOverviewSubResource {
  id: string;
  itemCount: number;
  status: "pending" | "synced" | "failed";
}

async function listOutlookOverview(
  connectionId: string,
  opts: { search: string | null; page: number; pageSize: number }
): Promise<
  PageResult<{
    upn: string;
    name: string;
    mail: OutlookOverviewSubResource;
    calendar: OutlookOverviewSubResource | null;
    contacts: OutlookOverviewSubResource | null;
  }>
> {
  const searchClause = `($2::text IS NULL OR cu.upn ILIKE '%' || $2 || '%' OR cu.display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_users cu WHERE cu.connection_id = $1 AND ${searchClause}`,
    [connectionId, opts.search]
  );
  const total = Number(countResult.rows[0]!.count);

  const result = await query<{
    upn: string;
    display_name: string | null;
    mail_id: string;
    mail_items: number;
    mail_status: OutlookOverviewSubResource["status"];
    calendar_id: string | null;
    calendar_events: number | null;
    calendar_status: OutlookOverviewSubResource["status"] | null;
    contacts_id: string | null;
    contact_count: number | null;
    contacts_status: OutlookOverviewSubResource["status"] | null;
  }>(
    `SELECT cu.upn, cu.display_name,
            cu.id AS mail_id, cu.item_count AS mail_items, cu.sync_status AS mail_status,
            cal.id AS calendar_id, cal.item_count AS calendar_events, cal.sync_status AS calendar_status,
            con.id AS contacts_id, con.item_count AS contact_count, con.sync_status AS contacts_status
     FROM connection_users cu
     LEFT JOIN connection_outlook_calendars cal ON cal.connection_id = cu.connection_id AND cal.graph_user_id = cu.graph_user_id
     LEFT JOIN connection_outlook_contacts con ON con.connection_id = cu.connection_id AND con.graph_user_id = cu.graph_user_id
     WHERE cu.connection_id = $1 AND ${searchClause}
     ORDER BY COALESCE(cu.display_name, cu.upn), cu.id
     LIMIT $3 OFFSET $4`,
    [connectionId, opts.search, opts.pageSize, (opts.page - 1) * opts.pageSize]
  );

  const rows = result.rows.map((r) => ({
    upn: r.upn,
    name: r.display_name ?? r.upn,
    mail: { id: r.mail_id, itemCount: r.mail_items, status: r.mail_status },
    calendar: r.calendar_id ? { id: r.calendar_id, itemCount: r.calendar_events ?? 0, status: r.calendar_status! } : null,
    contacts: r.contacts_id ? { id: r.contacts_id, itemCount: r.contact_count ?? 0, status: r.contacts_status! } : null,
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

cleaningRouter.get(
  "/connections/:id/outlook/overview",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "outlook");
    const { search, page, pageSize } = parsePageQuery(req);
    const { rows, total, page: p, pageSize: ps } = await listOutlookOverview(req.params.id!, { search, page, pageSize });
    res.json({ mailboxes: rows, total, page: p, pageSize: ps });
  })
);

async function latestScan(connectionId: string, scanType: "teams_structure" | "message_counts"): Promise<CleaningScanRow | null> {
  const result = await query<{
    id: string; scan_type: "teams_structure" | "message_counts"; status: string;
    total_items: number; processed_items: number; created_at: string; started_at: string | null; finished_at: string | null;
  }>(
    `SELECT id, scan_type, status, total_items, processed_items, created_at, started_at, finished_at
     FROM cleaning_scans WHERE connection_id = $1 AND scan_type = $2 ORDER BY created_at DESC LIMIT 1`,
    [connectionId, scanType]
  );
  return result.rows[0] ? toScanRow(result.rows[0]) : null;
}

const STALLED_AFTER_MS = 10 * 60 * 1000;

/**
 * A scan is worth silently retrying (no user action needed) if it failed outright, or if it's been
 * "running"/"queued" for implausibly long — a worker process restart (e.g. a deploy) can leave a
 * row stuck mid-job forever otherwise, since nothing else ever marks it failed.
 */
function needsRetry(scan: CleaningScanRow): boolean {
  if (scan.status === "failed") return true;
  if (scan.status === "queued" || scan.status === "running") {
    const reference = new Date(scan.startedAt ?? scan.createdAt).getTime();
    return Date.now() - reference > STALLED_AFTER_MS;
  }
  return false;
}

/** True if a not-yet-finished cleanup operation touches this connection — used to avoid a discovery scan racing a live deletion (harmless given cleanup snapshots its own data, but wastes throttle budget and confuses the UI mid-delete). */
async function hasActiveCleanup(connectionId: string): Promise<boolean> {
  const result = await query(
    `SELECT 1 FROM cleanup_operation_items coi
     JOIN cleanup_operations co ON co.id = coi.cleanup_operation_id
     WHERE coi.connection_id = $1 AND co.status IN ('queued', 'running') LIMIT 1`,
    [connectionId]
  );
  return result.rows.length > 0;
}

async function startScan(connectionId: string, scanType: "teams_structure" | "message_counts"): Promise<void> {
  const inserted = await query<{ id: string }>(`INSERT INTO cleaning_scans (connection_id, scan_type) VALUES ($1, $2) RETURNING id`, [
    connectionId,
    scanType,
  ]);
  await enqueueCleaningScanJob({ scanId: inserted.rows[0]!.id });
}

/** GET /api/cleaning/connections/:id/teams/summary — dashboard card + drives the first-visit structure scan. */
cleaningRouter.get(
  "/connections/:id/teams/summary",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "teams");
    const connectionId = req.params.id!;

    let structureScan = await latestScan(connectionId, "teams_structure");
    if ((!structureScan || needsRetry(structureScan)) && !(await hasActiveCleanup(connectionId))) {
      await startScan(connectionId, "teams_structure");
      structureScan = await latestScan(connectionId, "teams_structure");
    }
    const countScan = await latestScan(connectionId, "message_counts");

    const counts = await query<{
      team_count: string; channel_count: string; chat_count: string; messages_counted: string | null;
      awaiting_count: string; failed_count: string;
    }>(
      `SELECT
         (SELECT COUNT(DISTINCT team_id) FROM cleaning_channels WHERE connection_id = $1 AND is_active) AS team_count,
         (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND is_active) AS channel_count,
         (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND is_active) AS chat_count,
         (SELECT COALESCE(SUM(message_count), 0) FROM (
            SELECT message_count FROM cleaning_channels WHERE connection_id = $1 AND is_active AND count_status = 'completed'
            UNION ALL
            SELECT message_count FROM cleaning_chats WHERE connection_id = $1 AND is_active AND count_status = 'completed'
          ) AS m) AS messages_counted,
         (
           (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND is_active AND count_status IN ('pending','calculating')) +
           (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND is_active AND count_status IN ('pending','calculating'))
         ) AS awaiting_count,
         (
           (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND is_active AND count_status = 'failed') +
           (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND is_active AND count_status = 'failed')
         ) AS failed_count`,
      [connectionId]
    );
    const c = counts.rows[0]!;

    const summary: CleaningTeamsSummary = {
      teamCount: Number(c.team_count),
      channelCount: Number(c.channel_count),
      chatCount: Number(c.chat_count),
      messagesCountedSoFar: Number(c.messages_counted ?? 0),
      itemsAwaitingCount: Number(c.awaiting_count),
      itemsFailedCount: Number(c.failed_count),
      structureScan,
      countScan,
    };
    res.json(summary);
  })
);

/**
 * GET /api/cleaning/connections/:id/teams/channels — flat rows; frontend groups by teamId into a
 * tree, which isn't paginated in the UI yet, so this just asks for a generous page size rather
 * than implementing a pager for a view that doesn't have one — still offset-based for consistency
 * and so a future paginated tree view doesn't need an API change.
 */
cleaningRouter.get(
  "/connections/:id/teams/channels",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "teams");
    const { search, page, pageSize } = parsePageQuery(req);
    const searchClause = `($2::text IS NULL OR team_name ILIKE '%' || $2 || '%' OR channel_name ILIKE '%' || $2 || '%')`;

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND is_active AND ${searchClause}`,
      [req.params.id, search]
    );

    const result = await query<{
      id: string; team_id: string; team_name: string; channel_id: string; channel_name: string;
      message_count: number | null; count_status: string;
    }>(
      `SELECT id, team_id, team_name, channel_id, channel_name, message_count, count_status
       FROM cleaning_channels
       WHERE connection_id = $1 AND is_active AND ${searchClause}
       ORDER BY team_name, channel_name, id
       LIMIT $3 OFFSET $4`,
      [req.params.id, search, pageSize, (page - 1) * pageSize]
    );

    const channels: CleaningChannelRow[] = result.rows.map((r) => ({
      id: r.id,
      teamId: r.team_id,
      teamName: r.team_name,
      channelId: r.channel_id,
      channelName: r.channel_name,
      messageCount: r.message_count,
      countStatus: r.count_status as CleaningChannelRow["countStatus"],
    }));
    res.json({ channels, total: Number(countResult.rows[0]!.count), page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/teams/dms */
cleaningRouter.get(
  "/connections/:id/teams/dms",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "teams");
    const { search, page, pageSize } = parsePageQuery(req);
    const searchClause = `($2::text IS NULL OR participants::text ILIKE '%' || $2 || '%')`;

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND is_active AND ${searchClause}`,
      [req.params.id, search]
    );

    const result = await query<{
      id: string; chat_type: string; participants: { displayName: string | null; upn: string | null }[];
      message_count: number | null; count_status: string; last_message_at: string | null;
    }>(
      `SELECT id, chat_type, participants, message_count, count_status, last_message_at
       FROM cleaning_chats
       WHERE connection_id = $1 AND is_active AND ${searchClause}
       ORDER BY last_message_at DESC NULLS LAST, id
       LIMIT $3 OFFSET $4`,
      [req.params.id, search, pageSize, (page - 1) * pageSize]
    );

    const chats: CleaningChatRow[] = result.rows.map((r) => ({
      id: r.id,
      chatType: r.chat_type as CleaningChatRow["chatType"],
      participants: r.participants,
      messageCount: r.message_count,
      countStatus: r.count_status as CleaningChatRow["countStatus"],
      lastMessageAt: r.last_message_at,
    }));
    res.json({ chats, total: Number(countResult.rows[0]!.count), page, pageSize });
  })
);

/** POST /api/cleaning/connections/:id/teams/calculate-counts — kicks off the expensive part, never automatic. */
cleaningRouter.post(
  "/connections/:id/teams/calculate-counts",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId, "teams");
    const connectionId = req.params.id!;

    const running = await query(`SELECT 1 FROM cleaning_scans WHERE connection_id = $1 AND scan_type = 'message_counts' AND status IN ('queued','running') LIMIT 1`, [
      connectionId,
    ]);
    if (running.rows.length > 0) {
      throw new ApiError(409, "SCAN_ALREADY_RUNNING", "Message counting is already in progress for this connection");
    }
    if (await hasActiveCleanup(connectionId)) {
      throw new ApiError(409, "CLEANUP_IN_PROGRESS", "A cleanup is in progress for this connection");
    }

    await startScan(connectionId, "message_counts");
    res.status(202).json({ status: "queued" });
  })
);

/**
 * Cleanup (deletion) execution — see the cleanup-execution plan for the full design rationale.
 * Not connection-scoped in the URL: the existing Review-your-selection screen already aggregates
 * a selection across up to 3 connections of one tenant (OneDrive/SharePoint/Teams), so a cleanup
 * operation must be able to span all of them. The tenant is never accepted from the client — it's
 * derived from, and cross-checked against, every connectionId the manifest actually references.
 */

const manifestSlotSchema = z.object({ connectionId: z.string().uuid(), ids: z.array(z.string().uuid()).min(1) });
const cleanupManifestSchema = z.object({
  oneDrive: manifestSlotSchema.optional(),
  sharePoint: manifestSlotSchema.optional(),
  outlook: manifestSlotSchema.optional(),
  outlookCalendar: manifestSlotSchema.optional(),
  outlookContacts: manifestSlotSchema.optional(),
  channels: manifestSlotSchema.optional(),
  chats: manifestSlotSchema.optional(),
});

// A sibling field to the manifest, not part of CleanupManifest's own structure — chosen by the
// operator on the confirmation screen (components/cleaning/CleanupConfirmation.tsx). Defaults to
// the safer, recoverable behavior if missing/invalid, per the same "default to the safer behavior
// when configuration is absent" principle this app already applies elsewhere (e.g. HTTPS
// enforcement in app.ts) — an operator must explicitly opt into permanent deletion, never the
// reverse. Only OneDrive/SharePoint/Outlook-mail items ever consult this (graph/cleanupDeletion.ts);
// Teams/Calendar/Contacts are unaffected either way.
const deletionModeSchema = z.enum(["recycle_bin", "permanent"]).catch("recycle_bin");

/** `viewer` can validate; only `cleanup_admin` can execute/cancel/retry — mirrors the split routes/cleanup.ts already establishes for the legacy pipeline. */
async function requireCleanupAdmin(tenantId: string, operatorId: string): Promise<void> {
  const result = await query<{ role: OperatorRole }>(`SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2`, [
    tenantId,
    operatorId,
  ]);
  if (result.rows[0]?.role !== "cleanup_admin") {
    throw new ApiError(403, "FORBIDDEN", "Requires cleanup_admin role on this tenant");
  }
}

/** Derives the tenant from every connectionId the manifest references (never accepted from the client) and confirms they all agree — a manifest can never legitimately span two tenants. */
async function resolveManifestTenant(manifest: CleanupManifest, operatorId: string): Promise<string> {
  const connectionIds = [
    ...new Set(
      [
        manifest.oneDrive?.connectionId,
        manifest.sharePoint?.connectionId,
        manifest.outlook?.connectionId,
        manifest.outlookCalendar?.connectionId,
        manifest.outlookContacts?.connectionId,
        manifest.channels?.connectionId,
        manifest.chats?.connectionId,
        manifest.googleMyDrive?.connectionId,
        manifest.sharedDrives?.connectionId,
        manifest.gmail?.connectionId,
        manifest.googleChat?.connectionId,
      ].filter((id): id is string => Boolean(id))
    ),
  ];
  if (connectionIds.length === 0) {
    throw new ApiError(400, "EMPTY_MANIFEST", "Nothing selected");
  }

  const results = await Promise.all(connectionIds.map((id) => requireConnectionAccess(id, operatorId)));
  const tenantIds = new Set(results.map((r) => r.tenantId));
  if (tenantIds.size > 1) {
    // Every tenants row is exactly one directory of exactly one vendor (M365 xor Google — see
    // migrations/014_google_my_drive.sql) — this also naturally rejects a manifest that mixes an
    // M365 connection with a Google connection, not just two different M365 tenants.
    throw new ApiError(400, "TENANT_MISMATCH", "Selected items belong to more than one tenant");
  }
  return results[0]!.tenantId;
}

interface ResolvedManifestItem {
  connectionId: string;
  resourceType: CleanupResourceType;
  resourceId: string;
  displayName: string;
  graphRef: Record<string, string>;
  supported: boolean;
}

interface Queryable {
  query<T extends Record<string, unknown> = Record<string, unknown>>(text: string, params?: unknown[]): Promise<{ rows: T[] }>;
}

/**
 * Re-checks resource ownership (every id must actually belong to the connection it's claimed
 * under) and snapshots the Graph-facing ref + display name needed at execution time. Accepts a
 * `Queryable` so it can run against the plain pool (for /validate) or a transaction client (for
 * the real, TOCTOU-safe snapshot inside POST /cleanup) with identical logic.
 *
 * Re-verifies operator access + cloud type per slot via requireConnectionAccess, even though
 * resolveManifestTenant already checked access for every connectionId in the manifest once,
 * up front — defense-in-depth (each slot is independently authorized at the point it's actually
 * used), not a replacement for that first check.
 */
async function resolveManifestItems(
  manifest: CleanupManifest,
  db: Queryable,
  operatorId: string
): Promise<{ items: ResolvedManifestItem[]; errors: string[]; foundIds: CleanupValidationResult["foundIds"] }> {
  const items: ResolvedManifestItem[] = [];
  const errors: string[] = [];
  const foundIds: CleanupValidationResult["foundIds"] = {
    oneDrive: [],
    sharePoint: [],
    outlook: [],
    outlookCalendar: [],
    outlookContacts: [],
    channels: [],
    chats: [],
    googleMyDrive: [],
    sharedDrives: [],
    gmail: [],
    googleChat: [],
  };

  if (manifest.oneDrive) {
    await requireConnectionAccess(manifest.oneDrive.connectionId, operatorId, "onedrive");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.oneDrive.connectionId, manifest.oneDrive.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.oneDrive.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected OneDrive account is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.oneDrive.connectionId,
        resourceType: "onedrive_account",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userId: row.graph_user_id },
        supported: true,
      });
      foundIds.oneDrive.push(row.id);
    }
  }

  if (manifest.sharePoint) {
    await requireConnectionAccess(manifest.sharePoint.connectionId, operatorId, "sharepoint");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.sharePoint.connectionId, manifest.sharePoint.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.sharePoint.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected SharePoint site is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.sharePoint.connectionId,
        resourceType: "sharepoint_site",
        resourceId: row.id,
        // connection_users.graph_user_id doubles as the site's Graph id for sharepoint-type connections (see cloudSyncWorker.ts).
        displayName: row.display_name ?? row.upn,
        graphRef: { siteId: row.graph_user_id },
        supported: true,
      });
      foundIds.sharePoint.push(row.id);
    }
  }

  if (manifest.outlook) {
    await requireConnectionAccess(manifest.outlook.connectionId, operatorId, "outlook");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.outlook.connectionId, manifest.outlook.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.outlook.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Outlook mailbox is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.outlook.connectionId,
        resourceType: "outlook_mailbox",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userId: row.graph_user_id },
        supported: true,
      });
      foundIds.outlook.push(row.id);
    }
  }

  // Deliberately its own block, not folded into the manifest.outlook block above — same connection
  // type ("outlook"), different resource table/resourceType, per the no-shared-Outlook-helper rule.
  if (manifest.outlookCalendar) {
    await requireConnectionAccess(manifest.outlookCalendar.connectionId, operatorId, "outlook");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_outlook_calendars WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.outlookCalendar.connectionId, manifest.outlookCalendar.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.outlookCalendar.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Outlook calendar is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.outlookCalendar.connectionId,
        resourceType: "outlook_calendar",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userId: row.graph_user_id },
        supported: true,
      });
      foundIds.outlookCalendar.push(row.id);
    }
  }

  // Same reasoning as the outlookCalendar block above — its own block, own table, own resourceType.
  if (manifest.outlookContacts) {
    await requireConnectionAccess(manifest.outlookContacts.connectionId, operatorId, "outlook");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_outlook_contacts WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.outlookContacts.connectionId, manifest.outlookContacts.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.outlookContacts.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Outlook contacts mailbox is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.outlookContacts.connectionId,
        resourceType: "outlook_contacts",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userId: row.graph_user_id },
        supported: true,
      });
      foundIds.outlookContacts.push(row.id);
    }
  }

  if (manifest.channels) {
    await requireConnectionAccess(manifest.channels.connectionId, operatorId, "teams");
    const result = await db.query<{ id: string; team_id: string; team_name: string; channel_id: string; channel_name: string }>(
      `SELECT id, team_id, team_name, channel_id, channel_name FROM cleaning_channels WHERE connection_id = $1 AND is_active AND id = ANY($2::uuid[])`,
      [manifest.channels.connectionId, manifest.channels.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.channels.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Teams channel is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.channels.connectionId,
        resourceType: "channel",
        resourceId: row.id,
        displayName: `${row.team_name} / ${row.channel_name}`,
        graphRef: { teamId: row.team_id, channelId: row.channel_id },
        // Microsoft Graph has no application-permission (unattended) way to delete channel messages — delegated/signed-in-user only.
        supported: false,
      });
      foundIds.channels.push(row.id);
    }
  }

  if (manifest.chats) {
    await requireConnectionAccess(manifest.chats.connectionId, operatorId, "teams");
    const result = await db.query<{ id: string; chat_id: string; participants: { displayName: string | null; upn: string | null }[] }>(
      `SELECT id, chat_id, participants FROM cleaning_chats WHERE connection_id = $1 AND is_active AND id = ANY($2::uuid[])`,
      [manifest.chats.connectionId, manifest.chats.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.chats.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected conversation is no longer available`);
        continue;
      }
      const names = row.participants.map((p) => p.displayName ?? p.upn ?? "Unknown").join(" ↔ ");
      items.push({
        connectionId: manifest.chats.connectionId,
        resourceType: "chat",
        resourceId: row.id,
        displayName: names || "Conversation",
        graphRef: { chatId: row.chat_id },
        // Same Graph limitation as channel messages — no application-permission delete path exists.
        supported: false,
      });
      foundIds.chats.push(row.id);
    }
  }

  // Google My Drive — copy-shaped from the manifest.oneDrive block above; graphRef is keyed
  // userEmail (not userId) since domain-wide-delegation impersonation needs an email subject, not
  // Google's opaque Directory user id.
  if (manifest.googleMyDrive) {
    await requireConnectionAccess(manifest.googleMyDrive.connectionId, operatorId, "google_my_drive");
    const result = await db.query<{ id: string; display_name: string | null; upn: string }>(
      `SELECT id, display_name, upn FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.googleMyDrive.connectionId, manifest.googleMyDrive.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.googleMyDrive.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Google My Drive account is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.googleMyDrive.connectionId,
        resourceType: "google_my_drive_account",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userEmail: row.upn },
        supported: true,
      });
      foundIds.googleMyDrive.push(row.id);
    }
  }

  // Google Shared Drives — same shape as googleMyDrive above, but graphRef is keyed driveId (no
  // owning user to impersonate; executeSharedDriveItem impersonates the connection's admin
  // instead, read from cleanup_operation_items' joined connections row at execution time).
  if (manifest.sharedDrives) {
    await requireConnectionAccess(manifest.sharedDrives.connectionId, operatorId, "shared_drive");
    const result = await db.query<{ id: string; display_name: string | null; upn: string; graph_user_id: string }>(
      `SELECT id, display_name, upn, graph_user_id FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.sharedDrives.connectionId, manifest.sharedDrives.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.sharedDrives.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Shared Drive is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.sharedDrives.connectionId,
        resourceType: "shared_drive",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        // graph_user_id doubles as the Shared Drive's Google id for shared_drive-type connections (see migrations/015).
        graphRef: { driveId: row.graph_user_id },
        supported: true,
      });
      foundIds.sharedDrives.push(row.id);
    }
  }

  // Gmail — same shape as googleMyDrive above, reusing connection_users (graph_user_id = the
  // Directory user id, upn = their primary email, the impersonation subject for message access).
  if (manifest.gmail) {
    await requireConnectionAccess(manifest.gmail.connectionId, operatorId, "gmail");
    const result = await db.query<{ id: string; display_name: string | null; upn: string }>(
      `SELECT id, display_name, upn FROM connection_users WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.gmail.connectionId, manifest.gmail.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.gmail.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Gmail mailbox is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.gmail.connectionId,
        resourceType: "gmail_mailbox",
        resourceId: row.id,
        displayName: row.display_name ?? row.upn,
        graphRef: { userEmail: row.upn },
        supported: true,
      });
      foundIds.gmail.push(row.id);
    }
  }

  // Google Chat — its own table (connection_google_spaces, migrations/015), not connection_users;
  // graphRef is keyed spaceId, resolved to an impersonation target at execution time (a Space has
  // no single owning user — see jobs/googleChatCleanupExecution.ts).
  if (manifest.googleChat) {
    await requireConnectionAccess(manifest.googleChat.connectionId, operatorId, "google_chat");
    const result = await db.query<{ id: string; display_name: string | null; space_id: string }>(
      `SELECT id, display_name, space_id FROM connection_google_spaces WHERE connection_id = $1 AND id = ANY($2::uuid[])`,
      [manifest.googleChat.connectionId, manifest.googleChat.ids]
    );
    const found = new Map(result.rows.map((r) => [r.id, r]));
    for (const id of manifest.googleChat.ids) {
      const row = found.get(id);
      if (!row) {
        errors.push(`A selected Google Chat space is no longer available`);
        continue;
      }
      items.push({
        connectionId: manifest.googleChat.connectionId,
        resourceType: "google_chat_space",
        resourceId: row.id,
        displayName: row.display_name ?? row.space_id,
        graphRef: { spaceId: row.space_id },
        supported: true,
      });
      foundIds.googleChat.push(row.id);
    }
  }

  return { items, errors, foundIds };
}

function summarizeItems(items: ResolvedManifestItem[]): CleanupValidationResult["summary"] {
  return {
    oneDriveAccounts: items.filter((i) => i.resourceType === "onedrive_account").length,
    sharePointSites: items.filter((i) => i.resourceType === "sharepoint_site").length,
    outlookMailboxes: items.filter((i) => i.resourceType === "outlook_mailbox").length,
    outlookCalendars: items.filter((i) => i.resourceType === "outlook_calendar").length,
    outlookContacts: items.filter((i) => i.resourceType === "outlook_contacts").length,
    googleMyDriveAccounts: items.filter((i) => i.resourceType === "google_my_drive_account").length,
    sharedDrives: items.filter((i) => i.resourceType === "shared_drive").length,
    gmailMailboxes: items.filter((i) => i.resourceType === "gmail_mailbox").length,
    googleChatSpaces: items.filter((i) => i.resourceType === "google_chat_space").length,
    channels: items.filter((i) => i.resourceType === "channel").length,
    chats: items.filter((i) => i.resourceType === "chat").length,
  };
}

/** POST /api/cleaning/cleanup/validate — pure dry run, never writes to the database. */
cleaningRouter.post(
  "/cleanup/validate",
  asyncHandler(async (req, res) => {
    const manifest = cleanupManifestSchema.parse(req.body);
    await resolveManifestTenant(manifest, req.session!.operatorId); // authorizes every referenced connection; tenant itself isn't needed for a read-only validation
    const { items, errors, foundIds } = await resolveManifestItems(manifest, { query }, req.session!.operatorId);

    const result: CleanupValidationResult = {
      valid: errors.length === 0,
      summary: summarizeItems(items),
      unsupported: items.filter((i) => !i.supported).map((i) => ({ resourceType: i.resourceType, displayName: i.displayName })),
      errors,
      foundIds,
    };
    res.json(result);
  })
);

/**
 * Tenant-scoped concurrency checks shared by both /cleanup and /sync. Every existing "already
 * running" check in this codebase before this was a plain SELECT-then-INSERT race (confirmed: no
 * advisory locks or FOR UPDATE anywhere) — callers MUST run these only after taking
 * `pg_advisory_xact_lock(hashtext(tenantId))` on the same `db` (i.e. the same transaction client),
 * so the check-then-insert this guards actually is atomic.
 */
async function hasActiveTenantCleanup(db: Queryable, tenantId: string): Promise<boolean> {
  const cleanupOps = await db.query(`SELECT 1 FROM cleanup_operations WHERE tenant_id = $1 AND status IN ('queued', 'running') LIMIT 1`, [
    tenantId,
  ]);
  if (cleanupOps.rows.length > 0) return true;
  // The legacy tenant-scoped Cleanup pipeline (routes/cleanup.ts) is a separate system against the
  // same tenant — without this check it could run concurrently and race overlapping Graph deletes.
  const legacyJobs = await db.query(
    `SELECT 1 FROM cleanup_jobs WHERE tenant_id = $1 AND status IN ('export_in_progress', 'queued', 'running') LIMIT 1`,
    [tenantId]
  );
  return legacyJobs.rows.length > 0;
}

/**
 * Per-connection (not tenant-wide) — each cloud has its own independent "Sync" control, so syncing
 * OneDrive must not be blocked by, or block, a SharePoint/Teams sync for the same tenant. Only
 * `hasActiveTenantSync` (tenant-wide) still guards cleanup, which can span every connection at once.
 */
async function hasActiveConnectionSync(db: Queryable, connectionId: string, cloudType: CloudType): Promise<boolean> {
  if (cloudType === "teams") {
    const result = await db.query(
      `SELECT 1 FROM cleaning_scans WHERE connection_id = $1 AND scan_type = 'teams_structure' AND status IN ('queued', 'running') LIMIT 1`,
      [connectionId]
    );
    return result.rows.length > 0;
  }
  const result = await db.query(`SELECT 1 FROM sync_jobs WHERE connection_id = $1 AND status IN ('queued', 'running') LIMIT 1`, [connectionId]);
  return result.rows.length > 0;
}

async function hasActiveTenantSync(db: Queryable, tenantId: string): Promise<boolean> {
  const result = await db.query(
    `SELECT 1 FROM cleaning_sync_operations so
     LEFT JOIN sync_jobs sj1 ON sj1.id = so.onedrive_sync_job_id
     LEFT JOIN sync_jobs sj2 ON sj2.id = so.sharepoint_sync_job_id
     LEFT JOIN sync_jobs sj3 ON sj3.id = so.outlook_sync_job_id
     LEFT JOIN cleaning_scans cs ON cs.id = so.teams_scan_id
     WHERE so.tenant_id = $1
       AND (sj1.status IN ('queued', 'running') OR sj2.status IN ('queued', 'running') OR sj3.status IN ('queued', 'running') OR cs.status IN ('queued', 'running'))
     LIMIT 1`,
    [tenantId]
  );
  return result.rows.length > 0;
}

/** POST /api/cleaning/cleanup — validates + snapshots + enqueues. Requires cleanup_admin. */
cleaningRouter.post(
  "/cleanup",
  asyncHandler(async (req, res) => {
    const manifest = cleanupManifestSchema.parse(req.body);
    const deletionMode = deletionModeSchema.parse(req.body?.deletionMode);
    const operatorId = req.session!.operatorId;
    const tenantId = await resolveManifestTenant(manifest, operatorId);
    await requireCleanupAdmin(tenantId, operatorId);

    const { operationId, connectionIds } = await withTransaction(async (client) => {
      const db: Queryable = { query: client.query.bind(client) };

      // Atomic per-tenant mutex, held for the rest of this transaction — a concurrent sync/cleanup
      // request for the same tenant blocks here instead of racing the checks below.
      await client.query(`SELECT pg_advisory_xact_lock(hashtext($1))`, [tenantId]);

      if (await hasActiveTenantCleanup(db, tenantId)) {
        throw new ApiError(409, "CLEANUP_ALREADY_RUNNING", "A cleanup is already in progress for this Microsoft 365 connection");
      }
      if (await hasActiveTenantSync(db, tenantId)) {
        throw new ApiError(409, "SYNC_IN_PROGRESS", "A sync is currently running for this connection — try again once it finishes");
      }

      // Re-resolve for real here — never trust any earlier client-side validation for what
      // actually gets written, closing the window between validation and commit.
      const { items: freshItems, errors: freshErrors } = await resolveManifestItems(manifest, db, operatorId);
      if (freshErrors.length > 0) {
        throw new ApiError(400, "VALIDATION_FAILED", "Your selection has changed and needs to be reviewed again", { errors: freshErrors });
      }

      const touchedConnectionIds = [...new Set(freshItems.map((i) => i.connectionId))];
      if (touchedConnectionIds.length > 0) {
        const runningScan = await client.query(
          `SELECT 1 FROM cleaning_scans WHERE connection_id = ANY($1::uuid[]) AND status IN ('queued', 'running') LIMIT 1`,
          [touchedConnectionIds]
        );
        if (runningScan.rows.length > 0) {
          throw new ApiError(409, "SCAN_IN_PROGRESS", "Discovery is still running for this connection — try again once it finishes");
        }
      }

      const unsupportedCount = freshItems.filter((i) => !i.supported).length;
      const opInsert = await client.query<{ id: string }>(
        `INSERT INTO cleanup_operations (tenant_id, requested_by, total_items, skipped_items, deletion_mode) VALUES ($1, $2, $3, $4, $5) RETURNING id`,
        [tenantId, operatorId, freshItems.length, unsupportedCount, deletionMode]
      );
      const newOperationId = opInsert.rows[0]!.id;

      for (const item of freshItems) {
        await client.query(
          `INSERT INTO cleanup_operation_items (cleanup_operation_id, connection_id, resource_type, resource_id, display_name, graph_ref, status)
           VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [
            newOperationId,
            item.connectionId,
            item.resourceType,
            item.resourceId,
            item.displayName,
            JSON.stringify(item.graphRef),
            item.supported ? "pending" : "unsupported",
          ]
        );
      }

      return { operationId: newOperationId, connectionIds: touchedConnectionIds };
    });

    for (const connectionId of connectionIds) {
      await query(`INSERT INTO connection_events (connection_id, tenant_id, event, detail) VALUES ($1, $2, 'cleanup_requested', $3)`, [
        connectionId,
        tenantId,
        { operationId },
      ]);
    }

    await enqueueCleanupExecutionJob({ operationId });
    res.status(202).json({ operationId, status: "queued" });
  })
);

/** Same existence-hiding convention as requireConnectionAccess — 404, never 403, for both "doesn't exist" and "no access". */
async function requireCleanupOperationAccess(operationId: string, operatorId: string, minRole: OperatorRole): Promise<{ tenantId: string }> {
  const result = await query<{ tenant_id: string }>(`SELECT tenant_id FROM cleanup_operations WHERE id = $1`, [operationId]);
  const row = result.rows[0];
  if (!row) throw new ApiError(404, "CLEANUP_OPERATION_NOT_FOUND", "No such cleanup operation");

  const roleResult = await query<{ role: OperatorRole }>(`SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2`, [
    row.tenant_id,
    operatorId,
  ]);
  const role = roleResult.rows[0]?.role;
  if (!role || (minRole === "cleanup_admin" && role !== "cleanup_admin")) {
    throw new ApiError(404, "CLEANUP_OPERATION_NOT_FOUND", "No such cleanup operation");
  }
  return { tenantId: row.tenant_id };
}

interface CleanupOperationSqlRow {
  id: string;
  status: string;
  total_items: number;
  processed_items: number;
  successful_items: number;
  failed_items: number;
  skipped_items: number;
  retry_of_operation_id: string | null;
  started_at: string | null;
  completed_at: string | null;
  cancel_requested_at: string | null;
  created_at: string;
  error_message: string | null;
  requested_by_email: string | null;
  requested_by_display_name: string | null;
  label: string | null;
  deletion_mode: "recycle_bin" | "permanent";
}

/** Shared by the single-operation getter and the list endpoint — same columns, same joins (operators for "Requested by", one of the operation's own touched connections for a human label). */
const CLEANUP_OPERATION_SELECT = `
  co.id, co.status, co.total_items, co.processed_items, co.successful_items, co.failed_items, co.skipped_items,
  co.retry_of_operation_id, co.started_at, co.completed_at, co.cancel_requested_at, co.created_at, co.error_message,
  co.deletion_mode,
  o.email AS requested_by_email, o.display_name AS requested_by_display_name,
  (SELECT c.display_name FROM cleanup_operation_items coi
   JOIN connections c ON c.id = coi.connection_id
   WHERE coi.cleanup_operation_id = co.id LIMIT 1) AS label
`;

function toCleanupOperationRow(r: CleanupOperationSqlRow): CleanupOperationRow {
  return {
    id: r.id,
    status: r.status as CleanupOperationStatus,
    totalItems: r.total_items,
    processedItems: r.processed_items,
    successfulItems: r.successful_items,
    failedItems: r.failed_items,
    skippedItems: r.skipped_items,
    retryOfOperationId: r.retry_of_operation_id,
    startedAt: r.started_at,
    completedAt: r.completed_at,
    cancelRequestedAt: r.cancel_requested_at,
    createdAt: r.created_at,
    errorMessage: r.error_message,
    requestedBy: r.requested_by_email ? { email: r.requested_by_email, displayName: r.requested_by_display_name ?? r.requested_by_email } : null,
    label: r.label ?? "Microsoft 365",
    deletionMode: r.deletion_mode,
  };
}

const RESOURCE_TYPES: CleanupResourceType[] = [
  "onedrive_account",
  "sharepoint_site",
  "outlook_mailbox",
  "outlook_calendar",
  "outlook_contacts",
  "channel",
  "chat",
];

const RESOURCE_TYPE_REPORT_LABEL: Record<CleanupResourceType, string> = {
  onedrive_account: "OneDrive account",
  sharepoint_site: "SharePoint site",
  outlook_mailbox: "Outlook mailbox",
  outlook_calendar: "Outlook calendar event",
  outlook_contacts: "Outlook contact",
  channel: "Teams channel",
  chat: "Direct message",
  google_my_drive_account: "Google My Drive account",
  shared_drive: "Google Shared Drive",
  gmail_mailbox: "Gmail mailbox",
  google_chat_space: "Google Chat space",
};

/** Matches an operation whose touched connections include one with a matching display_name — used by both the list and its count query, so the two never disagree on what "matches" means. */
const OPERATION_SEARCH_CLAUSE = `(
  $3::text IS NULL OR EXISTS (
    SELECT 1 FROM cleanup_operation_items coi
    JOIN connections c ON c.id = coi.connection_id
    WHERE coi.cleanup_operation_id = co.id AND c.display_name ILIKE '%' || $3 || '%'
  )
)`;

/** GET /api/cleaning/cleanup/operations — paginated list of this operator's tenant's cleanup operations, for the Reports page. Optional ?status= and ?search= (matches the same connection label shown as "Cleanup Name") filters. */
cleaningRouter.get(
  "/cleanup/operations",
  asyncHandler(async (req, res) => {
    const { page, pageSize } = parsePageQuery(req);
    const statusFilter = typeof req.query.status === "string" ? req.query.status : null;
    const searchFilter = typeof req.query.search === "string" && req.query.search.trim() ? req.query.search.trim() : null;
    const statusClause = `($2::text IS NULL OR co.status = $2)`;

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*) FROM cleanup_operations co
       JOIN tenant_roles tr ON tr.tenant_id = co.tenant_id AND tr.operator_id = $1
       WHERE ${statusClause} AND ${OPERATION_SEARCH_CLAUSE}`,
      [req.session!.operatorId, statusFilter, searchFilter]
    );
    const result = await query<CleanupOperationSqlRow>(
      `SELECT ${CLEANUP_OPERATION_SELECT}
       FROM cleanup_operations co
       JOIN tenant_roles tr ON tr.tenant_id = co.tenant_id AND tr.operator_id = $1
       LEFT JOIN operators o ON o.id = co.requested_by
       WHERE ${statusClause} AND ${OPERATION_SEARCH_CLAUSE}
       ORDER BY co.created_at DESC
       LIMIT $4 OFFSET $5`,
      [req.session!.operatorId, statusFilter, searchFilter, pageSize, (page - 1) * pageSize]
    );

    res.json({
      operations: result.rows.map(toCleanupOperationRow),
      total: Number(countResult.rows[0]!.count),
      page,
      pageSize,
    });
  })
);

/** GET /api/cleaning/cleanup/operations/summary — aggregate totals across every operation this operator can see, for the Reports page's top stat strip. Same tenant_roles scoping as the list route — never aggregates across a tenant/operator the caller has no role on. */
cleaningRouter.get(
  "/cleanup/operations/summary",
  asyncHandler(async (req, res) => {
    const result = await query<{
      total_operations: string;
      processed_items: string;
      total_items: string;
      bytes_cleared: string;
      bytes_total: string;
    }>(
      `SELECT COUNT(*) AS total_operations,
              COALESCE(SUM(co.processed_items), 0) AS processed_items,
              COALESCE(SUM(co.total_items), 0) AS total_items,
              COALESCE(SUM(f.bytes_cleared), 0) AS bytes_cleared,
              COALESCE(SUM(f.bytes_total), 0) AS bytes_total
       FROM cleanup_operations co
       JOIN tenant_roles tr ON tr.tenant_id = co.tenant_id AND tr.operator_id = $1
       LEFT JOIN LATERAL (
         SELECT SUM(cof.file_size_bytes) FILTER (WHERE cof.status IN ('deleted', 'already_gone')) AS bytes_cleared,
                SUM(cof.file_size_bytes) AS bytes_total
         FROM cleanup_operation_item_files cof
         JOIN cleanup_operation_items coi ON coi.id = cof.cleanup_operation_item_id
         WHERE coi.cleanup_operation_id = co.id
       ) f ON true`,
      [req.session!.operatorId]
    );
    const r = result.rows[0]!;
    res.json({
      totalOperations: Number(r.total_operations),
      processedItems: Number(r.processed_items),
      totalItems: Number(r.total_items),
      bytesCleared: Number(r.bytes_cleared),
      bytesTotal: Number(r.bytes_total),
      updatedAt: new Date().toISOString(),
    });
  })
);

/** Shared by the progress route and the downloadable report, so the two never disagree on totals. */
async function computeCleanupProgress(operationId: string): Promise<CleanupProgress> {
  const opResult = await query<CleanupOperationSqlRow>(
    `SELECT ${CLEANUP_OPERATION_SELECT} FROM cleanup_operations co LEFT JOIN operators o ON o.id = co.requested_by WHERE co.id = $1`,
    [operationId]
  );
  const op = opResult.rows[0];
  if (!op) throw new ApiError(404, "CLEANUP_OPERATION_NOT_FOUND", "No such cleanup operation");

  const byTypeResult = await query<{ resource_type: CleanupResourceType; status: string; count: string }>(
    `SELECT resource_type, status, COUNT(*) AS count FROM cleanup_operation_items WHERE cleanup_operation_id = $1 GROUP BY resource_type, status`,
    [operationId]
  );
  const byType = Object.fromEntries(
    RESOURCE_TYPES.map((t) => [t, { total: 0, completed: 0, failed: 0, skipped: 0, unsupported: 0 }])
  ) as CleanupProgress["byType"];
  for (const r of byTypeResult.rows) {
    const bucket = byType[r.resource_type];
    const count = Number(r.count);
    bucket.total += count;
    if (r.status === "completed") bucket.completed += count;
    else if (r.status === "failed") bucket.failed += count;
    else if (r.status === "skipped") bucket.skipped += count;
    else if (r.status === "unsupported") bucket.unsupported += count;
  }

  const filesResult = await query<{ files_total: string; files_completed: string }>(
    `SELECT COALESCE(SUM(files_total), 0) AS files_total, COALESCE(SUM(files_completed), 0) AS files_completed
     FROM cleanup_operation_items WHERE cleanup_operation_id = $1`,
    [operationId]
  );

  // bytesTotal counts every discovered file (known as soon as it's listed, before deletion is attempted);
  // bytesCleared only 'deleted'/'already_gone' — never 'failed'/'pending' — so it reflects data actually removed.
  const bytesResult = await query<{ bytes_total: string; bytes_cleared: string }>(
    `SELECT COALESCE(SUM(cof.file_size_bytes), 0) AS bytes_total,
            COALESCE(SUM(cof.file_size_bytes) FILTER (WHERE cof.status IN ('deleted', 'already_gone')), 0) AS bytes_cleared
     FROM cleanup_operation_item_files cof
     JOIN cleanup_operation_items coi ON coi.id = cof.cleanup_operation_item_id
     WHERE coi.cleanup_operation_id = $1`,
    [operationId]
  );

  return {
    ...toCleanupOperationRow(op),
    byType,
    filesTotal: Number(filesResult.rows[0]!.files_total),
    filesCompleted: Number(filesResult.rows[0]!.files_completed),
    bytesTotal: Number(bytesResult.rows[0]!.bytes_total),
    bytesCleared: Number(bytesResult.rows[0]!.bytes_cleared),
  };
}

/** GET /api/cleaning/cleanup/:operationId — progress, for the Cleanup Progress screen's polling. */
cleaningRouter.get(
  "/cleanup/:operationId",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");
    res.json(await computeCleanupProgress(req.params.operationId!));
  })
);

/** GET /api/cleaning/cleanup/:operationId/recent-files — live "recently removed" feed for the progress screen. */
cleaningRouter.get(
  "/cleanup/:operationId/recent-files",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");
    const limit = Math.min(Math.max(1, Number(req.query.limit) || 10), 50);

    const result = await query<{ file_name: string; resource_name: string; status: string; completed_at: string }>(
      `SELECT cof.file_name, coi.display_name AS resource_name, cof.status, cof.completed_at
       FROM cleanup_operation_item_files cof
       JOIN cleanup_operation_items coi ON coi.id = cof.cleanup_operation_item_id
       WHERE coi.cleanup_operation_id = $1 AND cof.completed_at IS NOT NULL
       ORDER BY cof.completed_at DESC
       LIMIT $2`,
      [req.params.operationId, limit]
    );

    res.json({
      files: result.rows.map((r) => ({
        fileName: r.file_name,
        resourceName: r.resource_name,
        status: r.status as CleanupRecentFile["status"],
        completedAt: r.completed_at,
      })),
    });
  })
);

/** Human-readable size for the client-facing report — mirrors the frontend's formatBytes (kept separate; this file has no shared dependency on frontend code). */
function formatBytesForReport(bytes: number): string {
  if (bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB", "TB"];
  const exp = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** exp).toFixed(exp === 0 ? 0 : 1)} ${units[exp]}`;
}

/** Explains why an item was never attempted, for rows that have no other error message — otherwise an 'unsupported' row shows blank Notes and reads like an unexplained gap. */
function reportNote(status: string, errorMessage: string | null): string {
  if (errorMessage) return errorMessage;
  if (status === "unsupported") {
    return "Not supported: Microsoft Graph has no application-permission (unattended) path to delete Teams channel/chat messages.";
  }
  return "";
}

/**
 * GET /api/cleaning/cleanup/:operationId/report — downloadable CSV for the client: an operation
 * summary block (cloud, requester, timing, counts, total data cleared), then one row per file
 * (plus one row per skipped/unsupported item, which has no files), then a grand total row.
 */
cleaningRouter.get(
  "/cleanup/:operationId/report",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");

    const progress = await computeCleanupProgress(req.params.operationId!);

    // Whether this specific operation actually performed permanent deletion or an operator-chosen
    // recycle-bin-recoverable soft delete — a fact tied to which code path ran (see
    // migrations/013_cleanup_deletion_mode.sql), not something inferred from status/timestamps, so a
    // "deleted" row from a permanent-mode operation can honestly be reported as "permanently deleted"
    // without mislabeling an operator's own recycle-bin choice (or an old pre-cutover report).
    const isPermanent = progress.deletionMode === "permanent";

    const result = await query<{
      connection_label: string; resource_name: string; resource_type: CleanupResourceType; file_name: string | null;
      file_size_bytes: string | null; status: string; completed_at: string | null; error_message: string | null;
    }>(
      `SELECT c.display_name AS connection_label, coi.display_name AS resource_name, coi.resource_type, cof.file_name,
              cof.file_size_bytes,
              COALESCE(cof.status, coi.status) AS status,
              COALESCE(cof.completed_at, coi.completed_at) AS completed_at,
              COALESCE(cof.error_message, coi.error_message) AS error_message
       FROM cleanup_operation_items coi
       JOIN connections c ON c.id = coi.connection_id
       LEFT JOIN cleanup_operation_item_files cof ON cof.cleanup_operation_item_id = coi.id
       WHERE coi.cleanup_operation_id = $1
       ORDER BY coi.display_name, cof.file_name NULLS FIRST`,
      [req.params.operationId]
    );

    const csvEscape = (value: string) => `"${String(value).replace(/"/g, '""')}"`;
    const csvLine = (values: (string | number)[]) => values.map((v) => csvEscape(String(v))).join(",");

    const summary = [
      csvLine(["Microsoft 365 Cleanup Report"]),
      csvLine(["Cloud / Connection", progress.label]),
      csvLine(["Requested By", progress.requestedBy ? `${progress.requestedBy.displayName} <${progress.requestedBy.email}>` : "—"]),
      csvLine(["Status", progress.status.replace(/_/g, " ")]),
      csvLine(["Started", progress.startedAt ?? "—"]),
      csvLine(["Completed", progress.completedAt ?? "—"]),
      csvLine(["Total Items", progress.totalItems]),
      csvLine(["Successful", progress.successfulItems]),
      csvLine(["Failed", progress.failedItems]),
      csvLine(["Skipped", progress.skippedItems]),
      csvLine(["Total Data Cleared", `${formatBytesForReport(progress.bytesCleared)} (${progress.bytesCleared.toLocaleString()} bytes)`]),
      "",
    ];

    const header = csvLine(["Cloud / Connection", "Resource Type", "Resource Name", "File Name", "Status", "Size", "Completed At", "Notes"]);
    let clearedBytes = 0;
    const rows = result.rows.map((r) => {
      const size = r.file_name != null && r.file_size_bytes != null ? Number(r.file_size_bytes) : null;
      if (size != null && (r.status === "deleted" || r.status === "already_gone")) clearedBytes += size;
      // Only a genuinely permanent-mode operation's "deleted" rows get the stronger wording — an old
      // report re-pulled after this shipped must keep saying what actually happened to it back then.
      const statusText = isPermanent && r.status === "deleted" ? "permanently deleted" : r.status.replace(/_/g, " ");
      return csvLine([
        r.connection_label,
        RESOURCE_TYPE_REPORT_LABEL[r.resource_type],
        r.resource_name,
        r.file_name ?? "",
        statusText,
        size != null ? formatBytesForReport(size) : "",
        r.completed_at ?? "",
        reportNote(r.status, r.error_message),
      ]);
    });
    const totalLine = csvLine(["", "", "", "", "TOTAL DATA CLEARED", formatBytesForReport(clearedBytes), "", ""]);

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="cleanup-report-${req.params.operationId}.csv"`);
    res.send([...summary, header, ...rows, "", totalLine].join("\r\n"));
  })
);

/** GET /api/cleaning/cleanup/:operationId/items — paginated results table, optionally filtered by ?status= and/or ?resourceType=. The resourceType filter backs the progress screen's "expand a category to see its items" drill-down. */
cleaningRouter.get(
  "/cleanup/:operationId/items",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");
    const { page, pageSize } = parsePageQuery(req);
    const statusFilter = typeof req.query.status === "string" ? req.query.status : null;
    const resourceTypeFilter = typeof req.query.resourceType === "string" ? req.query.resourceType : null;
    const filterClause = `($2::text IS NULL OR status = $2) AND ($3::text IS NULL OR resource_type = $3)`;

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*) FROM cleanup_operation_items WHERE cleanup_operation_id = $1 AND ${filterClause}`,
      [req.params.operationId, statusFilter, resourceTypeFilter]
    );
    const result = await query<{
      id: string; connection_id: string; resource_type: CleanupResourceType; display_name: string; status: string;
      attempts: number; started_at: string | null; completed_at: string | null; error_code: string | null; error_message: string | null;
      files_total: number; files_completed: number;
    }>(
      `SELECT id, connection_id, resource_type, display_name, status, attempts, started_at, completed_at, error_code, error_message,
              files_total, files_completed
       FROM cleanup_operation_items
       WHERE cleanup_operation_id = $1 AND ${filterClause}
       ORDER BY display_name, id
       LIMIT $4 OFFSET $5`,
      [req.params.operationId, statusFilter, resourceTypeFilter, pageSize, (page - 1) * pageSize]
    );

    const items: CleanupOperationItemRow[] = result.rows.map((r) => ({
      id: r.id,
      connectionId: r.connection_id,
      resourceType: r.resource_type,
      displayName: r.display_name,
      status: r.status as CleanupOperationItemRow["status"],
      attempts: r.attempts,
      startedAt: r.started_at,
      completedAt: r.completed_at,
      errorCode: r.error_code,
      errorMessage: r.error_message,
      filesTotal: r.files_total,
      filesCompleted: r.files_completed,
    }));
    res.json({ items, total: Number(countResult.rows[0]!.count), page, pageSize });
  })
);

/** GET /api/cleaning/cleanup/:operationId/items/:itemId/files — one item's file list (the third drill-down level: operation → item → file). Joining on coi.cleanup_operation_id = $1 is what stops an itemId from a different operation being read through here. */
cleaningRouter.get(
  "/cleanup/:operationId/items/:itemId/files",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");
    const { page, pageSize } = parsePageQuery(req);

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*)
       FROM cleanup_operation_item_files cof
       JOIN cleanup_operation_items coi ON coi.id = cof.cleanup_operation_item_id
       WHERE coi.cleanup_operation_id = $1 AND coi.id = $2`,
      [req.params.operationId, req.params.itemId]
    );
    const result = await query<{
      id: string; file_name: string; status: string; file_size_bytes: string; error_message: string | null; completed_at: string | null;
    }>(
      `SELECT cof.id, cof.file_name, cof.status, cof.file_size_bytes, cof.error_message, cof.completed_at
       FROM cleanup_operation_item_files cof
       JOIN cleanup_operation_items coi ON coi.id = cof.cleanup_operation_item_id
       WHERE coi.cleanup_operation_id = $1 AND coi.id = $2
       ORDER BY cof.file_name, cof.id
       LIMIT $3 OFFSET $4`,
      [req.params.operationId, req.params.itemId, pageSize, (page - 1) * pageSize]
    );

    const files: CleanupItemFileRow[] = result.rows.map((r) => ({
      id: r.id,
      fileName: r.file_name,
      status: r.status as CleanupItemFileRow["status"],
      fileSizeBytes: Number(r.file_size_bytes),
      errorMessage: r.error_message,
      completedAt: r.completed_at,
    }));
    res.json({ files, total: Number(countResult.rows[0]!.count), page, pageSize });
  })
);

/** POST /api/cleaning/cleanup/:operationId/cancel — cooperative: the worker checks cancel_requested_at between items. */
cleaningRouter.post(
  "/cleanup/:operationId/cancel",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "cleanup_admin");

    const result = await query<{ status: CleanupOperationStatus }>(`SELECT status FROM cleanup_operations WHERE id = $1`, [req.params.operationId]);
    const status = result.rows[0]?.status;
    if (!status || (status !== "queued" && status !== "running")) {
      throw new ApiError(409, "CLEANUP_NOT_CANCELLABLE", "This cleanup has already finished");
    }

    await query(`UPDATE cleanup_operations SET cancel_requested_at = now() WHERE id = $1`, [req.params.operationId]);
    res.status(202).json({ status: "cancel_requested" });
  })
);

/** POST /api/cleaning/cleanup/:operationId/retry — creates a NEW operation scoped to only the previous failed items; the original operation's rows are never mutated. */
cleaningRouter.post(
  "/cleanup/:operationId/retry",
  asyncHandler(async (req, res) => {
    const { tenantId } = await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "cleanup_admin");

    const opResult = await query<{ status: CleanupOperationStatus; deletion_mode: "recycle_bin" | "permanent" }>(
      `SELECT status, deletion_mode FROM cleanup_operations WHERE id = $1`,
      [req.params.operationId]
    );
    const status = opResult.rows[0]?.status;
    if (!status || (status !== "completed_with_errors" && status !== "failed")) {
      throw new ApiError(409, "NOTHING_TO_RETRY", "This cleanup has no failed items to retry");
    }
    // Retry finishes the same job the same way it was configured — never re-asks, and never
    // silently changes an operator's earlier recycle-bin choice into a permanent one (or the reverse).
    const deletionMode = opResult.rows[0]!.deletion_mode;

    const running = await query(`SELECT 1 FROM cleanup_operations WHERE tenant_id = $1 AND status IN ('queued', 'running') LIMIT 1`, [tenantId]);
    if (running.rows.length > 0) {
      throw new ApiError(409, "CLEANUP_ALREADY_RUNNING", "A cleanup is already in progress for this Microsoft 365 connection");
    }

    const failedItems = await query<{
      connection_id: string; resource_type: CleanupResourceType; resource_id: string; display_name: string; graph_ref: Record<string, string>;
    }>(`SELECT connection_id, resource_type, resource_id, display_name, graph_ref FROM cleanup_operation_items WHERE cleanup_operation_id = $1 AND status = 'failed'`, [
      req.params.operationId,
    ]);
    if (failedItems.rows.length === 0) {
      throw new ApiError(409, "NOTHING_TO_RETRY", "This cleanup has no failed items to retry");
    }

    const newOperationId = await withTransaction(async (client) => {
      const opInsert = await client.query<{ id: string }>(
        `INSERT INTO cleanup_operations (tenant_id, requested_by, total_items, retry_of_operation_id, deletion_mode) VALUES ($1, $2, $3, $4, $5) RETURNING id`,
        [tenantId, req.session!.operatorId, failedItems.rows.length, req.params.operationId, deletionMode]
      );
      const newId = opInsert.rows[0]!.id;
      for (const item of failedItems.rows) {
        await client.query(
          `INSERT INTO cleanup_operation_items (cleanup_operation_id, connection_id, resource_type, resource_id, display_name, graph_ref, status)
           VALUES ($1, $2, $3, $4, $5, $6, 'pending')`,
          [newId, item.connection_id, item.resource_type, item.resource_id, item.display_name, JSON.stringify(item.graph_ref)]
        );
      }
      return newId;
    });

    await enqueueCleanupExecutionJob({ operationId: newOperationId });
    res.status(202).json({ operationId: newOperationId, status: "queued" });
  })
);

/**
 * "Sync Now" — a thin tenant-level wrapper around the EXISTING sync_jobs (OneDrive/SharePoint,
 * cloudSyncWorker.ts) and cleaning_scans (Teams, cleaningScanWorker.ts) mechanisms. No new Graph
 * code and no new worker: this only decides which of those to (re)start for a tenant's connections
 * and remembers the resulting ids so the frontend can poll one thing instead of up to three.
 * Not connection-scoped in the URL for the same reason /cleanup isn't — one sync spans a tenant's
 * up-to-3 connections. The tenant is never accepted from the client, only derived from and
 * cross-checked against every connectionId requested (same pattern as resolveManifestTenant).
 */

const syncRequestSchema = z.object({ connectionIds: z.array(z.string().uuid()).min(1) });

/** POST /api/cleaning/sync — kicks off a refresh of the given connections' discovery data. Viewer-level: sync is read/discovery only, not destructive. */
cleaningRouter.post(
  "/sync",
  asyncHandler(async (req, res) => {
    const { connectionIds } = syncRequestSchema.parse(req.body);
    const operatorId = req.session!.operatorId;

    const uniqueConnectionIds = [...new Set(connectionIds)];
    const accessResults = await Promise.all(uniqueConnectionIds.map((id) => requireConnectionAccess(id, operatorId)));
    const tenantIds = new Set(accessResults.map((r) => r.tenantId));
    if (tenantIds.size > 1) {
      throw new ApiError(400, "TENANT_MISMATCH", "Selected connections belong to more than one Microsoft 365 tenant");
    }
    const tenantId = accessResults[0]!.tenantId;

    const connectionRows = await query<{ id: string; cloud_type: CloudType; status: string }>(
      `SELECT id, cloud_type, status FROM connections WHERE id = ANY($1::uuid[])`,
      [uniqueConnectionIds]
    );

    const { operationId, onedriveSyncJobId, sharepointSyncJobId, outlookSyncJobId, teamsScanId } = await withTransaction(async (client) => {
      const db: Queryable = { query: client.query.bind(client) };

      // Same atomic per-tenant mutex /cleanup takes — a concurrent sync/cleanup for this tenant
      // blocks here rather than racing the checks below.
      await client.query(`SELECT pg_advisory_xact_lock(hashtext($1))`, [tenantId]);

      if (await hasActiveTenantCleanup(db, tenantId)) {
        throw new ApiError(409, "CLEANUP_IN_PROGRESS", "A cleanup is in progress for this connection — try again once it finishes");
      }
      for (const conn of connectionRows.rows) {
        if (conn.status === "disconnected") continue;
        if (await hasActiveConnectionSync(db, conn.id, conn.cloud_type)) {
          throw new ApiError(409, "SYNC_ALREADY_RUNNING", "A sync is already in progress for this connection");
        }
      }

      let newOnedriveSyncJobId: string | null = null;
      let newSharepointSyncJobId: string | null = null;
      let newOutlookSyncJobId: string | null = null;
      let newTeamsScanId: string | null = null;

      for (const conn of connectionRows.rows) {
        if (conn.status === "disconnected") continue;
        if (conn.cloud_type === "onedrive" || conn.cloud_type === "sharepoint" || conn.cloud_type === "outlook") {
          // Exact same insert cloudConnections.ts's POST /:id/resync already does — cloudSyncWorker.ts is untouched.
          const jobInsert = await client.query<{ id: string }>(`INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`, [
            conn.id,
          ]);
          if (conn.cloud_type === "onedrive") newOnedriveSyncJobId = jobInsert.rows[0]!.id;
          else if (conn.cloud_type === "sharepoint") newSharepointSyncJobId = jobInsert.rows[0]!.id;
          else newOutlookSyncJobId = jobInsert.rows[0]!.id;
        } else if (conn.cloud_type === "teams") {
          // Same insert startScan() does — decoupled from it here only because startScan() also
          // enqueues immediately, and this needs the id first to record on cleaning_sync_operations
          // before enqueuing after commit (same ordering POST /cleanup already uses).
          const scanInsert = await client.query<{ id: string }>(
            `INSERT INTO cleaning_scans (connection_id, scan_type) VALUES ($1, 'teams_structure') RETURNING id`,
            [conn.id]
          );
          newTeamsScanId = scanInsert.rows[0]!.id;
        }
      }

      if (!newOnedriveSyncJobId && !newSharepointSyncJobId && !newOutlookSyncJobId && !newTeamsScanId) {
        throw new ApiError(400, "NOTHING_TO_SYNC", "None of the selected connections can be synced right now");
      }

      const opInsert = await client.query<{ id: string }>(
        `INSERT INTO cleaning_sync_operations (tenant_id, requested_by, onedrive_sync_job_id, sharepoint_sync_job_id, outlook_sync_job_id, teams_scan_id)
         VALUES ($1, $2, $3, $4, $5, $6) RETURNING id`,
        [tenantId, operatorId, newOnedriveSyncJobId, newSharepointSyncJobId, newOutlookSyncJobId, newTeamsScanId]
      );

      return {
        operationId: opInsert.rows[0]!.id,
        onedriveSyncJobId: newOnedriveSyncJobId,
        sharepointSyncJobId: newSharepointSyncJobId,
        outlookSyncJobId: newOutlookSyncJobId,
        teamsScanId: newTeamsScanId,
      };
    });

    for (const conn of connectionRows.rows) {
      if (conn.status === "disconnected") continue;
      await query(
        `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail) VALUES ($1, $2, 'cleaning_sync_requested', $3, $4)`,
        [conn.id, tenantId, operatorId, { operationId }]
      );
    }

    if (onedriveSyncJobId) await enqueueCloudSyncJob({ syncJobId: onedriveSyncJobId });
    if (sharepointSyncJobId) await enqueueCloudSyncJob({ syncJobId: sharepointSyncJobId });
    if (outlookSyncJobId) await enqueueCloudSyncJob({ syncJobId: outlookSyncJobId });
    if (teamsScanId) await enqueueCleaningScanJob({ scanId: teamsScanId });

    res.status(202).json({ operationId, status: "queued" });
  })
);

/** Same 404-not-403 existence-hiding convention as requireConnectionAccess/requireCleanupOperationAccess. */
async function requireSyncOperationAccess(operationId: string, operatorId: string): Promise<{ tenantId: string }> {
  const result = await query<{ tenant_id: string }>(`SELECT tenant_id FROM cleaning_sync_operations WHERE id = $1`, [operationId]);
  const row = result.rows[0];
  if (!row) throw new ApiError(404, "SYNC_OPERATION_NOT_FOUND", "No such sync operation");

  const access = await query<{ role: OperatorRole }>(`SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2`, [
    row.tenant_id,
    operatorId,
  ]);
  if (!access.rows[0]) throw new ApiError(404, "SYNC_OPERATION_NOT_FOUND", "No such sync operation");
  return { tenantId: row.tenant_id };
}

async function fetchSubResourceStatus(
  table: "sync_jobs" | "cleaning_scans",
  id: string
): Promise<{ status: string; finishedAt: string | null; error: string | null; processed: number; total: number; connectionId: string }> {
  const processedCol = table === "sync_jobs" ? "processed_users" : "processed_items";
  const totalCol = table === "sync_jobs" ? "total_users" : "total_items";
  const result = await query<{
    status: string; finished_at: string | null; error_log: { message: string }[]; processed: number; total: number; connection_id: string;
  }>(`SELECT status, finished_at, error_log, ${processedCol} AS processed, ${totalCol} AS total, connection_id FROM ${table} WHERE id = $1`, [id]);
  const row = result.rows[0]!;
  const errorLog = Array.isArray(row.error_log) ? row.error_log : [];
  return {
    status: row.status,
    finishedAt: row.finished_at,
    error: errorLog.length > 0 ? errorLog[errorLog.length - 1]!.message : null,
    processed: row.processed,
    total: row.total,
    connectionId: row.connection_id,
  };
}

/**
 * "completed_with_errors" for a OneDrive/SharePoint sync_jobs row almost always just means some
 * accounts have no provisioned drive (never touched OneDrive) or Graph reported a tenant-side
 * access block for that specific site — not that the sync mechanism itself broke. Surfacing this
 * count (rather than a bare "some errors" or, worse, a plain X implying total failure) is what lets
 * the Cleaning page explain that distinction instead of alarming the user over normal per-account
 * gaps in the data.
 */
async function countUnavailable(connectionId: string): Promise<number> {
  const result = await query<{ count: string }>(`SELECT COUNT(*) FROM connection_users WHERE connection_id = $1 AND sync_status = 'failed'`, [
    connectionId,
  ]);
  return Number(result.rows[0]!.count);
}

interface SyncOperationRow {
  id: string;
  started_at: string;
  onedrive_sync_job_id: string | null;
  sharepoint_sync_job_id: string | null;
  outlook_sync_job_id: string | null;
  teams_scan_id: string | null;
}

/** Shared by GET /sync/operations/:id and GET /sync/latest — computes the unified view live from whichever sub-resources this operation actually touched. */
async function buildSyncOperationResult(op: SyncOperationRow): Promise<CleaningSyncOperation> {
  const byResource: CleaningSyncOperation["byResource"] = {};
  const subStatuses: string[] = [];
  let completedAt: string | null = null;
  const noteCompletion = (finishedAt: string | null) => {
    if (finishedAt) completedAt = !completedAt || finishedAt > completedAt ? finishedAt : completedAt;
  };

  if (op.onedrive_sync_job_id) {
    const r = await fetchSubResourceStatus("sync_jobs", op.onedrive_sync_job_id);
    byResource.onedrive = {
      status: r.status as CleaningSyncResourceStatus,
      error: r.error,
      processed: r.processed,
      total: r.total,
      unavailableCount: await countUnavailable(r.connectionId),
    };
    subStatuses.push(r.status);
    noteCompletion(r.finishedAt);
  }
  if (op.sharepoint_sync_job_id) {
    const r = await fetchSubResourceStatus("sync_jobs", op.sharepoint_sync_job_id);
    byResource.sharepoint = {
      status: r.status as CleaningSyncResourceStatus,
      error: r.error,
      processed: r.processed,
      total: r.total,
      unavailableCount: await countUnavailable(r.connectionId),
    };
    subStatuses.push(r.status);
    noteCompletion(r.finishedAt);
  }
  if (op.outlook_sync_job_id) {
    const r = await fetchSubResourceStatus("sync_jobs", op.outlook_sync_job_id);
    byResource.outlook = {
      status: r.status as CleaningSyncResourceStatus,
      error: r.error,
      processed: r.processed,
      total: r.total,
      unavailableCount: await countUnavailable(r.connectionId),
    };
    subStatuses.push(r.status);
    noteCompletion(r.finishedAt);
  }
  if (op.teams_scan_id) {
    const r = await fetchSubResourceStatus("cleaning_scans", op.teams_scan_id);
    byResource.teams = { status: r.status as CleaningSyncResourceStatus, error: r.error, processed: r.processed, total: r.total };
    subStatuses.push(r.status);
    noteCompletion(r.finishedAt);
  }

  return {
    id: op.id,
    status: computeSyncStatus(subStatuses),
    startedAt: op.started_at,
    completedAt,
    byResource,
  };
}

const SYNC_OPERATION_COLUMNS = "id, started_at, onedrive_sync_job_id, sharepoint_sync_job_id, outlook_sync_job_id, teams_scan_id";

/** GET /api/cleaning/sync/operations/:operationId — unified status, computed live from whichever sub-resources this operation actually touched. */
cleaningRouter.get(
  "/sync/operations/:operationId",
  asyncHandler(async (req, res) => {
    await requireSyncOperationAccess(req.params.operationId!, req.session!.operatorId);

    const opResult = await query<SyncOperationRow>(`SELECT ${SYNC_OPERATION_COLUMNS} FROM cleaning_sync_operations WHERE id = $1`, [
      req.params.operationId,
    ]);
    res.json(await buildSyncOperationResult(opResult.rows[0]!));
  })
);

/**
 * GET /api/cleaning/sync/latest?connectionIds=a,b,c — the most recent sync that actually touched
 * one of these connections, if any. Lets the Cleaning page resume tracking a sync after navigating
 * away and back (or reloading) — without this, sync progress/status only ever lived in the
 * Dashboard component's local state and was lost the moment it unmounted, even though the sync
 * itself kept running server-side.
 *
 * Deliberately scoped to "an operation whose sub-resource's own connection_id is one of these" —
 * not just "the tenant's latest operation" — since a per-connection CloudSyncControl (one per
 * cloud) must never surface a DIFFERENT connection's sync just because it happens to be the most
 * recent thing this tenant did (e.g. an older bundled operation that also touched another cloud).
 */
cleaningRouter.get(
  "/sync/latest",
  asyncHandler(async (req, res) => {
    const raw = typeof req.query.connectionIds === "string" ? req.query.connectionIds : "";
    const connectionIds = [...new Set(raw.split(",").map((s) => s.trim()).filter(Boolean))];
    if (connectionIds.length === 0) {
      res.json({ operation: null });
      return;
    }

    const operatorId = req.session!.operatorId;
    const accessResults = await Promise.all(connectionIds.map((id) => requireConnectionAccess(id, operatorId)));
    const tenantIds = new Set(accessResults.map((r) => r.tenantId));
    if (tenantIds.size > 1) {
      throw new ApiError(400, "TENANT_MISMATCH", "Selected connections belong to more than one Microsoft 365 tenant");
    }
    const tenantId = accessResults[0]!.tenantId;

    const opResult = await query<SyncOperationRow>(
      `SELECT cso.id, cso.started_at, cso.onedrive_sync_job_id, cso.sharepoint_sync_job_id, cso.outlook_sync_job_id, cso.teams_scan_id
       FROM cleaning_sync_operations cso
       LEFT JOIN sync_jobs sj1 ON sj1.id = cso.onedrive_sync_job_id
       LEFT JOIN sync_jobs sj2 ON sj2.id = cso.sharepoint_sync_job_id
       LEFT JOIN sync_jobs sj3 ON sj3.id = cso.outlook_sync_job_id
       LEFT JOIN cleaning_scans cs ON cs.id = cso.teams_scan_id
       WHERE cso.tenant_id = $1
         AND (sj1.connection_id = ANY($2::uuid[]) OR sj2.connection_id = ANY($2::uuid[]) OR sj3.connection_id = ANY($2::uuid[]) OR cs.connection_id = ANY($2::uuid[]))
       ORDER BY cso.created_at DESC LIMIT 1`,
      [tenantId, connectionIds]
    );
    const op = opResult.rows[0];
    res.json({ operation: op ? await buildSyncOperationResult(op) : null });
  })
);
