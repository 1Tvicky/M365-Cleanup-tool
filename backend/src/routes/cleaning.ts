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

/** Copied verbatim from routes/cloudConnections.ts rather than importing, to avoid touching that file at all. */
async function requireConnectionAccess(connectionId: string, operatorId: string): Promise<{ tenantId: string }> {
  const result = await query<{ tenant_id: string; has_access: boolean }>(
    `SELECT c.tenant_id, (tr.operator_id IS NOT NULL) AS has_access
     FROM connections c
     LEFT JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $2
     WHERE c.id = $1`,
    [connectionId, operatorId]
  );
  const row = result.rows[0];
  if (!row || !row.has_access) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");
  return { tenantId: row.tenant_id };
}

/** Resolves the Graph-callable tenant GUID for a connection already verified via requireConnectionAccess — never derived from anything frontend-supplied. */
async function resolveConnection(connectionId: string, expectedCloudType: CloudType): Promise<{ m365TenantId: string }> {
  const result = await query<{ cloud_type: CloudType; m365_tenant_id: string }>(
    `SELECT c.cloud_type, t.m365_tenant_id FROM connections c JOIN tenants t ON t.id = c.tenant_id WHERE c.id = $1`,
    [connectionId]
  );
  const row = result.rows[0];
  if (!row) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");
  if (row.cloud_type !== expectedCloudType) {
    throw new ApiError(400, "WRONG_CLOUD_TYPE", `This connection is ${row.cloud_type}, not ${expectedCloudType}`);
  }
  return { m365TenantId: row.m365_tenant_id };
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

/** Shared by OneDrive and SharePoint — both read connection_users directly, no Graph calls, no job. */
async function listCleaningResources(
  connectionId: string,
  opts: { search: string | null; sort: "storage" | "name"; page: number; pageSize: number }
): Promise<PageResult<CleaningResourceRow>> {
  const searchClause = `($2::text IS NULL OR upn ILIKE '%' || $2 || '%' OR display_name ILIKE '%' || $2 || '%')`;

  const countResult = await query<{ count: string }>(
    `SELECT COUNT(*) FROM connection_users WHERE connection_id = $1 AND ${searchClause}`,
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
  }>(
    `SELECT id, display_name, upn, storage_used_bytes, item_count, sync_status
     FROM connection_users
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
  }));

  return { rows, total, page: opts.page, pageSize: opts.pageSize };
}

/** GET /api/cleaning/connections/:id/onedrive — OneDrive Accounts table. */
cleaningRouter.get(
  "/connections/:id/onedrive",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "onedrive");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, parsePageQuery(req));
    res.json({ accounts: rows, total, page, pageSize });
  })
);

/** GET /api/cleaning/connections/:id/sharepoint — SharePoint Sites table. */
cleaningRouter.get(
  "/connections/:id/sharepoint",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "sharepoint");
    const { rows, total, page, pageSize } = await listCleaningResources(req.params.id!, parsePageQuery(req));
    res.json({ sites: rows, total, page, pageSize });
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
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
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
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
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
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
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
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
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
  channels: manifestSlotSchema.optional(),
  chats: manifestSlotSchema.optional(),
});

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
      [manifest.oneDrive?.connectionId, manifest.sharePoint?.connectionId, manifest.channels?.connectionId, manifest.chats?.connectionId].filter(
        (id): id is string => Boolean(id)
      )
    ),
  ];
  if (connectionIds.length === 0) {
    throw new ApiError(400, "EMPTY_MANIFEST", "Nothing selected");
  }

  const results = await Promise.all(connectionIds.map((id) => requireConnectionAccess(id, operatorId)));
  const tenantIds = new Set(results.map((r) => r.tenantId));
  if (tenantIds.size > 1) {
    throw new ApiError(400, "TENANT_MISMATCH", "Selected items belong to more than one Microsoft 365 tenant");
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
 */
async function resolveManifestItems(
  manifest: CleanupManifest,
  db: Queryable
): Promise<{ items: ResolvedManifestItem[]; errors: string[]; foundIds: CleanupValidationResult["foundIds"] }> {
  const items: ResolvedManifestItem[] = [];
  const errors: string[] = [];
  const foundIds: CleanupValidationResult["foundIds"] = { oneDrive: [], sharePoint: [], channels: [], chats: [] };

  if (manifest.oneDrive) {
    await resolveConnection(manifest.oneDrive.connectionId, "onedrive");
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
    await resolveConnection(manifest.sharePoint.connectionId, "sharepoint");
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

  if (manifest.channels) {
    await resolveConnection(manifest.channels.connectionId, "teams");
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
    await resolveConnection(manifest.chats.connectionId, "teams");
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

  return { items, errors, foundIds };
}

function summarizeItems(items: ResolvedManifestItem[]): CleanupValidationResult["summary"] {
  return {
    oneDriveAccounts: items.filter((i) => i.resourceType === "onedrive_account").length,
    sharePointSites: items.filter((i) => i.resourceType === "sharepoint_site").length,
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
    const { items, errors, foundIds } = await resolveManifestItems(manifest, { query });

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
     LEFT JOIN cleaning_scans cs ON cs.id = so.teams_scan_id
     WHERE so.tenant_id = $1
       AND (sj1.status IN ('queued', 'running') OR sj2.status IN ('queued', 'running') OR cs.status IN ('queued', 'running'))
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
      const { items: freshItems, errors: freshErrors } = await resolveManifestItems(manifest, db);
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
        `INSERT INTO cleanup_operations (tenant_id, requested_by, total_items, skipped_items) VALUES ($1, $2, $3, $4) RETURNING id`,
        [tenantId, operatorId, freshItems.length, unsupportedCount]
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

function toCleanupOperationRow(r: {
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
}): CleanupOperationRow {
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
  };
}

const RESOURCE_TYPES: CleanupResourceType[] = ["onedrive_account", "sharepoint_site", "channel", "chat"];

const RESOURCE_TYPE_REPORT_LABEL: Record<CleanupResourceType, string> = {
  onedrive_account: "OneDrive account",
  sharepoint_site: "SharePoint site",
  channel: "Teams channel",
  chat: "Direct message",
};

/** GET /api/cleaning/cleanup/:operationId — progress, for the Cleanup Progress screen's polling. */
cleaningRouter.get(
  "/cleanup/:operationId",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");

    const opResult = await query<{
      id: string; status: string; total_items: number; processed_items: number; successful_items: number;
      failed_items: number; skipped_items: number; retry_of_operation_id: string | null;
      started_at: string | null; completed_at: string | null; cancel_requested_at: string | null; created_at: string; error_message: string | null;
    }>(`SELECT * FROM cleanup_operations WHERE id = $1`, [req.params.operationId]);
    const op = opResult.rows[0];
    if (!op) throw new ApiError(404, "CLEANUP_OPERATION_NOT_FOUND", "No such cleanup operation");

    const byTypeResult = await query<{ resource_type: CleanupResourceType; status: string; count: string }>(
      `SELECT resource_type, status, COUNT(*) AS count FROM cleanup_operation_items WHERE cleanup_operation_id = $1 GROUP BY resource_type, status`,
      [req.params.operationId]
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
      [req.params.operationId]
    );

    const progress: CleanupProgress = {
      ...toCleanupOperationRow(op),
      byType,
      filesTotal: Number(filesResult.rows[0]!.files_total),
      filesCompleted: Number(filesResult.rows[0]!.files_completed),
    };
    res.json(progress);
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

/** GET /api/cleaning/cleanup/:operationId/report — downloadable CSV: one row per file (plus one row per skipped/unsupported item, which has no files). */
cleaningRouter.get(
  "/cleanup/:operationId/report",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");

    const result = await query<{
      resource_name: string; resource_type: CleanupResourceType; file_name: string | null;
      status: string; completed_at: string | null; error_message: string | null;
    }>(
      `SELECT coi.display_name AS resource_name, coi.resource_type, cof.file_name,
              COALESCE(cof.status, coi.status) AS status,
              COALESCE(cof.completed_at, coi.completed_at) AS completed_at,
              COALESCE(cof.error_message, coi.error_message) AS error_message
       FROM cleanup_operation_items coi
       LEFT JOIN cleanup_operation_item_files cof ON cof.cleanup_operation_item_id = coi.id
       WHERE coi.cleanup_operation_id = $1
       ORDER BY coi.display_name, cof.file_name NULLS FIRST`,
      [req.params.operationId]
    );

    const csvEscape = (value: string) => `"${value.replace(/"/g, '""')}"`;
    const header = ["Resource", "Type", "File", "Status", "Completed At", "Details"].map(csvEscape).join(",");
    const rows = result.rows.map((r) =>
      [
        r.resource_name,
        RESOURCE_TYPE_REPORT_LABEL[r.resource_type],
        r.file_name ?? "",
        r.status,
        r.completed_at ?? "",
        r.error_message ?? "",
      ]
        .map((v) => csvEscape(String(v)))
        .join(",")
    );

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="cleanup-report-${req.params.operationId}.csv"`);
    res.send([header, ...rows].join("\r\n"));
  })
);

/** GET /api/cleaning/cleanup/:operationId/items — paginated results table, optionally filtered by ?status=. */
cleaningRouter.get(
  "/cleanup/:operationId/items",
  asyncHandler(async (req, res) => {
    await requireCleanupOperationAccess(req.params.operationId!, req.session!.operatorId, "viewer");
    const { page, pageSize } = parsePageQuery(req);
    const statusFilter = typeof req.query.status === "string" ? req.query.status : null;
    const statusClause = `($2::text IS NULL OR status = $2)`;

    const countResult = await query<{ count: string }>(
      `SELECT COUNT(*) FROM cleanup_operation_items WHERE cleanup_operation_id = $1 AND ${statusClause}`,
      [req.params.operationId, statusFilter]
    );
    const result = await query<{
      id: string; connection_id: string; resource_type: CleanupResourceType; display_name: string; status: string;
      attempts: number; started_at: string | null; completed_at: string | null; error_code: string | null; error_message: string | null;
    }>(
      `SELECT id, connection_id, resource_type, display_name, status, attempts, started_at, completed_at, error_code, error_message
       FROM cleanup_operation_items
       WHERE cleanup_operation_id = $1 AND ${statusClause}
       ORDER BY display_name, id
       LIMIT $3 OFFSET $4`,
      [req.params.operationId, statusFilter, pageSize, (page - 1) * pageSize]
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
    }));
    res.json({ items, total: Number(countResult.rows[0]!.count), page, pageSize });
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

    const opResult = await query<{ status: CleanupOperationStatus }>(`SELECT status FROM cleanup_operations WHERE id = $1`, [req.params.operationId]);
    const status = opResult.rows[0]?.status;
    if (!status || (status !== "completed_with_errors" && status !== "failed")) {
      throw new ApiError(409, "NOTHING_TO_RETRY", "This cleanup has no failed items to retry");
    }

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
        `INSERT INTO cleanup_operations (tenant_id, requested_by, total_items, retry_of_operation_id) VALUES ($1, $2, $3, $4) RETURNING id`,
        [tenantId, req.session!.operatorId, failedItems.rows.length, req.params.operationId]
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

    const { operationId, onedriveSyncJobId, sharepointSyncJobId, teamsScanId } = await withTransaction(async (client) => {
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
      let newTeamsScanId: string | null = null;

      for (const conn of connectionRows.rows) {
        if (conn.status === "disconnected") continue;
        if (conn.cloud_type === "onedrive" || conn.cloud_type === "sharepoint") {
          // Exact same insert cloudConnections.ts's POST /:id/resync already does — cloudSyncWorker.ts is untouched.
          const jobInsert = await client.query<{ id: string }>(`INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`, [
            conn.id,
          ]);
          if (conn.cloud_type === "onedrive") newOnedriveSyncJobId = jobInsert.rows[0]!.id;
          else newSharepointSyncJobId = jobInsert.rows[0]!.id;
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

      if (!newOnedriveSyncJobId && !newSharepointSyncJobId && !newTeamsScanId) {
        throw new ApiError(400, "NOTHING_TO_SYNC", "None of the selected connections can be synced right now");
      }

      const opInsert = await client.query<{ id: string }>(
        `INSERT INTO cleaning_sync_operations (tenant_id, requested_by, onedrive_sync_job_id, sharepoint_sync_job_id, teams_scan_id)
         VALUES ($1, $2, $3, $4, $5) RETURNING id`,
        [tenantId, operatorId, newOnedriveSyncJobId, newSharepointSyncJobId, newTeamsScanId]
      );

      return {
        operationId: opInsert.rows[0]!.id,
        onedriveSyncJobId: newOnedriveSyncJobId,
        sharepointSyncJobId: newSharepointSyncJobId,
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

const SYNC_OPERATION_COLUMNS = "id, started_at, onedrive_sync_job_id, sharepoint_sync_job_id, teams_scan_id";

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
      `SELECT cso.id, cso.started_at, cso.onedrive_sync_job_id, cso.sharepoint_sync_job_id, cso.teams_scan_id
       FROM cleaning_sync_operations cso
       LEFT JOIN sync_jobs sj1 ON sj1.id = cso.onedrive_sync_job_id
       LEFT JOIN sync_jobs sj2 ON sj2.id = cso.sharepoint_sync_job_id
       LEFT JOIN cleaning_scans cs ON cs.id = cso.teams_scan_id
       WHERE cso.tenant_id = $1
         AND (sj1.connection_id = ANY($2::uuid[]) OR sj2.connection_id = ANY($2::uuid[]) OR cs.connection_id = ANY($2::uuid[]))
       ORDER BY cso.created_at DESC LIMIT 1`,
      [tenantId, connectionIds]
    );
    const op = opResult.rows[0];
    res.json({ operation: op ? await buildSyncOperationResult(op) : null });
  })
);
