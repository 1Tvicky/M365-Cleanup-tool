import { Router } from "express";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { enqueueCleaningScanJob } from "../jobs/queue.js";
import { ApiError } from "../types/index.js";
import type { CloudType } from "../types/connections.js";
import type { CleaningChannelRow, CleaningChatRow, CleaningResourceRow, CleaningScanRow, CleaningTeamsSummary } from "../types/cleaning.js";

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

/** Shared by OneDrive and SharePoint — both read connection_users directly, no Graph calls, no job. */
async function listCleaningResources(connectionId: string, opts: { search: string | null; sort: "storage" | "name"; cursor: string | null; limit: number }) {
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
     WHERE connection_id = $1
       AND ($2::text IS NULL OR upn ILIKE '%' || $2 || '%' OR display_name ILIKE '%' || $2 || '%')
       AND ($3::uuid IS NULL OR id > $3)
     ORDER BY ${opts.sort === "storage" ? "storage_used_bytes DESC" : "display_name NULLS LAST"}, id
     LIMIT $4`,
    [connectionId, opts.search, opts.cursor, opts.limit]
  );

  const rows: CleaningResourceRow[] = result.rows.map((r) => ({
    id: r.id,
    name: r.display_name ?? r.upn,
    detail: r.upn,
    storageUsedBytes: Number(r.storage_used_bytes),
    itemCount: r.item_count,
    status: r.sync_status as CleaningResourceRow["status"],
  }));
  return { rows, nextCursor: rows.length === opts.limit ? rows[rows.length - 1]!.id : null };
}

function parseListQuery(req: { query: Record<string, unknown> }) {
  const search = typeof req.query.search === "string" && req.query.search.trim() ? req.query.search.trim() : null;
  const sort: "name" | "storage" = req.query.sort === "name" ? "name" : "storage";
  const cursor = typeof req.query.cursor === "string" ? req.query.cursor : null;
  const limit = Math.min(Number(req.query.limit) || 50, 200);
  return { search, sort, cursor, limit };
}

/** GET /api/cleaning/connections/:id/onedrive — OneDrive Accounts table. */
cleaningRouter.get(
  "/connections/:id/onedrive",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "onedrive");
    const { rows, nextCursor } = await listCleaningResources(req.params.id!, parseListQuery(req));
    res.json({ accounts: rows, nextCursor });
  })
);

/** GET /api/cleaning/connections/:id/sharepoint — SharePoint Sites table. */
cleaningRouter.get(
  "/connections/:id/sharepoint",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "sharepoint");
    const { rows, nextCursor } = await listCleaningResources(req.params.id!, parseListQuery(req));
    res.json({ sites: rows, nextCursor });
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
    if (!structureScan || needsRetry(structureScan)) {
      await startScan(connectionId, "teams_structure");
      structureScan = await latestScan(connectionId, "teams_structure");
    }
    const countScan = await latestScan(connectionId, "message_counts");

    const counts = await query<{
      team_count: string; channel_count: string; chat_count: string; messages_counted: string | null;
      awaiting_count: string; failed_count: string;
    }>(
      `SELECT
         (SELECT COUNT(DISTINCT team_id) FROM cleaning_channels WHERE connection_id = $1) AS team_count,
         (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1) AS channel_count,
         (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1) AS chat_count,
         (SELECT COALESCE(SUM(message_count), 0) FROM (
            SELECT message_count FROM cleaning_channels WHERE connection_id = $1 AND count_status = 'completed'
            UNION ALL
            SELECT message_count FROM cleaning_chats WHERE connection_id = $1 AND count_status = 'completed'
          ) AS m) AS messages_counted,
         (
           (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND count_status IN ('pending','calculating')) +
           (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND count_status IN ('pending','calculating'))
         ) AS awaiting_count,
         (
           (SELECT COUNT(*) FROM cleaning_channels WHERE connection_id = $1 AND count_status = 'failed') +
           (SELECT COUNT(*) FROM cleaning_chats WHERE connection_id = $1 AND count_status = 'failed')
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

/** GET /api/cleaning/connections/:id/teams/channels — flat rows; frontend groups by teamId. */
cleaningRouter.get(
  "/connections/:id/teams/channels",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
    const { search, cursor, limit } = parseListQuery(req);

    const result = await query<{
      id: string; team_id: string; team_name: string; channel_id: string; channel_name: string;
      message_count: number | null; count_status: string;
    }>(
      `SELECT id, team_id, team_name, channel_id, channel_name, message_count, count_status
       FROM cleaning_channels
       WHERE connection_id = $1
         AND ($2::text IS NULL OR team_name ILIKE '%' || $2 || '%' OR channel_name ILIKE '%' || $2 || '%')
         AND ($3::uuid IS NULL OR id > $3)
       ORDER BY team_name, channel_name, id
       LIMIT $4`,
      [req.params.id, search, cursor, limit]
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
    res.json({ channels, nextCursor: channels.length === limit ? channels[channels.length - 1]!.id : null });
  })
);

/** GET /api/cleaning/connections/:id/teams/dms */
cleaningRouter.get(
  "/connections/:id/teams/dms",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    await resolveConnection(req.params.id!, "teams");
    const { cursor, limit } = parseListQuery(req);

    const result = await query<{
      id: string; chat_type: string; participants: { displayName: string | null; upn: string | null }[];
      message_count: number | null; count_status: string; last_message_at: string | null;
    }>(
      `SELECT id, chat_type, participants, message_count, count_status, last_message_at
       FROM cleaning_chats
       WHERE connection_id = $1 AND ($2::uuid IS NULL OR id > $2)
       ORDER BY last_message_at DESC NULLS LAST, id
       LIMIT $3`,
      [req.params.id, cursor, limit]
    );

    const chats: CleaningChatRow[] = result.rows.map((r) => ({
      id: r.id,
      chatType: r.chat_type as CleaningChatRow["chatType"],
      participants: r.participants,
      messageCount: r.message_count,
      countStatus: r.count_status as CleaningChatRow["countStatus"],
      lastMessageAt: r.last_message_at,
    }));
    res.json({ chats, nextCursor: chats.length === limit ? chats[chats.length - 1]!.id : null });
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

    await startScan(connectionId, "message_counts");
    res.status(202).json({ status: "queued" });
  })
);
