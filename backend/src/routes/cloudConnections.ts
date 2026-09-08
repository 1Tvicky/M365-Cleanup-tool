import { randomBytes } from "node:crypto";
import { Router } from "express";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireInternalAdmin } from "../middleware/internalAdmin.js";
import { invalidateTenantTokenCache, graphClientForTenant } from "../graph/client.js";
import { exchangeM365ConnectCode, getM365ConnectAuthorizeUrl } from "../services/m365ConnectAuth.js";
import { consumeConnectAttempt, InvalidOAuthStateError, startConnectAttempt } from "../services/oauthState.js";
import { encryptToken } from "../services/tokenEncryption.js";
import { enqueueCloudSyncJob } from "../jobs/queue.js";
import { getSiteById, getTeamById, getUserById, listAllTeams, listAllUsers, searchSites } from "../graph/cloudEnumeration.js";
import { ApiError } from "../types/index.js";
import {
  CLOUD_TYPES,
  isCloudType,
  type AvailableResourceRow,
  type CloudType,
  type ConnectionUserRow,
  type ManageCloudsRow,
  type SyncJobResourceRow,
} from "../types/connections.js";

export const cloudConnectionsRouter = Router();
cloudConnectionsRouter.use(requireSession);

/**
 * Verifies the operator has tenant_roles access to the connection's tenant — same RBAC scoping
 * the rest of the app uses (middleware/rbac.ts), applied here since connections aren't addressed
 * by :tenantId in the URL the way cleanup routes are.
 */
async function requireConnectionAccess(connectionId: string, operatorId: string): Promise<{ tenantId: string; cloudType: CloudType; m365TenantId: string }> {
  const result = await query<{ tenant_id: string; cloud_type: CloudType; m365_tenant_id: string; has_access: boolean }>(
    `SELECT c.tenant_id, c.cloud_type, t.m365_tenant_id, (tr.operator_id IS NOT NULL) AS has_access
     FROM connections c
     JOIN tenants t ON t.id = c.tenant_id
     LEFT JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $2
     WHERE c.id = $1`,
    [connectionId, operatorId]
  );
  const row = result.rows[0];
  if (!row || !row.has_access) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");
  return { tenantId: row.tenant_id, cloudType: row.cloud_type, m365TenantId: row.m365_tenant_id };
}

/**
 * The cheap, listing-only Graph call for a workload's browsable resources — never the expensive
 * per-resource sync calls (getUserDriveQuota/getSiteDriveQuota/etc). cloudType is always the
 * connection's own DB column (via requireConnectionAccess), never taken from the request, so this
 * can't be pointed at the wrong Graph endpoint by a client.
 */
async function listWorkloadResources(client: Awaited<ReturnType<typeof graphClientForTenant>>, cloudType: CloudType): Promise<AvailableResourceRow[]> {
  if (cloudType === "sharepoint") {
    const sites = await searchSites(client);
    return sites.map((s) => ({ id: s.id, displayName: s.displayName, secondary: s.webUrl }));
  }
  if (cloudType === "teams") {
    const teams = await listAllTeams(client);
    return teams.map((t) => ({ id: t.id, displayName: t.displayName }));
  }
  const users = await listAllUsers(client);
  return users.map((u) => ({ id: u.id, displayName: u.displayName ?? u.upn, secondary: u.upn }));
}

/**
 * Direct single-resource lookup, dispatched by the connection's own cloud_type (never the
 * request's) — used to validate/re-derive a *specific selected* resource without ever paginating
 * the whole tenant. This is what lets POST /:id/resync confirm a handful of selected ids belong to
 * this workload without the "enumerate everything, then filter" pattern the resource-level sync
 * feature exists to eliminate.
 */
async function getWorkloadResourceById(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  cloudType: CloudType,
  id: string
): Promise<AvailableResourceRow | null> {
  if (cloudType === "sharepoint") {
    const site = await getSiteById(client, id);
    return site ? { id: site.id, displayName: site.displayName, secondary: site.webUrl } : null;
  }
  if (cloudType === "teams") {
    const team = await getTeamById(client, id);
    return team ? { id: team.id, displayName: team.displayName } : null;
  }
  const user = await getUserById(client, id);
  return user ? { id: user.id, displayName: user.displayName ?? user.upn, secondary: user.upn } : null;
}

/**
 * Short-TTL in-process cache for the browse endpoint — same pattern as graph/client.ts's tenant
 * token cache (a plain Map keyed by a string id, entries carrying their own expiry), a separate
 * instance since this caches resource *listings* per connection, not Graph access tokens per
 * tenant. Avoids re-listing (paginating GET /users or /sites, potentially thousands of rows) on
 * every keystroke of search or page change — never persisted into connection_users, which stays
 * reserved for "state as of the last real sync," not "what Graph currently reports."
 */
const availableResourcesCache = new Map<string, { resources: AvailableResourceRow[]; expiresAt: number }>();
const AVAILABLE_RESOURCES_TTL_MS = 90_000;

async function getCachedWorkloadResources(
  connectionId: string,
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  cloudType: CloudType
): Promise<AvailableResourceRow[]> {
  const cached = availableResourcesCache.get(connectionId);
  if (cached && cached.expiresAt > Date.now()) return cached.resources;
  const resources = await listWorkloadResources(client, cloudType);
  availableResourcesCache.set(connectionId, { resources, expiresAt: Date.now() + AVAILABLE_RESOURCES_TTL_MS });
  return resources;
}

/** POST /api/clouds/:cloudType/connect/init — see docs/azure-ad-app-registration.md §4a. */
cloudConnectionsRouter.post(
  "/:cloudType/connect/init",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const cloudType = req.params.cloudType;
    if (!cloudType || !isCloudType(cloudType)) {
      throw new ApiError(400, "INVALID_CLOUD_TYPE", `cloudType must be one of: ${CLOUD_TYPES.join(", ")}`);
    }

    const attempt = await startConnectAttempt(cloudType, req.session!.operatorId);
    const authorizeUrl = await getM365ConnectAuthorizeUrl({
      state: attempt.authorizeParams.state,
      codeChallenge: attempt.authorizeParams.codeChallenge,
      codeChallengeMethod: attempt.authorizeParams.codeChallengeMethod,
    });

    res.json({ authorizeUrl, state: attempt.state });
  })
);

/**
 * GET /api/clouds/manage — one row per (tenant, cloud_type) connection the operator can see.
 *
 * total_users/processed_users used to always come from the single most-recently-created sync_jobs
 * row (DISTINCT ON (c.id) ... ORDER BY sj.created_at DESC). That stopped being sufficient once
 * resource-scoped syncs can run concurrently on the same connection (see POST /:id/resync below) —
 * an older-but-still-running scoped job would otherwise silently disappear from this figure the
 * moment a second one starts. So: when a connection has any sync_jobs row still queued/running,
 * SUM across every such row instead of picking just one; otherwise fall back to the single latest
 * row exactly as before (covers both "no job ever ran" and "last job already finished").
 */
cloudConnectionsRouter.get(
  "/manage",
  asyncHandler(async (req, res) => {
    const result = await query<{
      id: string;
      cloud_type: CloudType;
      display_name: string;
      admin_upn: string;
      admin_display_name: string | null;
      status: ManageCloudsRow["status"];
      connected_at: string | null;
      last_synced_at: string | null;
      last_error: string | null;
      total_users: number | null;
      processed_users: number | null;
      total_known_users: string | null;
      added_users: string | null;
      not_added_users: string | null;
    }>(
      `SELECT c.id, c.cloud_type, c.display_name, c.admin_upn, c.admin_display_name, c.status,
              c.connected_at, c.last_synced_at, c.last_error,
              COALESCE(active.total_users, latest.total_users) AS total_users,
              COALESCE(active.processed_users, latest.processed_users) AS processed_users,
              cu.total_known_users, cu.added_users, cu.not_added_users
       FROM connections c
       JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $1
       LEFT JOIN LATERAL (
         SELECT SUM(total_users) AS total_users, SUM(processed_users) AS processed_users
         FROM sync_jobs sj
         WHERE sj.connection_id = c.id AND sj.status IN ('queued', 'running')
       ) active ON active.total_users IS NOT NULL
       LEFT JOIN LATERAL (
         SELECT sj.total_users, sj.processed_users
         FROM sync_jobs sj
         WHERE sj.connection_id = c.id
         ORDER BY sj.created_at DESC
         LIMIT 1
       ) latest ON true
       LEFT JOIN (
         SELECT connection_id,
                COUNT(*) AS total_known_users,
                COUNT(*) FILTER (WHERE sync_status = 'synced') AS added_users,
                COUNT(*) FILTER (WHERE sync_status = 'failed') AS not_added_users
         FROM connection_users
         GROUP BY connection_id
       ) cu ON cu.connection_id = c.id
       WHERE c.status != 'disconnected'`,
      [req.session!.operatorId]
    );

    const connections: ManageCloudsRow[] = result.rows.map((r) => {
      // totalUsers is sourced from connection_users (the real, discovered count), never from
      // sync_jobs.total_users directly — a sync job's own total is per-job bookkeeping (e.g.
      // Outlook's syncOutlook runs Mail/Calendar/Contacts as separate internal passes over the
      // same job row) and isn't guaranteed to stay 1:1 with "how many users/sites we actually
      // know about" the way this row's headline count needs to be.
      const total = Number(r.total_known_users ?? r.total_users ?? 0);
      const processed = r.processed_users ?? 0;
      const added = Number(r.added_users ?? 0);
      return {
        id: r.id,
        cloudType: r.cloud_type,
        iconKey: r.cloud_type,
        displayName: r.display_name,
        adminEmail: r.admin_upn,
        adminDisplayName: r.admin_display_name,
        tenantDomain: r.display_name,
        totalUsers: total,
        processedUsers: processed,
        addedUsers: added,
        notAddedUsers: Number(r.not_added_users ?? 0),
        // Deliberately addedUsers/total, not processedUsers/total: processedUsers is scoped to
        // whichever single sync_jobs run happens to be "latest" for this connection, which is now
        // routinely a small resource-scoped run (e.g. 2 of 116) once resource-level sync is the
        // common case — dividing that by the whole tenant's total_known_users would make this
        // figure crash toward 0% after every scoped sync, even though addedUsers (this workload's
        // real, current, whole-tenant synced count from connection_users) hasn't actually dropped.
        // addedUsers already updates per-resource in real time during an active job too, so this
        // stays accurate whether idle or mid-sync, unlike processedUsers/total.
        percent: total > 0 ? Math.round((added / total) * 100) : 0,
        status: r.status,
        multiUser: true,
        connectedAt: r.connected_at,
        lastSyncedAt: r.last_synced_at,
        lastError: r.last_error,
      };
    });

    res.json({ connections });
  })
);

/**
 * GET /api/clouds/:id/status — lightweight polling target for the OAuth popup itself (see
 * popupResultPage below): same-origin, same browser, so the operator's session cookie is sent
 * automatically. Lets the popup show real enumeration progress instead of declaring "connected"
 * the instant the token exchange finishes, before any users have actually been discovered.
 */
cloudConnectionsRouter.get(
  "/:id/status",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const result = await query<{
      status: ManageCloudsRow["status"];
      job_status: string | null;
      total_users: number | null;
      processed_users: number | null;
    }>(
      `SELECT c.status,
              sj.status AS job_status, sj.total_users, sj.processed_users
       FROM connections c
       LEFT JOIN sync_jobs sj ON sj.connection_id = c.id
       WHERE c.id = $1
       ORDER BY sj.created_at DESC NULLS LAST
       LIMIT 1`,
      [req.params.id]
    );
    const row = result.rows[0];
    if (!row) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No such connection");

    res.json({
      connectionStatus: row.status,
      job: row.job_status
        ? { status: row.job_status, totalUsers: row.total_users ?? 0, processedUsers: row.processed_users ?? 0 }
        : null,
    });
  })
);

/**
 * GET /api/clouds/:id/available-resources — the browse step for resource-level sync: every
 * resource this workload currently has in Microsoft 365 (live, via the cheap listing-only Graph
 * calls, cached briefly — see getCachedWorkloadResources), NOT what's already synced. Deliberately
 * separate from GET /:id/users (which reads connection_users, i.e. "as of the last sync" and can be
 * empty/stale before a first sync or right after the Teams resource-shape change). Search/paging
 * happen server-side over the cached array, using the same page/pageSize convention as
 * routes/cleaning.ts's discovery routes (not GET /:id/users' cursor convention) so this plugs
 * directly into the existing DiscoveryTable component on the frontend.
 */
cloudConnectionsRouter.get(
  "/:id/available-resources",
  asyncHandler(async (req, res) => {
    const { cloudType, m365TenantId } = await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const client = await graphClientForTenant(m365TenantId);
    const all = await getCachedWorkloadResources(req.params.id!, client, cloudType);

    const search = typeof req.query.search === "string" ? req.query.search.trim().toLowerCase() : "";
    const filtered = search
      ? all.filter((r) => r.displayName.toLowerCase().includes(search) || r.secondary?.toLowerCase().includes(search))
      : all;

    const pageSize = Math.min(Number(req.query.pageSize) || 20, 200);
    const page = Math.max(Number(req.query.page) || 1, 1);
    const resources = filtered.slice((page - 1) * pageSize, page * pageSize);

    res.json({ resources, total: filtered.length, page, pageSize });
  })
);

/**
 * GET /api/clouds/:id/sync-jobs/:jobId/resources — per-resource progress for one resource-scoped
 * sync run, for the live "N of M completed" view. A job with no rows here is either pre-this-change
 * history or a legacy full-tenant run (see jobs/cloudSyncWorker.ts) — the frontend falls back to the
 * existing aggregate total_users/processed_users bar in that case, it doesn't treat an empty list as
 * an error.
 */
cloudConnectionsRouter.get(
  "/:id/sync-jobs/:jobId/resources",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const jobRow = await query(`SELECT 1 FROM sync_jobs WHERE id = $1 AND connection_id = $2`, [req.params.jobId, req.params.id]);
    if (jobRow.rows.length === 0) throw new ApiError(404, "SYNC_JOB_NOT_FOUND", "No such sync job for this connection");

    const result = await query<{
      id: string;
      graph_resource_id: string;
      display_name: string;
      secondary: string | null;
      status: SyncJobResourceRow["status"];
      error_message: string | null;
      started_at: string | null;
      completed_at: string | null;
    }>(
      `SELECT id, graph_resource_id, display_name, secondary, status, error_message, started_at, completed_at
       FROM sync_job_resources
       WHERE sync_job_id = $1
       ORDER BY display_name`,
      [req.params.jobId]
    );

    const resources: SyncJobResourceRow[] = result.rows.map((r) => ({
      id: r.id,
      graphResourceId: r.graph_resource_id,
      displayName: r.display_name,
      secondary: r.secondary,
      status: r.status,
      errorMessage: r.error_message,
      startedAt: r.started_at,
      completedAt: r.completed_at,
    }));

    res.json({ resources });
  })
);

/**
 * GET /api/clouds/:id/users — backs the expand chevron's "Failed Users Details" drill-in.
 * `?status=failed` filters to just the not-added users (the common case from the UI); omit for
 * the full list.
 */
cloudConnectionsRouter.get(
  "/:id/users",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const limit = Math.min(Number(req.query.limit) || 50, 200);
    const cursor = typeof req.query.cursor === "string" ? req.query.cursor : null;
    const statusFilter = req.query.status === "failed" || req.query.status === "synced" || req.query.status === "pending" ? req.query.status : null;

    const result = await query<{
      id: string;
      graph_user_id: string;
      upn: string;
      display_name: string | null;
      storage_used_bytes: string;
      item_count: number;
      sync_status: ConnectionUserRow["syncStatus"];
      last_synced_at: string | null;
      error_message: string | null;
    }>(
      `SELECT id, graph_user_id, upn, display_name, storage_used_bytes, item_count, sync_status, last_synced_at, error_message
       FROM connection_users
       WHERE connection_id = $1 AND ($2::uuid IS NULL OR id > $2) AND ($4::text IS NULL OR sync_status = $4)
       ORDER BY id
       LIMIT $3`,
      [req.params.id, cursor, limit, statusFilter]
    );

    const users: ConnectionUserRow[] = result.rows.map((r) => ({
      id: r.id,
      graphUserId: r.graph_user_id,
      upn: r.upn,
      displayName: r.display_name,
      storageUsedBytes: Number(r.storage_used_bytes),
      itemCount: r.item_count,
      syncStatus: r.sync_status,
      lastSyncedAt: r.last_synced_at,
      errorMessage: r.error_message,
    }));

    res.json({ users, nextCursor: users.length === limit ? users[users.length - 1]!.id : null });
  })
);

function csvEscape(value: string): string {
  return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** GET /api/clouds/:id/users/export — the full per-user list as a CSV download, no pagination cap (the on-screen table pages, the export doesn't). */
cloudConnectionsRouter.get(
  "/:id/users/export",
  asyncHandler(async (req, res) => {
    await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const result = await query<{
      upn: string;
      display_name: string | null;
      storage_used_bytes: string;
      item_count: number;
      sync_status: ConnectionUserRow["syncStatus"];
      error_message: string | null;
    }>(
      `SELECT upn, display_name, storage_used_bytes, item_count, sync_status, error_message
       FROM connection_users
       WHERE connection_id = $1
       ORDER BY COALESCE(display_name, upn)`,
      [req.params.id]
    );

    const header = ["Name", "Email", "Storage Used (bytes)", "Item Count", "Status", "Notes"].map(csvEscape).join(",");
    const rows = result.rows.map((r) =>
      [
        r.display_name ?? r.upn,
        r.upn,
        r.storage_used_bytes,
        r.item_count,
        r.sync_status === "synced" ? "Active" : r.sync_status === "failed" ? "Inactive" : "Pending",
        r.error_message ?? "",
      ]
        .map((v) => csvEscape(String(v)))
        .join(",")
    );

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="users-${req.params.id}.csv"`);
    res.send([header, ...rows].join("\r\n"));
  })
);

/**
 * POST /api/clouds/:id/resync — creates a fresh sync_jobs row rather than reusing the last one, so
 * sync history isn't overwritten. connections.last_synced_at updates on completion, not here — see
 * docs/cloud-connections-api.md for why that deviates from a literal "update on resync" reading.
 *
 * Optional `resourceIds` in the body scopes the sync to just those resources (the "Sync Selected"
 * flow) — omitted/empty preserves today's full-tenant behavior exactly (used by the plain "Resync"
 * icon, which stays a quick "sync everything" action). Resource ids are never trusted as-is: they're
 * intersected against a fresh live listing (the same one GET /:id/available-resources uses) keyed by
 * this connection's own cloud_type — never anything the client claims — so a ResourceId that
 * belongs to a different connection, tenant, or workload is silently excluded (reported in
 * `skipped`), never synced.
 */
cloudConnectionsRouter.post(
  "/:id/resync",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const { tenantId, cloudType, m365TenantId } = await requireConnectionAccess(req.params.id!, req.session!.operatorId);
    const requestedIds: string[] = Array.isArray(req.body?.resourceIds) ? req.body.resourceIds.filter((v: unknown) => typeof v === "string") : [];

    const connRow = await query<{ status: string }>(`SELECT status FROM connections WHERE id = $1`, [req.params.id]);
    const status = connRow.rows[0]?.status;
    if (status === "disconnected") {
      throw new ApiError(409, "CONNECTION_DISCONNECTED", "This connection is disconnected — reconnect instead of resyncing");
    }

    // Tier 1: a legacy/full-tenant job (no sync_job_resources rows of its own) touches every
    // resource, so it still blocks everything regardless of what's being requested now.
    const legacyRunning = await query(
      `SELECT 1 FROM sync_jobs sj
       WHERE sj.connection_id = $1 AND sj.status IN ('queued', 'running')
         AND NOT EXISTS (SELECT 1 FROM sync_job_resources sjr WHERE sjr.sync_job_id = sj.id)
       LIMIT 1`,
      [req.params.id]
    );
    if (legacyRunning.rows.length > 0) {
      throw new ApiError(409, "RESYNC_ALREADY_RUNNING", "A sync job is already in progress for this connection");
    }

    const acceptedResources: AvailableResourceRow[] = [];
    const skipped: { id: string; reason: "not_found" | "already_syncing" }[] = [];

    if (requestedIds.length > 0) {
      // Tier 2: a resource-scoped job only blocks the specific resource ids it's already working on.
      const alreadySyncing = await query<{ graph_resource_id: string }>(
        `SELECT DISTINCT sjr.graph_resource_id
         FROM sync_job_resources sjr
         JOIN sync_jobs sj ON sj.id = sjr.sync_job_id
         WHERE sj.connection_id = $1 AND sj.status IN ('queued', 'running')
           AND sjr.status IN ('pending', 'processing')
           AND sjr.graph_resource_id = ANY($2::text[])`,
        [req.params.id, requestedIds]
      );
      const lockedIds = new Set(alreadySyncing.rows.map((r) => r.graph_resource_id));

      // Direct per-id lookups, not a full tenant listing — a tenant with thousands of users
      // selecting 3 must only ever cost 3 Graph calls here, never one that pages the whole tenant.
      const client = await graphClientForTenant(m365TenantId);
      for (const id of requestedIds) {
        if (lockedIds.has(id)) {
          skipped.push({ id, reason: "already_syncing" });
          continue;
        }
        const resource = await getWorkloadResourceById(client, cloudType, id);
        if (resource) {
          acceptedResources.push(resource);
        } else {
          skipped.push({ id, reason: "not_found" });
        }
      }
      if (acceptedResources.length === 0) {
        throw new ApiError(400, "NOTHING_TO_SYNC", "None of the selected resources could be synced");
      }
    }
    // requestedIds empty/absent falls through with acceptedResources still [] — jobs/cloudSyncWorker.ts
    // treats a sync_jobs row with zero sync_job_resources children as "sync everything," which is
    // exactly today's behavior and exactly what the plain "Resync" icon (never sends resourceIds)
    // needs. The frontend's "Sync Selected" button is disabled at 0 selected, so a legitimate caller
    // never reaches this branch meaning to scope to nothing.

    const jobInsert = await query<{ id: string }>(
      `INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`,
      [req.params.id]
    );
    const jobId = jobInsert.rows[0]!.id;

    if (acceptedResources.length > 0) {
      for (const r of acceptedResources) {
        await query(
          `INSERT INTO sync_job_resources (sync_job_id, connection_id, graph_resource_id, display_name, secondary) VALUES ($1, $2, $3, $4, $5)`,
          [jobId, req.params.id, r.id, r.displayName, r.secondary ?? null]
        );
      }
    }

    await query(
      `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail)
       VALUES ($1, $2, 'resync_requested', $3, $4)`,
      [req.params.id, tenantId, req.session!.operatorId, { syncJobId: jobId, resourceCount: acceptedResources.length }]
    );

    await enqueueCloudSyncJob({ syncJobId: jobId });
    res.status(202).json({ jobId, status: "queued", acceptedCount: acceptedResources.length, skipped });
  })
);

/** DELETE /api/clouds/:id — soft disconnect. See docs/azure-ad-app-registration.md §6 for exactly what this can/can't revoke. */
cloudConnectionsRouter.delete(
  "/:id",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const { tenantId } = await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    await query(
      `UPDATE sync_jobs SET cancel_requested_at = now()
       WHERE connection_id = $1 AND status IN ('queued', 'running')`,
      [req.params.id]
    );

    await query(
      `UPDATE connections
       SET status = 'disconnected', disconnected_at = now(), encrypted_refresh_token = NULL
       WHERE id = $1`,
      [req.params.id]
    );

    await query(
      `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id)
       VALUES ($1, $2, 'disconnected', $3)`,
      [req.params.id, tenantId, req.session!.operatorId]
    );

    res.status(204).end();
  })
);

/* --- OAuth callback: mounted separately in app.ts at the fixed top-level path
   /api/auth/m365/callback, since that exact string is the registered Azure AD redirect URI. --- */

export const m365ConnectCallbackRouter = Router();

/**
 * helmet()'s default Cross-Origin-Opener-Policy is "same-origin" (app.ts). Since this callback
 * page is served from a different origin than the frontend that opened the popup
 * (localhost:4000 vs. localhost:5173, and in production the popup lands here after several
 * cross-origin hops through login.microsoftonline.com), that default severs `window.opener` the
 * moment this page loads — silently breaking both the postMessage handshake back to the opener
 * AND this page's own permission to call `window.close()` on itself (browsers tie "can this
 * script close its own window" to the browsing-context-group lineage that COOP just cut). Neither
 * failure throws or logs anything, which is why it looked like "it connected but the popup just
 * sits there and Manage Clouds never updates." Override to unsafe-none for this one route only.
 */
m365ConnectCallbackRouter.use((_req, res, next) => {
  res.setHeader("Cross-Origin-Opener-Policy", "unsafe-none");
  // helmet()'s default CSP (app.ts) is script-src 'self', which blocks inline <script> tags
  // outright with no exception — silently, no visible error, which is exactly why the postMessage
  // + window.close() script below never ran on any earlier attempt regardless of what the script
  // itself contained. This page is fully self-contained (no images/fonts/external resources), so
  // replace the default with a minimal CSP scoped to just this response, using a fresh per-request
  // nonce to allow only this one inline script to run.
  res.locals.cspNonce = randomBytes(16).toString("base64");
  res.setHeader(
    "Content-Security-Policy",
    `default-src 'none'; script-src 'nonce-${res.locals.cspNonce}'; style-src 'unsafe-inline'; base-uri 'none'; frame-ancestors 'self'`
  );
  next();
});

const CLOUD_TYPE_LABELS: Record<CloudType, string> = {
  onedrive: "OneDrive for Business",
  sharepoint: "SharePoint Online",
  teams: "Microsoft Teams",
  outlook: "Outlook",
};

/**
 * By the time this renders, the token exchange has already finished server-side — there's no
 * "in progress" state left to show (unlike the reference product's client-driven exchange). This
 * shows a brief branded success/error confirmation instead of an instant silent close, then
 * posts the result to the opener and closes itself.
 */
function popupResultPage(opts: { payload: unknown; ok: boolean; cloudType: CloudType | null; reason?: string; nonce: string }): string {
  const label = opts.cloudType ? CLOUD_TYPE_LABELS[opts.cloudType] : "your cloud";
  const message = opts.ok
    ? `Your ${label} account has been connected!`
    : `We couldn't connect ${label}${opts.reason ? ` — ${opts.reason}` : ""}.`;
  const iconColor = opts.ok ? "#1b2fc4" : "#dc2626";
  // The HTML parser scans for "</script" (even inside what's meant to be a JS string literal)
  // before the JS parser ever runs — so a payload field containing that literal substring (e.g. an
  // attacker-crafted ?error= query param reflected into payload.reason) could otherwise break out
  // of this script tag. Escaping "<" defeats that regardless of where it appears in the JSON.
  const payloadJson = JSON.stringify(opts.payload).replace(/</g, "\\u003c");

  return `<!doctype html><html><head><meta charset="utf-8"><title>CloudFuze</title></head>
<body style="margin:0;font-family:-apple-system,Segoe UI,Arial,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;background:#fff;">
  <div style="text-align:center;max-width:320px;padding:24px;">
    <div style="font-weight:700;font-size:20px;color:#1b2fc4;margin-bottom:20px;">CloudFuze</div>
    <div style="font-weight:600;font-size:16px;color:#1e293b;margin-bottom:20px;">${message}</div>
    <div style="width:48px;height:48px;margin:0 auto;border-radius:50%;border:3px solid ${iconColor};display:flex;align-items:center;justify-content:center;color:${iconColor};font-size:24px;">
      ${opts.ok ? "&#10003;" : "&#33;"}
    </div>
  </div>
  <script nonce="${opts.nonce}">
    // Target "*": this page (served from the backend's own origin) has no reliable way to know
    // the opener's real origin cross-origin (it can differ from ours, e.g. localhost:4000 here vs.
    // the frontend's localhost:5173 in dev) — window.location.origin would evaluate to THIS page's
    // origin, not the opener's, causing postMessage to silently drop the message on delivery.
    // The frontend validates the message came from this exact popup via event.source instead.
    window.opener && window.opener.postMessage(${payloadJson}, "*");
    setTimeout(function () { window.close(); }, 1200);
  </script>
</body></html>`;
}

/**
 * Success-path popup: the reference product's OAuth popup shows a spinner ("Your X authentication
 * is in progress....") for the full duration of enumeration, not just the OAuth exchange itself —
 * closing only once the sync job has actually finished, so Manage Clouds is already at its settled
 * state (not a live-ticking 0%) the moment the popup goes away. Our token exchange already
 * completed by the time this renders, so this polls the just-created connection's own sync job
 * status (same-origin, same browser — the operator's session cookie is sent automatically) instead
 * of re-deriving progress some other way.
 */
function popupProgressPage(opts: { connectionId: string; cloudType: CloudType; payload: unknown; nonce: string }): string {
  const label = CLOUD_TYPE_LABELS[opts.cloudType];
  const payloadJson = JSON.stringify(opts.payload).replace(/</g, "\\u003c");
  const connectionIdJson = JSON.stringify(opts.connectionId).replace(/</g, "\\u003c");

  return `<!doctype html><html><head><meta charset="utf-8"><title>CloudFuze</title></head>
<body style="margin:0;font-family:-apple-system,Segoe UI,Arial,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;background:#fff;">
  <div style="text-align:center;max-width:320px;padding:24px;">
    <div style="font-weight:700;font-size:22px;color:#1b2fc4;margin-bottom:20px;">CloudFuze</div>
    <div id="msg" style="font-weight:600;font-size:16px;color:#1e293b;margin-bottom:28px;">Your ${label} authentication is in progress....</div>
    <div id="spinner" style="width:56px;height:56px;margin:0 auto;border-radius:50%;border:5px solid #dbeafe;border-top-color:#1b2fc4;animation:spin 0.9s linear infinite;"></div>
    <div id="done" style="display:none;width:48px;height:48px;margin:0 auto;border-radius:50%;border:3px solid #1b2fc4;align-items:center;justify-content:center;color:#1b2fc4;font-size:24px;">&#10003;</div>
  </div>
  <style nonce="${opts.nonce}">@keyframes spin { to { transform: rotate(360deg); } }</style>
  <script nonce="${opts.nonce}">
    var connectionId = ${connectionIdJson};
    var TERMINAL = ["completed", "completed_with_errors", "failed", "cancelled"];
    var MAX_WAIT_MS = 90000; // safety net for a very large tenant — don't spin forever
    var startedAt = Date.now();

    function finish() {
      document.getElementById("spinner").style.display = "none";
      document.getElementById("done").style.display = "flex";
      document.getElementById("msg").textContent = "Your ${label} account has been connected!";
      window.opener && window.opener.postMessage(${payloadJson}, "*");
      setTimeout(function () { window.close(); }, 900);
    }

    function poll() {
      if (Date.now() - startedAt > MAX_WAIT_MS) { finish(); return; }
      fetch("/api/clouds/" + connectionId + "/status", { credentials: "include" })
        .then(function (r) { return r.json(); })
        .then(function (data) {
          var jobDone = !data.job || TERMINAL.indexOf(data.job.status) !== -1;
          var connDone = data.connectionStatus === "needs_reauth" || data.connectionStatus === "error";
          if (jobDone || connDone) { finish(); } else { setTimeout(poll, 1500); }
        })
        .catch(function () { setTimeout(poll, 1500); });
    }

    poll();
  </script>
</body></html>`;
}

m365ConnectCallbackRouter.get(
  "/",
  asyncHandler(async (req, res) => {
    const stateParam = typeof req.query.state === "string" ? req.query.state : "";

    let claims: Awaited<ReturnType<typeof consumeConnectAttempt>>["claims"];
    let codeVerifier: string;
    try {
      ({ claims, codeVerifier } = await consumeConnectAttempt(stateParam));
    } catch (err) {
      const reason = err instanceof InvalidOAuthStateError ? err.message : "invalid_state";
      res.type("html").send(
        popupResultPage({
          payload: { type: "m365-connect-complete", status: "error", cloudType: null, reason },
          ok: false,
          cloudType: null,
          reason: "your session expired, please try again",
          nonce: res.locals.cspNonce,
        })
      );
      return;
    }

    // Admin declined consent, or Microsoft reported an error — close the popup cleanly.
    if (req.query.error || typeof req.query.code !== "string") {
      const reason = typeof req.query.error === "string" ? req.query.error : "cancelled";
      res.type("html").send(
        popupResultPage({
          payload: { type: "m365-connect-complete", status: "error", cloudType: claims.cloudType, reason },
          ok: false,
          cloudType: claims.cloudType,
          reason: "sign-in was cancelled",
          nonce: res.locals.cspNonce,
        })
      );
      return;
    }

    try {
      const identity = await exchangeM365ConnectCode(req.query.code, codeVerifier);

      const tenantResult = await query<{ id: string }>(
        `INSERT INTO tenants (m365_tenant_id, display_name, status, connected_at)
         VALUES ($1, $2, 'connected', now())
         ON CONFLICT (m365_tenant_id) DO UPDATE SET status = 'connected', connected_at = now()
         RETURNING id`,
        [identity.m365TenantId, identity.tenantDomain]
      );
      const tenantId = tenantResult.rows[0]!.id;

      const connResult = await query<{ id: string }>(
        `INSERT INTO connections (tenant_id, cloud_type, admin_upn, admin_display_name, display_name, status, connected_at, encrypted_refresh_token, token_expiry, connected_by_operator_id)
         VALUES ($1, $2, $3, $4, $5, 'connecting', now(), $6, $7, $8)
         ON CONFLICT (tenant_id, cloud_type) DO UPDATE SET
           admin_upn = EXCLUDED.admin_upn,
           admin_display_name = EXCLUDED.admin_display_name,
           status = 'connecting',
           connected_at = now(),
           disconnected_at = NULL,
           encrypted_refresh_token = EXCLUDED.encrypted_refresh_token,
           token_expiry = EXCLUDED.token_expiry,
           connected_by_operator_id = EXCLUDED.connected_by_operator_id,
           last_error = NULL
         RETURNING id`,
        [
          tenantId,
          claims.cloudType,
          identity.adminUpn,
          identity.adminDisplayName,
          identity.tenantDomain,
          encryptToken(identity.refreshToken),
          identity.accessTokenExpiresOn,
          claims.operatorId,
        ]
      );
      const connectionId = connResult.rows[0]!.id;

      // A running backend process may already hold a cached application-permission token for this
      // tenant (graph/client.ts's tokenCache, keyed only by m365TenantId — shared across every
      // cloud type for the tenant, not just the one being (re)connected here) from before this
      // consent screen ran. Without this, a reconnect that adds a new permission (e.g. Mail.ReadWrite
      // added after a tenant's first consent) would silently keep using the stale pre-consent token
      // until the ~55-minute cache TTL expires — reconnecting would appear to succeed while sync
      // still failed with the exact same permission error as before.
      invalidateTenantTokenCache(identity.m365TenantId);

      // Every /api/clouds/* route (including GET /manage) scopes access through tenant_roles —
      // without this, the tenant this operator just connected would be invisible to them too,
      // since that join would simply exclude it. requireInternalAdmin already gated who could
      // reach this flow, so cleanup_admin here reflects that they're not merely a viewer.
      // ON CONFLICT DO NOTHING: don't downgrade a role a different grant may have already set.
      await query(
        `INSERT INTO tenant_roles (tenant_id, operator_id, role, granted_by)
         VALUES ($1, $2, 'cleanup_admin', $2)
         ON CONFLICT (tenant_id, operator_id) DO NOTHING`,
        [tenantId, claims.operatorId]
      );

      await query(
        `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail) VALUES ($1, $2, 'token_exchange', $3, $4)`,
        [connectionId, tenantId, claims.operatorId, { adminUpn: identity.adminUpn }]
      );
      await query(
        `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id) VALUES ($1, $2, 'connected', $3)`,
        [connectionId, tenantId, claims.operatorId]
      );

      const jobInsert = await query<{ id: string }>(
        `INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`,
        [connectionId]
      );
      await enqueueCloudSyncJob({ syncJobId: jobInsert.rows[0]!.id });

      res.type("html").send(
        popupProgressPage({
          connectionId,
          cloudType: claims.cloudType,
          payload: { type: "m365-connect-complete", status: "success", connectionId, cloudType: claims.cloudType },
          nonce: res.locals.cspNonce,
        })
      );
    } catch (err) {
      console.error("M365 connect callback failed", err);
      res.type("html").send(
        popupResultPage({
          payload: { type: "m365-connect-complete", status: "error", cloudType: claims.cloudType, reason: "exchange_failed" },
          ok: false,
          cloudType: claims.cloudType,
          reason: "something went wrong, please try again",
          nonce: res.locals.cspNonce,
        })
      );
    }
  })
);
