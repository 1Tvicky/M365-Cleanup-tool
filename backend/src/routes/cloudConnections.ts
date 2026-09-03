import { Router } from "express";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireInternalAdmin } from "../middleware/internalAdmin.js";
import { exchangeM365ConnectCode, getM365ConnectAuthorizeUrl } from "../services/m365ConnectAuth.js";
import { consumeConnectAttempt, InvalidOAuthStateError, startConnectAttempt } from "../services/oauthState.js";
import { encryptToken } from "../services/tokenEncryption.js";
import { enqueueCloudSyncJob } from "../jobs/queue.js";
import { ApiError } from "../types/index.js";
import { CLOUD_TYPES, isCloudType, type CloudType, type ConnectionUserRow, type ManageCloudsRow } from "../types/connections.js";

export const cloudConnectionsRouter = Router();
cloudConnectionsRouter.use(requireSession);

/**
 * Verifies the operator has tenant_roles access to the connection's tenant — same RBAC scoping
 * the rest of the app uses (middleware/rbac.ts), applied here since connections aren't addressed
 * by :tenantId in the URL the way cleanup routes are.
 */
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

/** GET /api/clouds/manage — one row per (tenant, cloud_type) connection the operator can see. */
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
      added_users: string | null;
      not_added_users: string | null;
    }>(
      `SELECT DISTINCT ON (c.id)
              c.id, c.cloud_type, c.display_name, c.admin_upn, c.admin_display_name, c.status,
              c.connected_at, c.last_synced_at, c.last_error,
              sj.total_users, sj.processed_users,
              cu.added_users, cu.not_added_users
       FROM connections c
       JOIN tenant_roles tr ON tr.tenant_id = c.tenant_id AND tr.operator_id = $1
       LEFT JOIN sync_jobs sj ON sj.connection_id = c.id
       LEFT JOIN (
         SELECT connection_id,
                COUNT(*) FILTER (WHERE sync_status = 'synced') AS added_users,
                COUNT(*) FILTER (WHERE sync_status = 'failed') AS not_added_users
         FROM connection_users
         GROUP BY connection_id
       ) cu ON cu.connection_id = c.id
       WHERE c.status != 'disconnected'
       ORDER BY c.id, sj.created_at DESC NULLS LAST`,
      [req.session!.operatorId]
    );

    const connections: ManageCloudsRow[] = result.rows.map((r) => {
      const total = r.total_users ?? 0;
      const processed = r.processed_users ?? 0;
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
        addedUsers: Number(r.added_users ?? 0),
        notAddedUsers: Number(r.not_added_users ?? 0),
        percent: total > 0 ? Math.round((processed / total) * 100) : 0,
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

/**
 * POST /api/clouds/:id/resync — creates a fresh sync_jobs row rather than reusing the last one, so
 * sync history isn't overwritten. connections.last_synced_at updates on completion, not here — see
 * docs/cloud-connections-api.md for why that deviates from a literal "update on resync" reading.
 */
cloudConnectionsRouter.post(
  "/:id/resync",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const { tenantId } = await requireConnectionAccess(req.params.id!, req.session!.operatorId);

    const connRow = await query<{ status: string }>(`SELECT status FROM connections WHERE id = $1`, [req.params.id]);
    const status = connRow.rows[0]?.status;
    if (status === "disconnected") {
      throw new ApiError(409, "CONNECTION_DISCONNECTED", "This connection is disconnected — reconnect instead of resyncing");
    }

    const running = await query(
      `SELECT 1 FROM sync_jobs WHERE connection_id = $1 AND status IN ('queued', 'running') LIMIT 1`,
      [req.params.id]
    );
    if (running.rows.length > 0) {
      throw new ApiError(409, "RESYNC_ALREADY_RUNNING", "A sync job is already in progress for this connection");
    }

    const jobInsert = await query<{ id: string }>(
      `INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`,
      [req.params.id]
    );
    const jobId = jobInsert.rows[0]!.id;

    await query(
      `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail)
       VALUES ($1, $2, 'resync_requested', $3, $4)`,
      [req.params.id, tenantId, req.session!.operatorId, { syncJobId: jobId }]
    );

    await enqueueCloudSyncJob({ syncJobId: jobId });
    res.status(202).json({ jobId, status: "queued" });
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

const CLOUD_TYPE_LABELS: Record<CloudType, string> = {
  onedrive: "OneDrive for Business",
  sharepoint: "SharePoint Online",
  teams: "Microsoft Teams",
};

/**
 * By the time this renders, the token exchange has already finished server-side — there's no
 * "in progress" state left to show (unlike the reference product's client-driven exchange). This
 * shows a brief branded success/error confirmation instead of an instant silent close, then
 * posts the result to the opener and closes itself.
 */
function popupResultPage(opts: { payload: unknown; ok: boolean; cloudType: CloudType | null; reason?: string }): string {
  const label = opts.cloudType ? CLOUD_TYPE_LABELS[opts.cloudType] : "your cloud";
  const message = opts.ok
    ? `Your ${label} account has been connected!`
    : `We couldn't connect ${label}${opts.reason ? ` — ${opts.reason}` : ""}.`;
  const iconColor = opts.ok ? "#1b2fc4" : "#dc2626";

  return `<!doctype html><html><head><meta charset="utf-8"><title>CloudFuze</title></head>
<body style="margin:0;font-family:-apple-system,Segoe UI,Arial,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;background:#fff;">
  <div style="text-align:center;max-width:320px;padding:24px;">
    <div style="font-weight:700;font-size:20px;color:#1b2fc4;margin-bottom:20px;">CloudFuze</div>
    <div style="font-weight:600;font-size:16px;color:#1e293b;margin-bottom:20px;">${message}</div>
    <div style="width:48px;height:48px;margin:0 auto;border-radius:50%;border:3px solid ${iconColor};display:flex;align-items:center;justify-content:center;color:${iconColor};font-size:24px;">
      ${opts.ok ? "&#10003;" : "&#33;"}
    </div>
  </div>
  <script>
    window.opener && window.opener.postMessage(${JSON.stringify(opts.payload)}, window.location.origin);
    setTimeout(function () { window.close(); }, 1200);
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
        popupResultPage({
          payload: { type: "m365-connect-complete", status: "success", connectionId, cloudType: claims.cloudType },
          ok: true,
          cloudType: claims.cloudType,
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
        })
      );
    }
  })
);
