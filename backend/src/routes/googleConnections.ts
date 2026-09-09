import { Router } from "express";
import { z } from "zod";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireInternalAdmin } from "../middleware/internalAdmin.js";
import { enqueueCloudSyncJob } from "../jobs/queue.js";
import { GoogleDelegationError, verifyDomainDelegation } from "../services/googleWorkspaceAuth.js";
import { ApiError } from "../types/index.js";

/**
 * Google Workspace connect flow — deliberately NOT modeled on cloudConnections.ts's M365
 * connect/init -> popup -> OAuth consent -> callback flow, because there is no redirect to derive
 * identity from: Google Workspace domain-wide delegation is granted out-of-band, by the customer's
 * own Workspace super-admin, in THEIR Admin Console (Security -> API controls -> Domain-wide
 * delegation), authorizing this app's service account Client ID with the exact scopes it needs.
 * This route's job is to verify that grant actually took effect and to establish the
 * tenant/connection from what Google's own API reports back — never from what the operator typed.
 * See docs/google-workspace-integration.md.
 */
export const googleConnectionsRouter = Router();
googleConnectionsRouter.use(requireSession);

const verifyBodySchema = z.object({
  domain: z.string().trim().min(1).max(253),
  adminEmail: z.string().trim().email(),
});

/**
 * POST /api/google-clouds/connect/verify — the entire Google "Connect" action. Synchronous:
 * returns as soon as delegation is confirmed and the initial sync job is queued, no
 * authorizeUrl/popup/state involved (contrast cloudConnectionsRouter's /:cloudType/connect/init).
 * The frontend then polls the existing, already-generic GET /api/clouds/:id/status, same as every
 * other connection's post-connect progress.
 */
googleConnectionsRouter.post(
  "/connect/verify",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const body = verifyBodySchema.parse(req.body);

    let identity;
    try {
      identity = await verifyDomainDelegation(body.domain, body.adminEmail);
    } catch (err) {
      if (err instanceof GoogleDelegationError) {
        throw new ApiError(403, "GOOGLE_DELEGATION_NOT_GRANTED", err.message);
      }
      throw err;
    }

    const tenantResult = await query<{ id: string }>(
      `INSERT INTO tenants (google_customer_id, display_name, status, connected_at, connected_by_admin_upn)
       VALUES ($1, $2, 'connected', now(), $3)
       ON CONFLICT (google_customer_id) DO UPDATE SET status = 'connected', connected_at = now(), connected_by_admin_upn = EXCLUDED.connected_by_admin_upn
       RETURNING id`,
      [identity.customerId, identity.primaryDomain, body.adminEmail]
    );
    const tenantId = tenantResult.rows[0]!.id;

    const connResult = await query<{ id: string }>(
      `INSERT INTO connections (tenant_id, cloud_type, admin_upn, admin_display_name, display_name, status, connected_at, connected_by_operator_id)
       VALUES ($1, 'google_my_drive', $2, $3, $4, 'connecting', now(), $5)
       ON CONFLICT (tenant_id, cloud_type) DO UPDATE SET
         admin_upn = EXCLUDED.admin_upn,
         admin_display_name = EXCLUDED.admin_display_name,
         status = 'connecting',
         connected_at = now(),
         disconnected_at = NULL,
         connected_by_operator_id = EXCLUDED.connected_by_operator_id,
         last_error = NULL
       RETURNING id`,
      [tenantId, body.adminEmail, identity.adminDisplayName, identity.primaryDomain, req.session!.operatorId]
    );
    const connectionId = connResult.rows[0]!.id;

    // Same "an operator who just connected must immediately be able to see it" reasoning as the
    // M365 callback — cleanup_admin here reflects that requireInternalAdmin already gated who
    // could reach this route, not merely a viewer.
    await query(
      `INSERT INTO tenant_roles (tenant_id, operator_id, role, granted_by)
       VALUES ($1, $2, 'cleanup_admin', $2)
       ON CONFLICT (tenant_id, operator_id) DO NOTHING`,
      [tenantId, req.session!.operatorId]
    );

    await query(
      `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail) VALUES ($1, $2, 'token_exchange', $3, $4)`,
      [connectionId, tenantId, req.session!.operatorId, { adminUpn: body.adminEmail, domain: identity.primaryDomain }]
    );

    const jobInsert = await query<{ id: string }>(`INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`, [connectionId]);
    await query(
      `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail) VALUES ($1, $2, 'connected', $3, $4)`,
      [connectionId, tenantId, req.session!.operatorId, { syncJobId: jobInsert.rows[0]!.id }]
    );
    await enqueueCloudSyncJob({ syncJobId: jobInsert.rows[0]!.id });

    res.json({ connectionId, status: "connecting" });
  })
);
