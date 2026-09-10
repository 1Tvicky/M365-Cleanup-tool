import { Router } from "express";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireInternalAdmin } from "../middleware/internalAdmin.js";
import { enqueueCloudSyncJob } from "../jobs/queue.js";
import { exchangeGoogleConnectCode, getGoogleConnectAuthorizeUrl } from "../services/googleConnectAuth.js";
import { GoogleDelegationError, verifyDomainDelegation } from "../services/googleWorkspaceAuth.js";
import { consumeConnectAttempt, InvalidOAuthStateError, startConnectAttempt } from "../services/oauthState.js";
import { popupCallbackHeaders, popupProgressPage, popupResultPage } from "./cloudConnections.js";
import { GOOGLE_CLOUD_TYPES, isCloudType } from "../types/connections.js";
import { ApiError } from "../types/index.js";

/**
 * Google Workspace connect flow — a real OAuth popup, same shape and same reused HTML/state
 * helpers as cloudConnections.ts's M365 flow (POST /connect/init opens a popup,
 * googleConnectCallbackRouter is the fixed-path redirect target). The one structural difference:
 * this OAuth exchange only ever identifies the connecting admin (see
 * services/googleConnectAuth.ts) — it is NOT what grants data access. Data access is domain-wide
 * delegation via services/googleWorkspaceAuth.ts's service account, which the callback verifies
 * immediately after identity is confirmed, before creating the connection. If the customer's
 * Workspace super-admin hasn't authorized that service account yet, the popup shows a clear error
 * instead of creating a connection that can never actually sync. See
 * docs/google-workspace-integration.md.
 */
export const googleConnectionsRouter = Router();
googleConnectionsRouter.use(requireSession);

/** POST /api/google-clouds/:cloudType/connect/init — mirrors cloudConnectionsRouter's M365 route exactly. */
googleConnectionsRouter.post(
  "/:cloudType/connect/init",
  requireInternalAdmin,
  asyncHandler(async (req, res) => {
    const cloudType = req.params.cloudType;
    if (!cloudType || !isCloudType(cloudType) || !GOOGLE_CLOUD_TYPES.includes(cloudType)) {
      throw new ApiError(400, "INVALID_CLOUD_TYPE", `cloudType must be one of: ${GOOGLE_CLOUD_TYPES.join(", ")}`);
    }

    const attempt = await startConnectAttempt(cloudType, req.session!.operatorId);
    const authorizeUrl = getGoogleConnectAuthorizeUrl(attempt.authorizeParams.state);

    res.json({ authorizeUrl, state: attempt.state });
  })
);

/* --- OAuth callback: mounted separately in app.ts at a fixed top-level path, since that exact
   string must match this Google Cloud OAuth client's registered redirect URI. --- */

export const googleConnectCallbackRouter = Router();
googleConnectCallbackRouter.use(popupCallbackHeaders);

googleConnectCallbackRouter.get(
  "/",
  asyncHandler(async (req, res) => {
    const stateParam = typeof req.query.state === "string" ? req.query.state : "";

    let claims: Awaited<ReturnType<typeof consumeConnectAttempt>>["claims"];
    try {
      ({ claims } = await consumeConnectAttempt(stateParam));
    } catch (err) {
      const reason = err instanceof InvalidOAuthStateError ? err.message : "invalid_state";
      res.type("html").send(
        popupResultPage({
          payload: { type: "google-connect-complete", status: "error", cloudType: null, reason },
          ok: false,
          cloudType: null,
          reason: "your session expired, please try again",
          nonce: res.locals.cspNonce,
        })
      );
      return;
    }

    // Admin declined consent, or Google reported an error — close the popup cleanly.
    if (req.query.error || typeof req.query.code !== "string") {
      const reason = typeof req.query.error === "string" ? req.query.error : "cancelled";
      res.type("html").send(
        popupResultPage({
          payload: { type: "google-connect-complete", status: "error", cloudType: claims.cloudType, reason },
          ok: false,
          cloudType: claims.cloudType,
          reason: "sign-in was cancelled",
          nonce: res.locals.cspNonce,
        })
      );
      return;
    }

    try {
      const identity = await exchangeGoogleConnectCode(req.query.code);
      const domain = identity.adminEmail.split("@")[1] ?? "";

      // The OAuth step above only proves WHO is connecting — it grants no data access. Verify
      // domain-wide delegation actually works before creating anything: without this, a connection
      // could be created that can never sync, with no clear explanation why.
      let delegation;
      try {
        delegation = await verifyDomainDelegation(domain, identity.adminEmail);
      } catch (err) {
        const reason =
          err instanceof GoogleDelegationError
            ? err.message
            : "Google Workspace domain-wide delegation isn't set up for this domain yet.";
        res.type("html").send(
          popupResultPage({
            payload: { type: "google-connect-complete", status: "error", cloudType: claims.cloudType, reason },
            ok: false,
            cloudType: claims.cloudType,
            reason,
            nonce: res.locals.cspNonce,
          })
        );
        return;
      }

      const tenantResult = await query<{ id: string }>(
        `INSERT INTO tenants (google_customer_id, display_name, status, connected_at, connected_by_admin_upn)
         VALUES ($1, $2, 'connected', now(), $3)
         ON CONFLICT (google_customer_id) DO UPDATE SET status = 'connected', connected_at = now(), connected_by_admin_upn = EXCLUDED.connected_by_admin_upn
         RETURNING id`,
        [delegation.customerId, delegation.primaryDomain, identity.adminEmail]
      );
      const tenantId = tenantResult.rows[0]!.id;

      const connResult = await query<{ id: string }>(
        `INSERT INTO connections (tenant_id, cloud_type, admin_upn, admin_display_name, display_name, status, connected_at, connected_by_operator_id)
         VALUES ($1, $2, $3, $4, $5, 'connecting', now(), $6)
         ON CONFLICT (tenant_id, cloud_type) DO UPDATE SET
           admin_upn = EXCLUDED.admin_upn,
           admin_display_name = EXCLUDED.admin_display_name,
           status = 'connecting',
           connected_at = now(),
           disconnected_at = NULL,
           connected_by_operator_id = EXCLUDED.connected_by_operator_id,
           last_error = NULL
         RETURNING id`,
        [tenantId, claims.cloudType, identity.adminEmail, identity.adminDisplayName, delegation.primaryDomain, claims.operatorId]
      );
      const connectionId = connResult.rows[0]!.id;

      await query(
        `INSERT INTO tenant_roles (tenant_id, operator_id, role, granted_by)
         VALUES ($1, $2, 'cleanup_admin', $2)
         ON CONFLICT (tenant_id, operator_id) DO NOTHING`,
        [tenantId, claims.operatorId]
      );

      await query(
        `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id, detail) VALUES ($1, $2, 'token_exchange', $3, $4)`,
        [connectionId, tenantId, claims.operatorId, { adminUpn: identity.adminEmail }]
      );
      await query(
        `INSERT INTO connection_events (connection_id, tenant_id, event, operator_id) VALUES ($1, $2, 'connected', $3)`,
        [connectionId, tenantId, claims.operatorId]
      );

      const jobInsert = await query<{ id: string }>(`INSERT INTO sync_jobs (connection_id, status) VALUES ($1, 'queued') RETURNING id`, [connectionId]);
      await enqueueCloudSyncJob({ syncJobId: jobInsert.rows[0]!.id });

      res.type("html").send(
        popupProgressPage({
          connectionId,
          cloudType: claims.cloudType,
          payload: { type: "google-connect-complete", status: "success", connectionId, cloudType: claims.cloudType },
          nonce: res.locals.cspNonce,
        })
      );
    } catch (err) {
      console.error("Google connect callback failed", err);
      res.type("html").send(
        popupResultPage({
          payload: { type: "google-connect-complete", status: "error", cloudType: claims.cloudType, reason: "exchange_failed" },
          ok: false,
          cloudType: claims.cloudType,
          reason: "something went wrong, please try again",
          nonce: res.locals.cspNonce,
        })
      );
    }
  })
);
