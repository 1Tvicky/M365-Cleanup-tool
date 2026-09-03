import { Router } from "express";
import { query } from "../db/pool.js";
import { invalidateTenantTokenCache } from "../graph/client.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireTenantRole } from "../middleware/rbac.js";
import { ApiError } from "../types/index.js";
import type { TenantSummary } from "../types/index.js";

export const tenantsRouter = Router();
tenantsRouter.use(requireSession);

tenantsRouter.get(
  "/",
  asyncHandler(async (req, res) => {
    const operatorId = req.session!.operatorId;
    const result = await query<TenantSummary & { m365_tenant_id: string; connected_by_admin_upn: string; connected_at: string; last_token_refresh_at: string }>(
      `SELECT t.id, t.display_name AS "displayName", t.m365_tenant_id AS "m365TenantId", t.status,
              t.connected_at AS "connectedAt", t.connected_by_admin_upn AS "connectedByAdminUpn",
              t.last_token_refresh_at AS "lastTokenRefreshAt"
       FROM tenants t
       JOIN tenant_roles tr ON tr.tenant_id = t.id
       WHERE tr.operator_id = $1
       ORDER BY t.connected_at DESC NULLS LAST
       LIMIT 200`,
      [operatorId]
    );
    res.json({ tenants: result.rows.map((r) => ({ ...r, workloads: ["teams", "onedrive", "sharepoint"] })), nextCursor: null });
  })
);

tenantsRouter.get(
  "/:tenantId",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const result = await query(
      `SELECT id, display_name AS "displayName", m365_tenant_id AS "m365TenantId", status,
              connected_at AS "connectedAt", connected_by_admin_upn AS "connectedByAdminUpn",
              last_token_refresh_at AS "lastTokenRefreshAt", permissions_granted AS "permissionsGranted"
       FROM tenants WHERE id = $1`,
      [req.params.tenantId]
    );
    const tenant = result.rows[0];
    if (!tenant) throw new ApiError(404, "TENANT_NOT_FOUND", "No such tenant");

    const jobs = await query(
      `SELECT id, status, total_items AS "totalItems", completed_items AS "completedItems",
              failed_items AS "failedItems", confirmed_at AS "confirmedAt"
       FROM cleanup_jobs WHERE tenant_id = $1 ORDER BY confirmed_at DESC LIMIT 10`,
      [req.params.tenantId]
    );

    res.json({ ...tenant, recentJobs: jobs.rows });
  })
);

/**
 * Disconnect stops job scheduling and purges cached tokens; it cannot itself revoke the customer
 * tenant's Enterprise Application consent grant — only the customer's own admin can do that from
 * their Enterprise Applications page. See docs/azure-ad-app-registration.md §6.
 */
tenantsRouter.post(
  "/:tenantId/disconnect",
  requireTenantRole("cleanup_admin"),
  asyncHandler(async (req, res) => {
    const result = await query<{ m365_tenant_id: string }>(
      `UPDATE tenants SET status = 'disconnected', disconnected_at = now()
       WHERE id = $1 RETURNING m365_tenant_id`,
      [req.params.tenantId]
    );
    const tenant = result.rows[0];
    if (!tenant) throw new ApiError(404, "TENANT_NOT_FOUND", "No such tenant");

    invalidateTenantTokenCache(tenant.m365_tenant_id);
    res.status(204).end();
  })
);
