import { Router } from "express";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import { listTeams } from "../graph/teams.js";
import { listSites } from "../graph/sharepoint.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireTenantRole } from "../middleware/rbac.js";
import { ApiError } from "../types/index.js";

export const discoveryRouter = Router();
discoveryRouter.use(requireSession);

async function m365TenantIdFor(tenantId: string): Promise<string> {
  const row = await query<{ m365_tenant_id: string }>(
    `SELECT m365_tenant_id FROM tenants WHERE id = $1 AND status = 'connected'`,
    [tenantId]
  );
  const tenant = row.rows[0];
  if (!tenant) throw new ApiError(404, "TENANT_NOT_FOUND", "No connected tenant with this id");
  return tenant.m365_tenant_id;
}

discoveryRouter.get(
  "/tenants/:tenantId/users",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const m365TenantId = await m365TenantIdFor(req.params.tenantId!);
    const client = await graphClientForTenant(m365TenantId);
    const search = typeof req.query.search === "string" ? req.query.search : "";
    const limit = Math.min(Number(req.query.limit) || 25, 200);

    let request = client.api("/users").top(limit).select("id,displayName,userPrincipalName");
    if (search) request = request.filter(`startswith(displayName,'${search.replace(/'/g, "''")}')`);
    const result = await request.get();

    res.json({ users: (result.value as any[]).map((u) => ({ id: u.id, displayName: u.displayName, upn: u.userPrincipalName })) });
  })
);

discoveryRouter.get(
  "/tenants/:tenantId/teams",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const m365TenantId = await m365TenantIdFor(req.params.tenantId!);
    const client = await graphClientForTenant(m365TenantId);
    const search = typeof req.query.search === "string" ? req.query.search : "";
    const limit = Math.min(Number(req.query.limit) || 25, 200);

    const teams = await listTeams(client, search, limit);
    res.json({ teams });
  })
);

discoveryRouter.get(
  "/tenants/:tenantId/sites",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const m365TenantId = await m365TenantIdFor(req.params.tenantId!);
    const client = await graphClientForTenant(m365TenantId);
    const search = typeof req.query.search === "string" ? req.query.search : "";
    const limit = Math.min(Number(req.query.limit) || 25, 200);

    const sites = await listSites(client, search, limit);
    res.json({ sites });
  })
);
