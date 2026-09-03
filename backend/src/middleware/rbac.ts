import type { NextFunction, Request, Response } from "express";
import { query } from "../db/pool.js";
import { ApiError } from "../types/index.js";
import type { OperatorRole } from "../types/index.js";

/**
 * Enforces per-tenant role, server-side — see docs/api-spec.md "Roles" table and
 * docs/rollback-safety.md point 5. `viewer` can reach preview/reports; only `cleanup_admin` can
 * confirm, execute, cancel, or disconnect. Role is per (tenant, operator), not global.
 */
export function requireTenantRole(minRole: OperatorRole) {
  return async (req: Request, _res: Response, next: NextFunction) => {
    try {
      const tenantId = req.params.tenantId;
      const operatorId = req.session?.operatorId;
      if (!tenantId || !operatorId) {
        throw new ApiError(400, "MISSING_TENANT_CONTEXT", "tenantId and session are required");
      }

      const result = await query<{ role: OperatorRole }>(
        "SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2",
        [tenantId, operatorId]
      );
      const role = result.rows[0]?.role;

      if (!role || (minRole === "cleanup_admin" && role !== "cleanup_admin")) {
        throw new ApiError(403, "FORBIDDEN", `Requires ${minRole} role on this tenant`);
      }
      next();
    } catch (err) {
      next(err);
    }
  };
}
