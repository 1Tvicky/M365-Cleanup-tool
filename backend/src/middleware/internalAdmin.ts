import type { NextFunction, Request, Response } from "express";
import { query } from "../db/pool.js";
import { ApiError } from "../types/index.js";

/**
 * Gates connect/resync/disconnect on `operators.is_internal_admin` — a global flag, deliberately
 * separate from the per-tenant `tenant_roles.role` used by the Cleanup module (middleware/rbac.ts).
 * Connecting or disconnecting a customer tenant is a broader trust decision (it grants or revokes
 * this app's standing access) than running a cleanup job inside a tenant that's already
 * connected — see docs/cloud-connections-api.md "Roles". Must run after requireSession.
 */
export async function requireInternalAdmin(req: Request, _res: Response, next: NextFunction): Promise<void> {
  try {
    const operatorId = req.session?.operatorId;
    if (!operatorId) {
      throw new ApiError(401, "UNAUTHENTICATED", "Missing session");
    }

    const result = await query<{ is_internal_admin: boolean }>(
      `SELECT is_internal_admin FROM operators WHERE id = $1`,
      [operatorId]
    );
    if (!result.rows[0]?.is_internal_admin) {
      throw new ApiError(403, "FORBIDDEN", "Requires internal admin access to connect or disconnect a cloud");
    }
    next();
  } catch (err) {
    next(err);
  }
}
