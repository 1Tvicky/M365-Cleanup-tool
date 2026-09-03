import { Router } from "express";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireTenantRole } from "../middleware/rbac.js";
import { ApiError } from "../types/index.js";
import type { AuditEntry, JobStatus } from "../types/index.js";

export const jobsRouter = Router();
jobsRouter.use(requireSession);

jobsRouter.get(
  "/jobs/:jobId",
  asyncHandler(async (req, res) => {
    const row = await query<{
      id: string;
      tenant_id: string;
      status: JobStatus;
      started_at: string | null;
      finished_at: string | null;
      total_items: number;
      completed_items: number;
      failed_items: number;
      export_manifest_url: string | null;
    }>(
      `SELECT id, tenant_id, status, started_at, finished_at, total_items, completed_items, failed_items, export_manifest_url
       FROM cleanup_jobs WHERE id = $1`,
      [req.params.jobId]
    );
    const job = row.rows[0];
    if (!job) throw new ApiError(404, "JOB_NOT_FOUND", "No such job");

    res.json({
      jobId: job.id,
      tenantId: job.tenant_id,
      status: job.status,
      startedAt: job.started_at,
      finishedAt: job.finished_at,
      progress: { total: job.total_items, completed: job.completed_items, failed: job.failed_items },
      exportManifestUrl: job.export_manifest_url,
      currentThrottle: null, // live throttle state is pushed via /jobs/:jobId/stream (SSE), not polled here
    });
  })
);

/** Best-effort cancel — in-flight batch calls finish; already-deleted items stay deleted. See docs/rollback-safety.md point 6. */
jobsRouter.post(
  "/jobs/:jobId/cancel",
  asyncHandler(async (req, res) => {
    const jobRow = await query<{ tenant_id: string }>(`SELECT tenant_id FROM cleanup_jobs WHERE id = $1`, [
      req.params.jobId,
    ]);
    const job = jobRow.rows[0];
    if (!job) throw new ApiError(404, "JOB_NOT_FOUND", "No such job");

    // requireTenantRole reads tenantId from route params, not available here — re-check inline.
    const roleCheck = await query<{ role: string }>(
      `SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2`,
      [job.tenant_id, req.session!.operatorId]
    );
    if (roleCheck.rows[0]?.role !== "cleanup_admin") {
      throw new ApiError(403, "FORBIDDEN", "Requires cleanup_admin role on this tenant");
    }

    await query(`UPDATE cleanup_jobs SET cancel_requested_at = now() WHERE id = $1`, [req.params.jobId]);
    res.status(202).json({ jobId: req.params.jobId, status: "cancel_requested" });
  })
);

jobsRouter.get(
  "/jobs/:jobId/audit",
  asyncHandler(async (req, res) => {
    const rows = await query<AuditEntry & { size_bytes: number; error_code: string | null; occurred_at: string }>(
      `SELECT item_type AS "itemType", item_id AS "itemId", item_path AS "itemPath", size_bytes AS "sizeBytes",
              result, error_code AS "errorCode", occurred_at AS "timestamp"
       FROM audit_entries WHERE job_id = $1 ORDER BY occurred_at ASC LIMIT 5000`,
      [req.params.jobId]
    );
    res.json({ entries: rows.rows });
  })
);

jobsRouter.get(
  "/jobs/:jobId/audit/export",
  asyncHandler(async (req, res) => {
    const rows = await query<AuditEntry>(
      `SELECT item_type AS "itemType", item_id AS "itemId", item_path AS "itemPath", size_bytes AS "sizeBytes",
              result, error_code AS "errorCode", occurred_at AS "timestamp"
       FROM audit_entries WHERE job_id = $1 ORDER BY occurred_at ASC`,
      [req.params.jobId]
    );

    const header = "itemType,itemId,itemPath,sizeBytes,result,errorCode,timestamp";
    const lines = rows.rows.map((r) =>
      [r.itemType, r.itemId, csvEscape(r.itemPath), r.sizeBytes, r.result, r.errorCode ?? "", r.timestamp].join(",")
    );

    res.setHeader("Content-Type", "text/csv");
    res.setHeader("Content-Disposition", `attachment; filename="audit-${req.params.jobId}.csv"`);
    res.send([header, ...lines].join("\n"));
  })
);

jobsRouter.get(
  "/tenants/:tenantId/jobs",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const rows = await query(
      `SELECT id, status, total_items AS "totalItems", completed_items AS "completedItems",
              failed_items AS "failedItems", confirmed_at AS "confirmedAt", finished_at AS "finishedAt"
       FROM cleanup_jobs WHERE tenant_id = $1 ORDER BY confirmed_at DESC LIMIT 200`,
      [req.params.tenantId]
    );
    res.json({ jobs: rows.rows, nextCursor: null });
  })
);

jobsRouter.get(
  "/tenants/:tenantId/reports/summary",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const rows = await query<{ total_bytes: string; total_items: string; job_count: string; failed_items: string }>(
      `SELECT COALESCE(SUM(a.size_bytes) FILTER (WHERE a.result = 'deleted'), 0) AS total_bytes,
              COUNT(*) FILTER (WHERE a.result = 'deleted') AS total_items,
              COUNT(DISTINCT cj.id) AS job_count,
              COUNT(*) FILTER (WHERE a.result = 'failed') AS failed_items
       FROM cleanup_jobs cj
       LEFT JOIN audit_entries a ON a.job_id = cj.id
       WHERE cj.tenant_id = $1`,
      [req.params.tenantId]
    );
    const summary = rows.rows[0]!;
    res.json({
      totalBytesReclaimed: Number(summary.total_bytes),
      totalItemsDeleted: Number(summary.total_items),
      jobsRun: Number(summary.job_count),
      itemsFailed: Number(summary.failed_items),
    });
  })
);

function csvEscape(value: string): string {
  return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}
