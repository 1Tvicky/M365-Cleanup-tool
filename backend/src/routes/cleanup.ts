import { Router } from "express";
import { z } from "zod";
import { config } from "../config/index.js";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import { listChannels } from "../graph/teams.js";
import { listUserFiles } from "../graph/onedrive.js";
import { listDocumentLibraries } from "../graph/sharepoint.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { requireTenantRole } from "../middleware/rbac.js";
import { enqueueCleanupJob } from "../jobs/queue.js";
import { ApiError } from "../types/index.js";
import type { PreviewResult, PreviewScope } from "../types/index.js";

export const cleanupRouter = Router();
cleanupRouter.use(requireSession);

const previewRequestSchema = z.object({
  scope: z.object({
    users: z.array(z.string()).optional(),
    teams: z.array(z.object({ teamId: z.string(), channelIds: z.array(z.string()).optional() })).optional(),
    sites: z.array(z.object({ siteId: z.string(), libraryIds: z.array(z.string()).optional() })).optional(),
    removeM365Groups: z.boolean().optional(),
    fileFilter: z.object({ olderThanCutoff: z.string().optional() }).optional(),
  }),
});

/** Dry-run only — never calls a Graph delete endpoint. See docs/api-spec.md. */
cleanupRouter.post(
  "/tenants/:tenantId/cleanup/preview",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const { scope } = previewRequestSchema.parse(req.body);
    const tenantId = req.params.tenantId;

    const tenantRow = await query<{ m365_tenant_id: string }>(
      `SELECT m365_tenant_id FROM tenants WHERE id = $1 AND status = 'connected'`,
      [tenantId]
    );
    const tenant = tenantRow.rows[0];
    if (!tenant) throw new ApiError(404, "TENANT_NOT_FOUND", "No connected tenant with this id");

    const result = await computePreview(tenant.m365_tenant_id, scope);

    const insert = await query<{ id: string }>(
      `INSERT INTO previews (tenant_id, requested_by, scope, result, expires_at)
       VALUES ($1, $2, $3, $4, now() + ($5 || ' minutes')::interval)
       RETURNING id`,
      [tenantId, req.session!.operatorId, scope, result, config.previewTtlMinutes]
    );

    res.json({ ...result, previewId: insert.rows[0]!.id });
  })
);

cleanupRouter.get(
  "/tenants/:tenantId/cleanup/preview/:previewId",
  requireTenantRole("viewer"),
  asyncHandler(async (req, res) => {
    const row = await query<{ result: PreviewResult; expires_at: string }>(
      `SELECT result, expires_at FROM previews WHERE id = $1 AND tenant_id = $2`,
      [req.params.previewId, req.params.tenantId]
    );
    const preview = row.rows[0];
    if (!preview) throw new ApiError(404, "PREVIEW_NOT_FOUND", "Preview not found or expired");
    if (new Date(preview.expires_at).getTime() < Date.now()) {
      throw new ApiError(410, "PREVIEW_EXPIRED", "Preview expired, request a new one");
    }
    res.json(preview.result);
  })
);

const confirmRequestSchema = z.object({
  previewId: z.string().uuid(),
  typedConfirmation: z.string(),
  exportManifestOnly: z.boolean(),
});

/**
 * Requires the typed "DELETE" phrase (server-verified, not just a UI checkbox) and always starts
 * with a mandatory export before any delete is queued — see docs/rollback-safety.md points 2-3.
 */
cleanupRouter.post(
  "/tenants/:tenantId/cleanup/confirm",
  requireTenantRole("cleanup_admin"),
  asyncHandler(async (req, res) => {
    const body = confirmRequestSchema.parse(req.body);
    const tenantId = req.params.tenantId;

    if (body.typedConfirmation !== "DELETE") {
      throw new ApiError(400, "CONFIRMATION_MISMATCH", 'Typed confirmation must exactly equal "DELETE"');
    }

    const previewRow = await query<{ scope: PreviewScope; expires_at: string }>(
      `SELECT scope, expires_at FROM previews WHERE id = $1 AND tenant_id = $2`,
      [body.previewId, tenantId]
    );
    const preview = previewRow.rows[0];
    if (!preview) throw new ApiError(404, "PREVIEW_NOT_FOUND", "Preview not found or expired");
    if (new Date(preview.expires_at).getTime() < Date.now()) {
      throw new ApiError(410, "PREVIEW_EXPIRED", "Preview expired, request a new one");
    }

    const jobInsert = await query<{ id: string }>(
      `INSERT INTO cleanup_jobs (tenant_id, preview_id, confirmed_by, export_manifest_only)
       VALUES ($1, $2, $3, $4) RETURNING id`,
      [tenantId, body.previewId, req.session!.operatorId, body.exportManifestOnly]
    );
    const jobId = jobInsert.rows[0]!.id;

    const auditInsert = await query<{ id: string }>(
      `INSERT INTO job_audit_log (job_id, tenant_id, operator_id, scope_summary, typed_confirmation_matched)
       VALUES ($1, $2, $3, $4, true) RETURNING id`,
      [jobId, tenantId, req.session!.operatorId, preview.scope]
    );

    await enqueueCleanupJob({ jobId });

    res.status(202).json({ jobId, status: "export_in_progress", auditLogId: auditInsert.rows[0]!.id });
  })
);

async function computePreview(m365TenantId: string, scope: PreviewScope): Promise<PreviewResult> {
  const client = await graphClientForTenant(m365TenantId);
  const warnings: string[] = [];

  const teamsBreakdown = [];
  let teamsItemCount = 0;
  for (const t of scope.teams ?? []) {
    const channels = await listChannels(client, t.teamId);
    const toDelete = t.channelIds ? channels.filter((c) => t.channelIds!.includes(c.id)) : channels;
    if (toDelete.some((c) => c.membershipType === "private")) {
      warnings.push(`Team ${t.teamId} includes a private channel — deletion is not restorable via Graph.`);
    }
    teamsItemCount += toDelete.length;
    teamsBreakdown.push({
      teamId: t.teamId,
      displayName: t.teamId,
      channelsToDelete: toDelete.length,
      groupWillBeRemoved: Boolean(scope.removeM365Groups),
    });
  }

  const onedriveBreakdown = [];
  let onedriveItemCount = 0;
  let onedriveBytes = 0;
  for (const userId of scope.users ?? []) {
    const files = await listUserFiles(client, userId, scope.fileFilter?.olderThanCutoff);
    const size = files.reduce((sum, f) => sum + f.sizeBytes, 0);
    onedriveItemCount += files.length;
    onedriveBytes += size;
    onedriveBreakdown.push({ userId, upn: userId, fileCount: files.length, sizeBytes: size });
  }

  const sharepointBreakdown = [];
  let sharepointBytes = 0;
  for (const s of scope.sites ?? []) {
    const libraries = await listDocumentLibraries(client, s.siteId);
    const selected = s.libraryIds ? libraries.filter((l) => s.libraryIds!.includes(l.id)) : libraries;
    const size = selected.reduce((sum, l) => sum + l.sizeBytes, 0);
    sharepointBytes += size;
    sharepointBreakdown.push({
      siteId: s.siteId,
      displayName: s.siteId,
      libraryCount: selected.length,
      sizeBytes: size,
    });
  }

  return {
    previewId: "", // filled by caller once the preview row is inserted
    generatedAt: new Date().toISOString(),
    totals: {
      itemCount: teamsItemCount + onedriveItemCount + sharepointBreakdown.length,
      totalSizeBytes: onedriveBytes + sharepointBytes,
    },
    breakdown: {
      teams: teamsBreakdown,
      onedrive: onedriveBreakdown,
      sharepoint: sharepointBreakdown,
      chats: [], // report-only chat counts populated via listUserChatsForReport when users are in scope; omitted here for brevity
    },
    warnings,
  };
}
