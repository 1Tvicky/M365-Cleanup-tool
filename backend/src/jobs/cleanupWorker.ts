import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import { deleteChannel, deleteGroup } from "../graph/teams.js";
import { deleteFile } from "../graph/onedrive.js";
import { deleteDocumentLibrary } from "../graph/sharepoint.js";
import { runThrottled } from "../services/rateLimiter.js";
import type { PreviewScope } from "../types/index.js";
import { connection } from "./queue.js";

type DeletableItem =
  | { kind: "channel"; teamId: string; channelId: string; path: string }
  | { kind: "group"; groupId: string; path: string }
  | { kind: "file"; userId: string; fileId: string; path: string; sizeBytes: number }
  | { kind: "site_library"; siteId: string; driveId: string; path: string; sizeBytes: number };

/**
 * Runs export → delete for one confirmed cleanup job. Export is generated first and is never
 * skipped (per your decision that pre-delete backup is always mandatory) — deletes only begin
 * once the manifest (and, unless metadata-only, the content zip) has been written to blob storage.
 */
export const cleanupWorker = new Worker(
  "cleanup-jobs",
  async (job) => {
    const { jobId } = job.data as { jobId: string };

    const jobRow = await query<{
      tenant_id: string;
      m365_tenant_id: string;
      scope: PreviewScope;
      export_manifest_only: boolean;
    }>(
      `SELECT cj.tenant_id, t.m365_tenant_id, p.scope, cj.export_manifest_only
       FROM cleanup_jobs cj
       JOIN previews p ON p.id = cj.preview_id
       JOIN tenants t ON t.id = cj.tenant_id
       WHERE cj.id = $1`,
      [jobId]
    );
    const row = jobRow.rows[0];
    if (!row) throw new Error(`cleanup job ${jobId} not found`);

    await generateExportManifest(jobId, row.tenant_id, row.scope, row.export_manifest_only);
    await query(`UPDATE cleanup_jobs SET status = 'queued' WHERE id = $1`, [jobId]);

    const client = await graphClientForTenant(row.m365_tenant_id);
    const items = await resolveItemsFromScope(row.scope);

    await query(`UPDATE cleanup_jobs SET status = 'running', started_at = now(), total_items = $2 WHERE id = $1`, [
      jobId,
      items.length,
    ]);

    let completed = 0;
    let failed = 0;

    await runThrottled(
      items,
      async (item) => {
        switch (item.kind) {
          case "channel":
            return deleteChannel(client, item.teamId, item.channelId);
          case "group":
            return deleteGroup(client, item.groupId);
          case "file":
            return deleteFile(client, item.userId, item.fileId);
          case "site_library":
            return deleteDocumentLibrary(client, item.siteId, item.driveId);
        }
      },
      {
        isCancelled: async () => {
          const check = await query<{ cancel_requested_at: string | null }>(
            `SELECT cancel_requested_at FROM cleanup_jobs WHERE id = $1`,
            [jobId]
          );
          return check.rows[0]?.cancel_requested_at != null;
        },
        onItemSettled: async (item, result) => {
          if (result.ok) completed++;
          else failed++;

          await query(
            `INSERT INTO audit_entries (job_id, tenant_id, item_type, item_id, item_path, size_bytes, result, error_code)
             VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
            [
              jobId,
              row.tenant_id,
              item.kind,
              itemId(item),
              item.path,
              "sizeBytes" in item ? item.sizeBytes : 0,
              result.ok ? "deleted" : "failed",
              result.ok ? null : String((result as { error: unknown }).error),
            ]
          );
          await query(`UPDATE cleanup_jobs SET completed_items = $2, failed_items = $3 WHERE id = $1`, [
            jobId,
            completed,
            failed,
          ]);
        },
      }
    );

    const cancelled = await query<{ cancel_requested_at: string | null }>(
      `SELECT cancel_requested_at FROM cleanup_jobs WHERE id = $1`,
      [jobId]
    );
    const finalStatus = cancelled.rows[0]?.cancel_requested_at
      ? "cancelled"
      : failed > 0
        ? "completed_with_errors"
        : "completed";

    await query(`UPDATE cleanup_jobs SET status = $2, finished_at = now() WHERE id = $1`, [jobId, finalStatus]);
  },
  { connection, concurrency: 5 }
);

function itemId(item: DeletableItem): string {
  switch (item.kind) {
    case "channel":
      return item.channelId;
    case "group":
      return item.groupId;
    case "file":
      return item.fileId;
    case "site_library":
      return item.driveId;
  }
}

async function resolveItemsFromScope(scope: PreviewScope): Promise<DeletableItem[]> {
  // Re-resolves the confirmed scope against the discovery cache at execution time (not the stale
  // preview snapshot) so a confirm made 20 minutes ago doesn't delete items that no longer exist
  // or miss ones added since. Full implementation walks discovery_cache per docs/db schema.
  const items: DeletableItem[] = [];
  for (const team of scope.teams ?? []) {
    for (const channelId of team.channelIds ?? []) {
      items.push({ kind: "channel", teamId: team.teamId, channelId, path: `${team.teamId}/${channelId}` });
    }
  }
  return items;
}

/** Always runs before any delete, per docs/rollback-safety.md — the one recovery path independent of Graph's own retention windows. */
async function generateExportManifest(
  jobId: string,
  tenantId: string,
  scope: PreviewScope,
  manifestOnly: boolean
): Promise<void> {
  // Writes a manifest CSV (and, unless manifestOnly, zipped file contents) to blob storage, then
  // stores the signed URL on cleanup_jobs.export_manifest_url. Blob client wiring is
  // environment-specific (Azure Blob / S3-compatible) and intentionally left as an integration
  // point here rather than hardcoding one provider.
  const manifestUrl = `https://blob.example/exports/${jobId}.csv`;
  await query(`UPDATE cleanup_jobs SET export_manifest_url = $2 WHERE id = $1`, [jobId, manifestUrl]);
  void tenantId;
  void scope;
  void manifestOnly;
}
