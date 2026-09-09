import { query } from "../db/pool.js";
import { getDriveClientAs } from "../services/googleWorkspaceAuth.js";
import { classifyGoogleDeleteError, deleteMyDriveItem } from "../graph/googleDriveDeletion.js";
import { listRootDriveItems } from "../graph/googleDriveEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import type { PendingItem } from "./cleanupExecutionWorker.js";

/**
 * Google My Drive cleanup execution — the Google equivalent of cleanupExecutionWorker.ts's
 * executeItem. Deletes every top-level file/folder in the impersonated user's My Drive; never
 * touches the user, the Workspace account, or the domain. Builds its own per-item impersonated
 * Drive client (item.graph_ref.userEmail) rather than sharing one client for the whole job, unlike
 * every M365 execute* function — Google's domain-wide delegation requires impersonating a specific
 * user per call.
 */
export async function executeGoogleMyDriveItem(item: PendingItem, operationId: string, permanent: boolean): Promise<void> {
  const userEmail = item.graph_ref.userEmail!;
  const drive = await getDriveClientAs(userEmail);

  const children = await listRootDriveItems(drive);
  if (children.length === 0) return; // already empty — nothing to do, counts as success

  // Seed one row per file up front — same convention as executeItem, so the report/live
  // "recently removed" feed reflects the full file list from the start.
  for (const child of children) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, child.name, child.id, child.sizeBytes]
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, children.length]);

  let firstError: unknown = null;
  await runThrottled(children, (child) => deleteMyDriveItem(drive, child.id, permanent), {
    label: "GoogleMyDrive",
    batchSize: 10,
    onItemSettled: async (child, result) => {
      const fileStatus = result.ok ? result.value : "failed"; // "deleted" | "already_gone" | "failed"
      const errorMessage = result.ok ? null : classifyGoogleDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, child.id, fileStatus, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}

/** 401/403s that specifically indicate domain-wide delegation was revoked, not a transient per-file failure. Same shape as jobs/googleDriveSync.ts's isGoogleReauthError — duplicated rather than shared since sync and cleanup execution are deliberately separate modules (see docs/google-workspace-integration.md). */
export function isGoogleReauthCleanupError(err: unknown): boolean {
  const code = (err as { code?: number })?.code;
  const reason = (err as { errors?: { reason?: string }[] })?.errors?.[0]?.reason ?? "";
  return code === 401 || code === 403 || /unauthorized_client|invalid_grant|forbidden|domainPolicy/i.test(reason);
}
