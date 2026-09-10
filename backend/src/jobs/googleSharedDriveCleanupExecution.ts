import { query } from "../db/pool.js";
import { getDriveClientAs } from "../services/googleWorkspaceAuth.js";
import { classifySharedDriveDeleteError, deleteSharedDriveItem } from "../graph/googleSharedDriveDeletion.js";
import { listSharedDriveRootItems } from "../graph/googleSharedDriveEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import type { PendingItem } from "./cleanupExecutionWorker.js";

/**
 * Google Shared Drive cleanup execution — mirrors jobs/googleDriveCleanupExecution.ts, but
 * impersonates the tenant's admin (passed in, not read from item.graph_ref) since a Shared Drive
 * has no owning user. Deletes only the drive's top-level content, never the drive itself.
 */
export async function executeSharedDriveItem(adminUpn: string, item: PendingItem, operationId: string, permanent: boolean): Promise<void> {
  const driveId = item.graph_ref.driveId!;
  const drive = await getDriveClientAs(adminUpn);

  const children = await listSharedDriveRootItems(drive, driveId);
  if (children.length === 0) return;

  for (const child of children) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, child.name, child.id, child.sizeBytes]
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, children.length]);

  let firstError: unknown = null;
  await runThrottled(children, (child) => deleteSharedDriveItem(drive, child.id, permanent), {
    label: "SharedDrives",
    batchSize: 10,
    onItemSettled: async (child, result) => {
      const fileStatus = result.ok ? result.value : "failed";
      const errorMessage = result.ok ? null : classifySharedDriveDeleteError(result.error).message;
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
