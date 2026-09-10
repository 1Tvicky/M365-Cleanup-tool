import { query } from "../db/pool.js";
import { getGmailClientAs } from "../services/googleWorkspaceAuth.js";
import { deleteAllMailboxMessages } from "../graph/gmailDeletion.js";
import type { PendingItem } from "./cleanupExecutionWorker.js";

/**
 * Gmail cleanup execution — deliberately NOT shaped like executeGoogleMyDriveItem/
 * executeSharedDriveItem. Those seed one cleanup_operation_item_files row per file up front (fine
 * for the dozens/thousands of files a Drive/Shared Drive item typically has); a Gmail mailbox can
 * hold millions of messages, so this only ever updates the two existing aggregate counters
 * (files_total/files_completed) on the cleanup_operation_items row itself — see
 * graph/gmailDeletion.ts's header comment for the full reasoning. Preserves the mailbox/user/
 * account; only ever deletes messages.
 */
export async function executeGmailMailboxItem(item: PendingItem, operationId: string, permanent: boolean, isCancelled: () => Promise<boolean>): Promise<void> {
  const userEmail = item.graph_ref.userEmail!;
  const gmail = await getGmailClientAs(userEmail);

  const result = await deleteAllMailboxMessages(gmail, userEmail, permanent, {
    isCancelled,
    onProgress: async (progress) => {
      await query(`UPDATE cleanup_operation_items SET files_total = $2, files_completed = $3 WHERE id = $1`, [
        item.id,
        progress.requested,
        progress.completed + progress.failed,
      ]);
    },
  });

  // Same all-or-nothing-per-item convention as executeItem/executeSharedDriveItem: any message
  // failure marks the whole mailbox item 'failed' (cleanupExecutionWorker.ts's outer
  // onItemSettled only has a binary completed/failed per item — there's no partial-success
  // status). files_total/files_completed above still carry the exact requested/completed/failed
  // reconciliation for the report, even when the item itself reads as failed.
  if (result.failed > 0) {
    throw new Error(`${result.failed.toLocaleString()} of ${result.requested.toLocaleString()} message(s) couldn't be removed. It can be retried.`);
  }
}
