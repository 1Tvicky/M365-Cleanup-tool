import { query } from "../db/pool.js";
import { getGmailClientAs } from "../services/googleWorkspaceAuth.js";
import { getMailboxSummary } from "../graph/gmailEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import { isCancelled, markProcessing, settleResource, upsertConnectionUser, type ResourceRowMap } from "./cloudSyncWorker.js";
import type { BasicGoogleUser } from "../graph/googleDriveEnumeration.js";

/**
 * Syncs Gmail mailboxes — mirrors jobs/googleDriveSync.ts's syncGoogleMyDrive exactly, one user
 * per impersonated client. storageUsedBytes is always 0 — Gmail has no storage-bytes API field at
 * all (see graph/gmailEnumeration.ts), same "not meaningful" convention as Outlook mailboxes.
 */
export async function syncGmail(connectionId: string, syncJobId: string, users: BasicGoogleUser[], resourceRows: ResourceRowMap): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  await runThrottled(
    users,
    async (user) => {
      const gmail = await getGmailClientAs(user.email);
      return getMailboxSummary(gmail, user.email);
    },
    {
      isCancelled: () => isCancelled(syncJobId),
      label: "Gmail",
      batchSize: 20,
      onItemSettled: async (user, result) => {
        if (result.ok && result.value !== null) {
          await upsertConnectionUser(connectionId, {
            graphUserId: user.id,
            upn: user.email,
            displayName: user.displayName,
            storageUsedBytes: 0,
            itemCount: result.value.itemCount,
            syncStatus: "synced",
            errorMessage: null,
          });
          await settleResource(resourceRows, user.id, true, null);
        } else {
          failed++;
          const errorMessage = result.ok ? "No Gmail mailbox for this user" : String(result.error);
          await upsertConnectionUser(connectionId, {
            graphUserId: user.id,
            upn: user.email,
            displayName: user.displayName,
            storageUsedBytes: 0,
            itemCount: 0,
            syncStatus: "failed",
            errorMessage,
          });
          await settleResource(resourceRows, user.id, false, errorMessage);
        }
        await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
      },
    }
  );
  return failed;
}
