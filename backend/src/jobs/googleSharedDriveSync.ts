import { query } from "../db/pool.js";
import { getDriveClientAs } from "../services/googleWorkspaceAuth.js";
import { getSharedDriveUsage, listAllSharedDrives, type BasicSharedDrive } from "../graph/googleSharedDriveEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import { isCancelled, markProcessing, settleResource, upsertConnectionUser, type ResourceRowMap } from "./cloudSyncWorker.js";

/**
 * Syncs Google Shared Drives — mirrors jobs/googleDriveSync.ts's syncGoogleMyDrive, but every call
 * is impersonated as the connection's admin (adminUpn), never a per-resource user, since a Shared
 * Drive has no owning user. connection_users reuse: graph_user_id = the shared drive's id, upn =
 * its display name (no email/URL equivalent — see migrations/015's comment).
 */
export async function syncSharedDrives(adminUpn: string, connectionId: string, syncJobId: string, drives: BasicSharedDrive[], resourceRows: ResourceRowMap): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  const client = await getDriveClientAs(adminUpn);
  await runThrottled(
    drives,
    (d) => getSharedDriveUsage(client, d.id),
    {
      isCancelled: () => isCancelled(syncJobId),
      label: "SharedDrives",
      batchSize: 10,
      onItemSettled: async (d, result) => {
        if (result.ok) {
          await upsertConnectionUser(connectionId, {
            graphUserId: d.id,
            upn: d.name,
            displayName: d.name,
            storageUsedBytes: result.value.usedBytes,
            itemCount: result.value.itemCount,
            syncStatus: "synced",
            errorMessage: null,
          });
          await settleResource(resourceRows, d.id, true, null);
        } else {
          failed++;
          const errorMessage = String(result.error);
          await upsertConnectionUser(connectionId, {
            graphUserId: d.id,
            upn: d.name,
            displayName: d.name,
            storageUsedBytes: 0,
            itemCount: 0,
            syncStatus: "failed",
            errorMessage,
          });
          await settleResource(resourceRows, d.id, false, errorMessage);
        }
        await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
      },
    }
  );
  return failed;
}

export async function listAllSharedDrivesForAdmin(adminUpn: string): Promise<BasicSharedDrive[]> {
  const client = await getDriveClientAs(adminUpn);
  return listAllSharedDrives(client);
}
