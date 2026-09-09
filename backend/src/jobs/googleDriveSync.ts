import { query } from "../db/pool.js";
import { getDirectoryClientAs, getDriveClientAs } from "../services/googleWorkspaceAuth.js";
import { getUserDriveUsage, listDomainUsers, type BasicGoogleUser } from "../graph/googleDriveEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import {
  isCancelled,
  markProcessing,
  settleResource,
  upsertConnectionUser,
  type ResourceRowMap,
} from "./cloudSyncWorker.js";

/**
 * Syncs Google My Drive — the Google equivalent of cloudSyncWorker.ts's syncOneDrive, called from
 * that file's job processor when connections.cloud_type === 'google_my_drive'. Structurally
 * different from every M365 sync* function: Microsoft's app-only client is one shared client for
 * the whole tenant; Google's domain-wide delegation requires impersonating a SPECIFIC user per
 * Drive call, so a per-user client is built inside the runThrottled callback below, not once
 * outside the loop. getDriveClientAs's underlying client cache means repeat calls for the same
 * user are still cheap.
 */
export async function syncGoogleMyDrive(connectionId: string, syncJobId: string, users: BasicGoogleUser[], resourceRows: ResourceRowMap): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  await runThrottled(
    users,
    async (user) => {
      const drive = await getDriveClientAs(user.email);
      return getUserDriveUsage(drive);
    },
    {
      isCancelled: () => isCancelled(syncJobId),
      label: "GoogleMyDrive",
      batchSize: 20,
      onItemSettled: async (user, result) => {
        if (result.ok && result.value !== null) {
          const usage = result.value;
          await upsertConnectionUser(connectionId, {
            graphUserId: user.id,
            upn: user.email,
            displayName: user.displayName,
            storageUsedBytes: usage.usedBytes,
            itemCount: usage.itemCount,
            syncStatus: "synced",
            errorMessage: null,
          });
          await settleResource(resourceRows, user.id, true, null);
        } else {
          failed++;
          const errorMessage = result.ok ? "No Drive access for this user" : String(result.error);
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

/** 401/403s that specifically indicate domain-wide delegation was revoked or scopes changed, not a transient per-item failure. Google's client library shape (.code / .errors[0].reason) differs from Graph's (.statusCode) — a distinct check, not a shared function with isReauthError. */
export function isGoogleReauthError(err: unknown): boolean {
  const code = (err as { code?: number })?.code;
  const reason = (err as { errors?: { reason?: string }[] })?.errors?.[0]?.reason ?? "";
  return code === 401 || code === 403 || /unauthorized_client|invalid_grant|forbidden|domainPolicy/i.test(reason);
}

/**
 * Lists domain users for a full-domain (unscoped) Google My Drive sync — the Google equivalent of
 * cloudSyncWorker.ts's `listAllUsers(client)` call for onedrive, kept in this file (rather than
 * re-exported from googleDriveEnumeration.ts) purely so cloudSyncWorker.ts's import list only ever
 * needs one Google-specific import for the full sync path.
 */
export async function listAllGoogleDomainUsers(adminUpn: string) {
  const directory = await getDirectoryClientAs(adminUpn);
  const domain = adminUpn.split("@")[1]!;
  return listDomainUsers(directory, domain);
}
