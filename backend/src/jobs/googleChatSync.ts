import { query } from "../db/pool.js";
import { getChatAdminClientAs } from "../services/googleWorkspaceAuth.js";
import { getChatSpaceMembership, listAllChatSpaces, type BasicChatSpace } from "../graph/googleChatEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import { isCancelled } from "./cloudSyncWorker.js";
import type { ResourceRowMap } from "./cloudSyncWorker.js";

/**
 * Syncs Google Chat spaces into connection_google_spaces (migrations/015) — its own dedicated
 * table, not connection_users, since a Space isn't a per-user resource (see the migration's
 * comment). Every call here impersonates the tenant's admin — enumeration and membership counting
 * both use spaces.search/members.list's admin-access bypass, never a per-space member (that's only
 * needed later, at cleanup time, for message-level access — see jobs/googleChatCleanupExecution.ts).
 */

async function upsertSpace(connectionId: string, spaceId: string, displayName: string | null, memberCount: number, messageCount: number, status: "synced" | "failed", errorMessage: string | null): Promise<void> {
  await query(
    `INSERT INTO connection_google_spaces (connection_id, space_id, display_name, member_count, message_count, sync_status, error_message, last_synced_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, now())
     ON CONFLICT (connection_id, space_id) DO UPDATE SET
       display_name = EXCLUDED.display_name,
       member_count = EXCLUDED.member_count,
       message_count = EXCLUDED.message_count,
       sync_status = EXCLUDED.sync_status,
       error_message = EXCLUDED.error_message,
       last_synced_at = now()`,
    [connectionId, spaceId, displayName, memberCount, messageCount, status, errorMessage]
  );
}

async function pruneStaleSpaces(connectionId: string, currentIds: string[]): Promise<void> {
  await query(`DELETE FROM connection_google_spaces WHERE connection_id = $1 AND NOT (space_id = ANY($2::text[]))`, [connectionId, currentIds]);
}

export async function syncGoogleChatSpaces(adminUpn: string, connectionId: string, syncJobId: string, spaces: BasicChatSpace[], resourceRows: ResourceRowMap): Promise<number> {
  let failed = 0;
  if (resourceRows && resourceRows.size > 0) {
    await query(`UPDATE sync_job_resources SET status = 'processing', started_at = now() WHERE id = ANY($1::uuid[])`, [[...resourceRows.values()]]);
  }
  const chat = await getChatAdminClientAs(adminUpn);
  await runThrottled(
    spaces,
    (s) => getChatSpaceMembership(chat, s.id),
    {
      isCancelled: () => isCancelled(syncJobId),
      label: "GoogleChat",
      batchSize: 10,
      onItemSettled: async (s, result) => {
        const rowId = resourceRows?.get(s.id);
        if (result.ok) {
          // message_count here is a placeholder (0) — an accurate count would mean impersonating a
          // member and paginating spaces.messages.list for every space during sync, which is
          // exactly the expensive per-space enumeration this sync pass is meant to avoid; the exact
          // count is only computed at actual cleanup time (jobs/googleChatCleanupExecution.ts),
          // same tradeoff Teams channels already accept in cleaning_channels (count_status
          // 'pending' until a separate, explicit count pass runs).
          await upsertSpace(connectionId, s.id, s.displayName, result.value.memberCount, 0, "synced", null);
          if (rowId) await query(`UPDATE sync_job_resources SET status = 'completed', completed_at = now(), updated_at = now() WHERE id = $1`, [rowId]);
        } else {
          failed++;
          const errorMessage = String(result.error);
          await upsertSpace(connectionId, s.id, s.displayName, 0, 0, "failed", errorMessage);
          if (rowId) await query(`UPDATE sync_job_resources SET status = 'failed', error_message = $2, completed_at = now(), updated_at = now() WHERE id = $1`, [rowId, errorMessage]);
        }
        await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
      },
    }
  );
  if (!resourceRows) await pruneStaleSpaces(connectionId, spaces.map((s) => s.id));
  return failed;
}

export async function listAllChatSpacesForAdmin(adminUpn: string): Promise<BasicChatSpace[]> {
  const chat = await getChatAdminClientAs(adminUpn);
  return listAllChatSpaces(chat);
}
