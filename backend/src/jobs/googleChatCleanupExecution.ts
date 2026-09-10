import { query } from "../db/pool.js";
import { getChatAdminClientAs, getChatClientAsMember, getDirectoryClientAs } from "../services/googleWorkspaceAuth.js";
import { classifyChatDeleteError, deleteChatMessage } from "../graph/googleChatDeletion.js";
import { getChatSpaceMembership, listSpaceMessages } from "../graph/googleChatEnumeration.js";
import { getGoogleUserById } from "../graph/googleDriveEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import type { PendingItem } from "./cleanupExecutionWorker.js";

/**
 * Google Chat cleanup execution. `permanent` is accepted for signature consistency with every
 * other Google execute* function but deliberately unused — see graph/googleChatDeletion.ts's
 * header comment: Chat's messages.delete has no soft-delete/trash alternative, so there is only
 * one way to delete a message here regardless of the operation's deletion_mode.
 *
 * Never deletes the Space itself, and never deletes messages authored by a bot/app it can't
 * impersonate — only messages reachable through an actual human member's own delegated access
 * (see the member-resolution steps below; messages.list/messages.delete have no admin-access
 * bypass, verified against the API docs, unlike spaces.search/members.list).
 */
export async function executeGoogleChatSpaceItem(adminUpn: string, item: PendingItem, operationId: string, _permanent: boolean): Promise<void> {
  const spaceId = item.graph_ref.spaceId!;

  const adminChat = await getChatAdminClientAs(adminUpn);
  const membership = await getChatSpaceMembership(adminChat, spaceId);
  if (!membership.aHumanMemberId) {
    // No human member to impersonate — nothing in this space is reachable under domain-wide
    // delegation (Chat has no admin-bypass for message access). Reported as a clear failure, not
    // a silent no-op, so the report explains why this space's messages weren't removed.
    throw new Error("This space has no human members to act on its behalf — Google Chat has no admin-level access to message content.");
  }

  const directory = await getDirectoryClientAs(adminUpn);
  const member = await getGoogleUserById(directory, membership.aHumanMemberId);
  if (!member) {
    throw new Error("Could not resolve the space member needed to remove its messages.");
  }

  const memberChat = await getChatClientAsMember(member.email);
  const messages = await listSpaceMessages(memberChat, spaceId);
  if (messages.length === 0) return;

  // Seed one row per message up front — same convention as executeItem/executeSharedDriveItem,
  // so the report/live "recently removed" feed reflects the full message list from the start.
  // graph_item_id stores the message's full resource name (spaces/{s}/messages/{m}), the exact id
  // deleteChatMessage/onItemSettled below key on — messages have no meaningful byte size, so
  // file_size_bytes is always 0 here.
  for (const m of messages) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, 0)`,
      [item.id, m.name, m.name]
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, messages.length]);

  let firstError: unknown = null;
  await runThrottled(messages, (m) => deleteChatMessage(memberChat, m.name), {
    label: "GoogleChat",
    batchSize: 10,
    onItemSettled: async (m, result) => {
      const status = result.ok ? result.value : "failed";
      const errorMessage = result.ok ? null : classifyChatDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, m.name, status, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}
