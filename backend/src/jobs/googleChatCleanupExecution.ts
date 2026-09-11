import { getChatAdminClientAs } from "../services/googleWorkspaceAuth.js";
import { deleteChatSpace } from "../graph/googleChatDeletion.js";
import type { PendingItem } from "./cleanupExecutionWorker.js";

/**
 * Google Chat cleanup execution. `permanent` is accepted for signature consistency with every
 * other Google execute* function but deliberately unused — see graph/googleChatDeletion.ts's
 * header comment: neither of Chat's delete APIs has a soft-delete/trash alternative, so there is
 * only one way to delete a Space regardless of the operation's deletion_mode.
 *
 * Deletes the whole Space (a cascading delete that also removes its messages/memberships — see
 * deleteChatSpace) via the admin-impersonated client. No member impersonation/resolution needed
 * here (unlike the message-only deletion this function used to perform): admin access covers
 * space-level operations directly, since `spaces.delete` (unlike `spaces.messages.delete`) does
 * support the admin-access bypass — verified against the Chat API reference.
 *
 * `deleteChatSpace` treats "already gone" (404) the same as a fresh delete — both resolve
 * normally here, which the outer worker loop (cleanupExecutionWorker.ts) records as the item
 * completing successfully, the same idempotency convention used by every other resource type in
 * this app for an item with no per-file sub-rows.
 */
export async function executeGoogleChatSpaceItem(adminUpn: string, item: PendingItem, _operationId: string, _permanent: boolean): Promise<void> {
  const spaceId = item.graph_ref.spaceId!;
  const adminChat = await getChatAdminClientAs(adminUpn);
  await deleteChatSpace(adminChat, spaceId);
}
