import type { chat_v1 } from "googleapis";

/**
 * Google Chat deletion. Verified against the current Chat API reference: neither messages.delete
 * nor spaces.delete has a documented trash/recoverable state (unlike Drive's trashed flag or
 * Gmail's trash()) — both are immediate, unconditional removals. So unlike every other Google
 * workload in this app, Chat deletion does NOT branch on the operation's deletion_mode — there is
 * only one kind of delete Chat's API offers, for either resource.
 *
 * deleteChatMessage: `force: true` is required to also remove threaded replies (without it,
 * deleting a message with replies fails outright) — always passed, since "delete this content"
 * should not silently skip whole threads. Requires impersonating an actual human member of the
 * space (see services/googleWorkspaceAuth.ts's getChatClientAsMember) — messages.delete has no
 * admin-access bypass. Only still used where a Space's messages need clearing without deleting the
 * Space itself; the cleanup path for a *selected Space* uses deleteChatSpace below instead.
 *
 * deleteChatSpace: the whole-Space cleanup path (jobs/googleChatCleanupExecution.ts) — a cascading
 * delete of the Space and its messages/memberships together, via the admin-impersonated client
 * (no member impersonation needed for this one).
 */

export type ChatMessageDeleteResult = "deleted" | "already_gone";

function isNotFoundError(err: unknown): boolean {
  return (err as { code?: number })?.code === 404;
}

export async function deleteChatMessage(chat: chat_v1.Chat, messageName: string): Promise<ChatMessageDeleteResult> {
  try {
    await chat.spaces.messages.delete({ name: messageName, force: true });
    return "deleted";
  } catch (err) {
    if (isNotFoundError(err)) return "already_gone";
    throw err;
  }
}

/**
 * Deletes the whole Space (verified against the current Chat API reference,
 * developers.google.com/workspace/chat/api/reference/rest/v1/spaces/delete): a cascading delete —
 * Google's own docs state the space's child resources (messages, memberships) are removed along
 * with it, so there is no separate message-by-message pass needed or wanted here. Requires
 * `useAdminAccess: true` plus the `chat.admin.delete` scope (Developer Preview) on the
 * admin-impersonated client — unlike message deletion, this does NOT need a human member to
 * impersonate, since admin access covers space-level operations directly.
 */
export async function deleteChatSpace(chat: chat_v1.Chat, spaceId: string): Promise<ChatMessageDeleteResult> {
  try {
    await chat.spaces.delete({ name: `spaces/${spaceId}`, useAdminAccess: true });
    return "deleted";
  } catch (err) {
    if (isNotFoundError(err)) return "already_gone";
    throw err;
  }
}

/** Same mapping shape as classifyGoogleDeleteError — duplicated per this codebase's workload-isolation convention (see docs/google-workspace-integration.md). */
export function classifyChatDeleteError(err: unknown): { code: string; message: string } {
  const code = (err as { code?: number })?.code;
  const reason = (err as { errors?: { reason?: string }[] })?.errors?.[0]?.reason;
  if (code === 403) {
    return { code: "INSUFFICIENT_PERMISSION", message: "This connection doesn't currently have permission to delete this data." };
  }
  if (code === 429) {
    return { code: "RATE_LIMITED", message: "Google rate-limited this request. It can be retried." };
  }
  return { code: reason ?? String(code ?? "UNKNOWN"), message: String((err as { message?: string })?.message ?? err) };
}
