import type { chat_v1 } from "googleapis";

/**
 * Google Chat message deletion. Verified against the current Chat API reference: messages.delete
 * has no documented trash/recoverable state (unlike Drive's trashed flag or Gmail's trash()) — it's
 * always an immediate, unconditional removal. So unlike every other Google workload in this app,
 * Chat message deletion does NOT branch on the operation's deletion_mode — there is only one kind
 * of delete Chat's API offers. `force: true` is required to also remove threaded replies (without
 * it, deleting a message with replies fails outright) — always passed, since "delete the space's
 * content" should not silently skip whole threads.
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
