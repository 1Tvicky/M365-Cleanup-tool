import type { chat_v1 } from "googleapis";

/**
 * Google Chat calls — split across two impersonation levels, verified against the current Chat
 * API reference (not guessed):
 *
 * - spaces.search / spaces.members.list support `useAdminAccess: true`, letting an impersonated
 *   DOMAIN ADMIN enumerate every space and space membership without being a member themselves
 *   (services/googleWorkspaceAuth.ts's getChatAdminClientAs). This module only ever uses that
 *   admin-level access for enumeration — never for reading/deleting messages.
 * - spaces.messages.list/delete have NO admin-access bypass: "Lists messages in a space that the
 *   caller is a member of" is the caller's literal restriction, with no useAdminAccess parameter
 *   documented for either method. Reading/deleting messages therefore requires impersonating an
 *   actual HUMAN member of the space (getChatClientAsMember) — see
 *   jobs/googleChatCleanupExecution.ts for how that member is chosen.
 */

export interface BasicChatSpace {
  id: string; // the "spaces/{id}" resource name's id portion
  displayName: string | null;
}

function isResourceNotFoundError(err: unknown): boolean {
  const code = (err as { code?: number })?.code;
  return code === 404 || code === 400;
}

/** Paginated spaces.search with useAdminAccess:true — every real ("SPACE" type) space in the domain, not import-mode or the admin's own memberships only. */
export async function listAllChatSpaces(chat: chat_v1.Chat): Promise<BasicChatSpace[]> {
  const spaces: BasicChatSpace[] = [];
  let pageToken: string | undefined;
  do {
    const res = await chat.spaces.search({
      useAdminAccess: true,
      query: "customer = 'customers/my_customer' AND spaceType = 'SPACE'",
      pageSize: 1000,
      pageToken,
    });
    for (const s of res.data.spaces ?? []) {
      const id = s.name?.replace(/^spaces\//, "");
      if (!id) continue;
      spaces.push({ id, displayName: s.displayName ?? null });
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return spaces;
}

export async function getChatSpaceById(chat: chat_v1.Chat, spaceId: string): Promise<BasicChatSpace | null> {
  try {
    const res = await chat.spaces.get({ name: `spaces/${spaceId}`, useAdminAccess: true });
    if (!res.data.name) return null;
    return { id: spaceId, displayName: res.data.displayName ?? null };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
}

export interface ChatSpaceMembership {
  memberCount: number;
  /** The numeric Directory/People API user id (Chat's User.name, "users/{id}") of one HUMAN member — the impersonation target for reading/deleting this space's messages. null if the space has no human members (e.g. bots only). */
  aHumanMemberId: string | null;
}

/** One admin-impersonated call gets both the member count and a candidate to impersonate for message access — avoids a second listing pass. */
export async function getChatSpaceMembership(chat: chat_v1.Chat, spaceId: string): Promise<ChatSpaceMembership> {
  let memberCount = 0;
  let aHumanMemberId: string | null = null;
  let pageToken: string | undefined;
  do {
    const res = await chat.spaces.members.list({
      parent: `spaces/${spaceId}`,
      useAdminAccess: true,
      filter: "member.type = \"HUMAN\"",
      pageSize: 1000,
      pageToken,
    });
    for (const m of res.data.memberships ?? []) {
      memberCount++;
      if (!aHumanMemberId && m.member?.name) aHumanMemberId = m.member.name.replace(/^users\//, "");
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return { memberCount, aHumanMemberId };
}

export interface ChatMessage {
  name: string; // "spaces/{space}/messages/{message}" — the exact id delete needs
}

/** Paginated spaces.messages.list — must be called with a client impersonating an actual member of the space (see this file's header comment). */
export async function listSpaceMessages(chat: chat_v1.Chat, spaceId: string): Promise<ChatMessage[]> {
  const messages: ChatMessage[] = [];
  let pageToken: string | undefined;
  do {
    const res = await chat.spaces.messages.list({ parent: `spaces/${spaceId}`, pageSize: 1000, pageToken });
    for (const m of res.data.messages ?? []) {
      if (!m.name) continue;
      messages.push({ name: m.name });
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return messages;
}
