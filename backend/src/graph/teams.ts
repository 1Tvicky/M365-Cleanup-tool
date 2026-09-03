import type { Client } from "@microsoft/microsoft-graph-client";

export interface TeamSummary {
  id: string;
  displayName: string;
  channelCount: number;
  memberCount: number;
}

export interface ChannelSummary {
  id: string;
  displayName: string;
  membershipType: "standard" | "private" | "shared";
}

export async function listTeams(client: Client, search: string, limit: number): Promise<TeamSummary[]> {
  const filter = search
    ? `startswith(displayName,'${search.replace(/'/g, "''")}')`
    : undefined;
  const res = await client
    .api("/groups")
    .filter(["resourceProvisioningOptions/Any(x:x eq 'Team')", filter].filter(Boolean).join(" and "))
    .top(limit)
    .get();

  return (res.value as any[]).map((g) => ({
    id: g.id,
    displayName: g.displayName,
    channelCount: 0, // resolved lazily per-team in preview to avoid N calls on every list render
    memberCount: 0,
  }));
}

export async function listChannels(client: Client, teamId: string): Promise<ChannelSummary[]> {
  const res = await client.api(`/teams/${teamId}/channels`).get();
  return (res.value as any[]).map((c) => ({
    id: c.id,
    displayName: c.displayName,
    membershipType: c.membershipType ?? "standard",
  }));
}

/**
 * Deletes a channel. Private channels have NO Graph restore endpoint — see
 * docs/rollback-safety.md. Standard channels are restorable for 21 days via
 * POST /teams/{id}/channels/{id}/restore, but this function only performs the delete; restore is
 * an out-of-band admin action, not part of this tool's v1 scope.
 */
export async function deleteChannel(client: Client, teamId: string, channelId: string): Promise<void> {
  await client.api(`/teams/${teamId}/channels/${channelId}`).delete();
}

/** Removes the M365 Group backing a Team once its channels are cleaned up (Azure AD soft-delete, 30 days). */
export async function deleteGroup(client: Client, groupId: string): Promise<void> {
  await client.api(`/groups/${groupId}`).delete();
}

export interface ChatSummary {
  chatId: string;
  approxSizeBytes: number;
}

/**
 * Report-only in v1: chat/DM counts and approximate sizes for the preview breakdown. No delete
 * counterpart exists here on purpose — see docs/azure-ad-app-registration.md, Chat.Read.All is
 * granted for this call only, deletion is deferred to v2.
 */
export async function listUserChatsForReport(client: Client, userId: string): Promise<ChatSummary[]> {
  const res = await client.api(`/users/${userId}/chats`).get();
  return (res.value as any[]).map((c) => ({ chatId: c.id, approxSizeBytes: 0 }));
}
