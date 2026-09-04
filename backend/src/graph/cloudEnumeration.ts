import type { Client } from "@microsoft/microsoft-graph-client";

/**
 * Graph calls specific to Add Clouds / Manage Clouds enumeration (jobs/cloudSyncWorker.ts) — kept
 * separate from graph/onedrive.ts, graph/sharepoint.ts, graph/teams.ts, which back the (separate,
 * out-of-scope-here) Cleanup preview/execute flow.
 */

export interface BasicUser {
  id: string;
  upn: string;
  displayName: string | null;
}

/**
 * Paginates GET /users tenant-wide. Large tenants can have thousands of users — this is the slow
 * part of a OneDrive/Teams sync.
 *
 * Deliberately unfiltered: an earlier version scoped this to accountEnabled=true and
 * userType=Member, on the assumption that guests/disabled accounts were inflating the count past
 * what an admin would recognize as "their users." A real tenant's exported Microsoft 365 admin
 * center "Active users" list (392 rows) proved that assumption wrong — that view includes guest
 * accounts (the `name_domain.com#EXT#@tenant.onmicrosoft.com` UPN pattern) and blocked-credential
 * accounts too; "active" there means "not soft-deleted," not "enabled Member account." Matching
 * that same definition (i.e. matching what /users already returns, since Graph excludes
 * soft-deleted objects by default) is what makes our total agree with what the customer sees in
 * their own admin center.
 */
export async function listAllUsers(client: Client): Promise<BasicUser[]> {
  const users: BasicUser[] = [];
  let url: string | undefined = "/users?$select=id,userPrincipalName,displayName&$top=999";

  while (url) {
    const res: any = await client.api(url).get();
    for (const u of res.value as any[]) {
      users.push({ id: u.id, upn: u.userPrincipalName, displayName: u.displayName ?? null });
    }
    url = res["@odata.nextLink"];
  }
  return users;
}

export interface DriveQuota {
  usedBytes: number;
  itemCount: number;
}

/**
 * `quota.used` gives total storage; Graph has no cheap recursive item count, so `itemCount` here
 * is the root folder's immediate child count (not recursive) — same flat-walk trade-off as
 * graph/onedrive.ts's listUserFiles. Good enough for the Manage Clouds summary figure; the
 * Cleanup module's own preview does a real walk when it actually needs exact counts.
 */
export async function getUserDriveQuota(client: Client, userId: string): Promise<DriveQuota | null> {
  try {
    const drive: any = await client.api(`/users/${userId}/drive`).select("quota,root").get();
    return {
      usedBytes: drive.quota?.used ?? 0,
      itemCount: drive.root?.folder?.childCount ?? 0,
    };
  } catch (err) {
    // No OneDrive provisioned for this user (common — not every licensed user has touched OneDrive) — not a sync failure.
    if ((err as { statusCode?: number })?.statusCode === 404) return null;
    throw err;
  }
}

export interface SiteSummary {
  id: string;
  webUrl: string;
  displayName: string;
}

/**
 * Site discovery via the search API (search-index-backed — see docs/graph-api-limitations.md for
 * the completeness/freshness trade-off vs. a full site-collection crawl).
 */
export async function searchSites(client: Client, query = "*"): Promise<SiteSummary[]> {
  const sites: SiteSummary[] = [];
  let url: string | undefined = `/sites?search=${encodeURIComponent(query)}&$top=200`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const s of res.value as any[]) {
      sites.push({ id: s.id, webUrl: s.webUrl, displayName: s.displayName ?? s.name ?? s.webUrl });
    }
    url = res["@odata.nextLink"];
  }
  return sites;
}

export async function getSiteDriveQuota(client: Client, siteId: string): Promise<DriveQuota | null> {
  try {
    const drive: any = await client.api(`/sites/${siteId}/drive`).select("quota,root").get();
    return {
      usedBytes: drive.quota?.used ?? 0,
      itemCount: drive.root?.folder?.childCount ?? 0,
    };
  } catch (err) {
    if ((err as { statusCode?: number })?.statusCode === 404) return null;
    throw err;
  }
}

/** Count of teams a user has joined — used as the Teams connection's per-user `item_count` (storage size isn't meaningful for Teams, left at 0). */
export async function getUserJoinedTeamsCount(client: Client, userId: string): Promise<number> {
  let count = 0;
  let url: string | undefined = `/users/${userId}/joinedTeams?$select=id`;
  while (url) {
    const res: any = await client.api(url).get();
    count += (res.value as any[]).length;
    url = res["@odata.nextLink"];
  }
  return count;
}

/*
 * Below: Cleaning module (discovery phase) additions. Teams has no equivalent of the OneDrive/
 * SharePoint enumeration above — nothing existing lists actual teams/channels/chats or message
 * counts, only a per-user joined-teams count. See docs/graph-api-limitations.md for the message
 * count / chat-listing constraints these functions work around.
 */

export interface TeamSummary {
  id: string;
  displayName: string;
}

/**
 * There is no direct "list all teams" Graph endpoint under application permissions — the standard
 * approach is filtering the tenant's groups down to the ones provisioned as a Team. Uses the
 * existing Group.ReadWrite.All permission (no new grant needed).
 */
export async function listAllTeams(client: Client): Promise<TeamSummary[]> {
  const teams: TeamSummary[] = [];
  let url: string | undefined =
    "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName&$top=999";

  while (url) {
    const res: any = await client.api(url).get();
    for (const g of res.value as any[]) {
      teams.push({ id: g.id, displayName: g.displayName ?? g.id });
    }
    url = res["@odata.nextLink"];
  }
  return teams;
}

export interface ChannelSummary {
  id: string;
  displayName: string;
}

export async function listChannels(client: Client, teamId: string): Promise<ChannelSummary[]> {
  const channels: ChannelSummary[] = [];
  let url: string | undefined = `/teams/${teamId}/channels?$select=id,displayName`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const c of res.value as any[]) {
      channels.push({ id: c.id, displayName: c.displayName ?? c.id });
    }
    url = res["@odata.nextLink"];
  }
  return channels;
}

/**
 * Graph has no message-count endpoint for a channel — the only way to get one is to paginate
 * every root message (and each root message's own reply thread) and count them. Requires the
 * ChannelMessage.Read.All application permission, which is NOT part of this app's current
 * permission set — see docs/azure-ad-app-registration.md. Deliberately not using the bulk
 * `channel: getAllMessages` export endpoint: Microsoft documents it as a metered/billed capability
 * aimed at compliance/eDiscovery export, not a fit for a plain discovery count.
 */
export async function countChannelMessages(client: Client, teamId: string, channelId: string): Promise<number> {
  let count = 0;
  let url: string | undefined = `/teams/${teamId}/channels/${channelId}/messages?$top=50`;

  while (url) {
    const res: any = await client.api(url).get();
    const messages = res.value as any[];
    count += messages.length;
    for (const m of messages) {
      if (m.replies?.length !== undefined) continue; // replies aren't expanded inline by this endpoint
      const replyCount = await countChannelMessageReplies(client, teamId, channelId, m.id);
      count += replyCount;
    }
    url = res["@odata.nextLink"];
  }
  return count;
}

async function countChannelMessageReplies(client: Client, teamId: string, channelId: string, messageId: string): Promise<number> {
  let count = 0;
  let url: string | undefined = `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?$top=50`;
  while (url) {
    const res: any = await client.api(url).get();
    count += (res.value as any[]).length;
    url = res["@odata.nextLink"];
  }
  return count;
}

export interface ChatSummary {
  id: string;
  chatType: "oneOnOne" | "group" | "meeting" | "unknownFutureValue";
  topic: string | null;
  lastUpdatedDateTime: string | null;
  participants: { displayName: string | null; upn: string | null }[];
}

/**
 * GET /chats (list every chat in the tenant) is not supported under application permissions —
 * confirmed against current Microsoft Graph docs. The only way to discover 1:1/group chats
 * tenant-wide is enumerating each user's own chats and merging by chat id (a 1:1 chat appears in
 * both participants' lists) — see docs/graph-api-limitations.md. Uses the existing Chat.Read.All
 * permission, no new grant needed.
 */
export async function listUserChats(client: Client, userId: string): Promise<ChatSummary[]> {
  const chats: ChatSummary[] = [];
  let url: string | undefined =
    `/users/${userId}/chats?$expand=members&$select=id,chatType,topic,lastUpdatedDateTime&$top=50`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const c of res.value as any[]) {
      const participants = ((c.members as any[]) ?? []).map((m) => ({
        displayName: m.displayName ?? null,
        upn: m.email ?? m.userPrincipalName ?? null,
      }));
      chats.push({
        id: c.id,
        chatType: c.chatType ?? "unknownFutureValue",
        topic: c.topic ?? null,
        lastUpdatedDateTime: c.lastUpdatedDateTime ?? null,
        participants,
      });
    }
    url = res["@odata.nextLink"];
  }
  return chats;
}

/** Same "no count endpoint, must paginate" situation as countChannelMessages. */
export async function countChatMessages(client: Client, chatId: string): Promise<number> {
  let count = 0;
  let url: string | undefined = `/chats/${chatId}/messages?$top=50`;
  while (url) {
    const res: any = await client.api(url).get();
    count += (res.value as any[]).length;
    url = res["@odata.nextLink"];
  }
  return count;
}
