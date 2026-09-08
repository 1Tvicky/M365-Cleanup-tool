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
 * A resource id that doesn't belong to this workload isn't always a clean 404 — Graph rejects a
 * malformed id (wrong shape/format, e.g. a plain string where a SharePoint site id needs its
 * composite `hostname,collectionId,siteId` form, confirmed live: "400 invalidRequest — Invalid
 * hostname for this tenancy") with 400, not 404. Both mean the same thing to a caller asking "is
 * this a valid resource for this workload" — treat them identically as "no such resource," never
 * let either surface as an unhandled exception (a client submitting garbage as a resourceId must
 * get a clean rejection, not a 500).
 */
function isResourceNotFoundError(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  return status === 404 || status === 400;
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

/**
 * Direct single-object lookup — the resource-scoped Sync flow's validation/re-derivation step uses
 * this instead of listAllUsers, so confirming a handful of selected ids never requires paginating
 * the whole tenant (see docs/azure-ad-app-registration.md and the "no accidental tenant-wide
 * enumeration" requirement this backs). Returns null on a not-found-shaped error (404, or Graph's
 * 400 for a malformed id — see isResourceNotFoundError), not just a real 404, since a client
 * submitting garbage as a resourceId must get a clean rejection here, never an unhandled exception.
 */
export async function getUserById(client: Client, userId: string): Promise<BasicUser | null> {
  try {
    const u: any = await client.api(`/users/${userId}`).select("id,userPrincipalName,displayName").get();
    return { id: u.id, upn: u.userPrincipalName, displayName: u.displayName ?? null };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
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

export interface MailSummary {
  itemCount: number;
}

/**
 * Unlike OneDrive/SharePoint, Graph has no single-call mailbox byte-size quota under application
 * permissions (no equivalent of drive.quota.used) — Reports.Read.All's mailbox usage report is the
 * closest thing, but it's a daily-refreshed batch export, not a live per-user call, and out of step
 * with every other cheap-per-user-call in this file. So, same trade-off already made for Teams
 * (storageUsedBytes: 0, "not meaningful") — mailbox size isn't tracked at this summary level; item
 * count (top-level mail folders' totalItemCount, summed) is. Real per-message byte sizes ARE
 * tracked during actual cleanup execution (message.size), which is where "data cleared" is reported.
 */
export async function getUserMailSummary(client: Client, userId: string): Promise<MailSummary | null> {
  try {
    let itemCount = 0;
    let url: string | undefined = `/users/${userId}/mailFolders?$select=totalItemCount&$top=100`;
    while (url) {
      const res: any = await client.api(url).get();
      for (const folder of res.value as any[]) itemCount += folder.totalItemCount ?? 0;
      url = res["@odata.nextLink"];
    }
    return { itemCount };
  } catch (err) {
    // No mailbox provisioned for this user (e.g. a licenseless or resource account) — not a sync failure.
    if ((err as { statusCode?: number })?.statusCode === 404) return null;
    throw err;
  }
}

export interface CalendarSummary {
  itemCount: number;
}

/**
 * Deliberately separate from getUserMailSummary — its own Graph calls, its own function, never a
 * shared "which resource" helper (see docs/azure-ad-app-registration.md's Outlook isolation note).
 *
 * GET /users/{id}/events only returns the *default* calendar's events — a mailbox with secondary or
 * shared calendars would silently undercount. So this enumerates every calendar the user owns first
 * (/users/{id}/calendars, which per Graph's documented behavior already includes calendars in every
 * calendar group — no separate /calendarGroups traversal needed) and sums each one's event count.
 *
 * Filtered to canEdit=true: confirmed live that Graph's own auto-generated per-user calendars
 * ("Birthdays", "United States holidays" — owned by the user but canEdit=false) reject event
 * deletion with "Read-only calendars can't be modified." Counting their events here would promise a
 * cleanup that execution can never actually perform, the same way Mail never counts a distinguished
 * folder as if it could be deleted.
 */
export async function getUserCalendarEventCount(client: Client, userId: string): Promise<CalendarSummary | null> {
  try {
    const calendarIds: string[] = [];
    let calUrl: string | undefined = `/users/${userId}/calendars?$select=id,canEdit&$top=100`;
    while (calUrl) {
      const res: any = await client.api(calUrl).get();
      for (const cal of res.value as any[]) if (cal.canEdit) calendarIds.push(cal.id);
      calUrl = res["@odata.nextLink"];
    }

    let itemCount = 0;
    for (const calendarId of calendarIds) {
      let url: string | undefined = `/users/${userId}/calendars/${calendarId}/events?$select=id&$top=999`;
      while (url) {
        const res: any = await client.api(url).get();
        itemCount += (res.value as any[]).length;
        url = res["@odata.nextLink"];
      }
    }
    return { itemCount };
  } catch (err) {
    // No mailbox provisioned for this user — same rationale as getUserMailSummary.
    if ((err as { statusCode?: number })?.statusCode === 404) return null;
    throw err;
  }
}

export interface ContactSummary {
  itemCount: number;
}

/**
 * Deliberately separate from getUserMailSummary/getUserCalendarEventCount — own Graph calls, own
 * function. Contact folders nest (like mail folders), so this BFS-walks /contactFolders +
 * /childFolders the same shape as graph/cleanupDeletion.ts's listMailFoldersRecursive, plus the
 * implicit root "Contacts" folder (contacts with no folder, reachable only via /users/{id}/contacts
 * directly — there is no folder id for it).
 */
export async function getUserContactCount(client: Client, userId: string): Promise<ContactSummary | null> {
  try {
    let itemCount = 0;

    // Root contacts (no folder).
    let rootUrl: string | undefined = `/users/${userId}/contacts?$select=id&$top=999`;
    while (rootUrl) {
      const res: any = await client.api(rootUrl).get();
      itemCount += (res.value as any[]).length;
      rootUrl = res["@odata.nextLink"];
    }

    // Every contact folder, at every depth.
    const folderIds: string[] = [];
    let queue: string[] = [];
    let topUrl: string | undefined = `/users/${userId}/contactFolders?$select=id&$top=100`;
    while (topUrl) {
      const res: any = await client.api(topUrl).get();
      for (const folder of res.value as any[]) queue.push(folder.id);
      topUrl = res["@odata.nextLink"];
    }
    while (queue.length > 0) {
      const folderId = queue.shift()!;
      folderIds.push(folderId);
      let childUrl: string | undefined = `/users/${userId}/contactFolders/${folderId}/childFolders?$select=id&$top=100`;
      while (childUrl) {
        const res: any = await client.api(childUrl).get();
        for (const folder of res.value as any[]) queue.push(folder.id);
        childUrl = res["@odata.nextLink"];
      }
    }

    for (const folderId of folderIds) {
      let url: string | undefined = `/users/${userId}/contactFolders/${folderId}/contacts?$select=id&$top=999`;
      while (url) {
        const res: any = await client.api(url).get();
        itemCount += (res.value as any[]).length;
        url = res["@odata.nextLink"];
      }
    }

    return { itemCount };
  } catch (err) {
    // No mailbox provisioned for this user — same rationale as getUserMailSummary.
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

/** Same direct-lookup rationale as getUserById, for SharePoint's site resources. */
export async function getSiteById(client: Client, siteId: string): Promise<SiteSummary | null> {
  try {
    const s: any = await client.api(`/sites/${siteId}`).select("id,webUrl,displayName,name").get();
    return { id: s.id, webUrl: s.webUrl, displayName: s.displayName ?? s.name ?? s.webUrl };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
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

/**
 * Same direct-lookup rationale as getUserById, for Teams' team resources. Also re-checks
 * resourceProvisioningOptions the same way listAllTeams' $filter already does — a plain
 * GET /groups/{id} would happily return any Microsoft 365 group, Team-provisioned or not, and a
 * non-Team group id must never be accepted as a valid Teams-workload resource.
 */
export async function getTeamById(client: Client, teamId: string): Promise<TeamSummary | null> {
  try {
    const g: any = await client.api(`/groups/${teamId}`).select("id,displayName,resourceProvisioningOptions").get();
    if (!(g.resourceProvisioningOptions as string[] | undefined)?.includes("Team")) return null;
    return { id: g.id, displayName: g.displayName ?? g.id };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
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
