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

/** Paginates GET /users tenant-wide. Large tenants can have thousands of users — this is the slow part of a OneDrive/Teams sync. */
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
