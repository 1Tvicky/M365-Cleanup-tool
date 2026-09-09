import type { admin_directory_v1, drive_v3 } from "googleapis";

/**
 * Google Workspace calls specific to Add Clouds / Manage Clouds enumeration for My Drive
 * (jobs/googleDriveSync.ts) — the Google equivalent of graph/cloudEnumeration.ts. Every function
 * here takes an already-impersonated client (see services/googleWorkspaceAuth.ts) — this module
 * never builds auth itself.
 */

export interface BasicGoogleUser {
  id: string;
  email: string;
  displayName: string | null;
}

function isResourceNotFoundError(err: unknown): boolean {
  const code = (err as { code?: number })?.code;
  return code === 404 || code === 400;
}

/**
 * Paginates Directory API users.list for one domain. Mirrors cloudEnumeration.ts's listAllUsers —
 * deliberately unfiltered (no isSuspended/orgUnit filter) to match "every user the admin center
 * shows," same reasoning as the Graph side.
 */
export async function listDomainUsers(directory: admin_directory_v1.Admin, domain: string): Promise<BasicGoogleUser[]> {
  const users: BasicGoogleUser[] = [];
  let pageToken: string | undefined;
  do {
    const res = await directory.users.list({ domain, maxResults: 500, pageToken });
    for (const u of res.data.users ?? []) {
      if (!u.id || !u.primaryEmail) continue;
      users.push({ id: u.id, email: u.primaryEmail, displayName: u.name?.fullName ?? null });
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return users;
}

/**
 * Direct single-user lookup — the resource-scoped Sync flow's validation step, mirrors
 * cloudEnumeration.ts's getUserById: confirming a handful of selected ids never requires
 * paginating the whole domain. Returns null on a not-found-shaped error, never throws for a
 * client-submitted garbage id.
 */
export async function getGoogleUserById(directory: admin_directory_v1.Admin, userId: string): Promise<BasicGoogleUser | null> {
  try {
    const res = await directory.users.get({ userKey: userId });
    const u = res.data;
    if (!u.id || !u.primaryEmail) return null;
    return { id: u.id, email: u.primaryEmail, displayName: u.name?.fullName ?? null };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
}

export interface GoogleDriveUsage {
  usedBytes: number;
  itemCount: number;
}

/**
 * about.get's storageQuota is per-user and closer to real-time than the Admin SDK Reports API's
 * usageReport (which documents a ~2-day lag) — see docs/google-workspace-integration.md for the
 * tradeoff writeup and the still-open question of whether this field itself lags similarly (to be
 * confirmed against a live test domain). itemCount here is the immediate root-folder child count
 * (one extra files.list call), matching Graph's root.folder.childCount — not a recursive walk.
 */
export async function getUserDriveUsage(drive: drive_v3.Drive): Promise<GoogleDriveUsage | null> {
  try {
    const about = await drive.about.get({ fields: "storageQuota" });
    const usedBytes = Number(about.data.storageQuota?.usage ?? 0);
    let itemCount = 0;
    let pageToken: string | undefined;
    do {
      const res = await drive.files.list({
        q: "'root' in parents and trashed = false",
        fields: "nextPageToken, files(id)",
        pageSize: 1000,
        pageToken,
      });
      itemCount += res.data.files?.length ?? 0;
      pageToken = res.data.nextPageToken ?? undefined;
    } while (pageToken);
    return { usedBytes, itemCount };
  } catch (err) {
    // No Drive provisioned / access denied for this user — not a sync failure, same convention as getUserDriveQuota.
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
}

export interface GoogleDriveItem {
  id: string;
  name: string;
  sizeBytes: number;
}

/**
 * Root-level (top-level, non-recursive) items only — mirrors the existing OneDrive cleanup
 * semantics of "delete selected root children," never a full recursive walk up front (child files
 * inside a deleted folder are removed by Drive's own cascade, same as Graph's driveItem delete).
 * size is absent for native Google Docs/Sheets/Slides (no bytes on disk) — defaults to 0, no
 * special-case retry needed (Drive doesn't reject the query the way Graph does for some
 * Outlook-only fields).
 */
export async function listRootDriveItems(drive: drive_v3.Drive): Promise<GoogleDriveItem[]> {
  const items: GoogleDriveItem[] = [];
  let pageToken: string | undefined;
  do {
    const res = await drive.files.list({
      q: "'root' in parents and trashed = false",
      fields: "nextPageToken, files(id, name, size)",
      pageSize: 1000,
      pageToken,
    });
    for (const f of res.data.files ?? []) {
      if (!f.id) continue;
      items.push({ id: f.id, name: f.name ?? "(untitled)", sizeBytes: Number(f.size ?? 0) });
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return items;
}
