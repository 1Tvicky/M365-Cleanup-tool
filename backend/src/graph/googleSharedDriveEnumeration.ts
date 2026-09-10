import type { drive_v3 } from "googleapis";

/**
 * Google Shared Drives calls — the Google equivalent of graph/googleDriveEnumeration.ts, but for
 * shared_drive rather than google_my_drive. Structurally different from My Drive: a Shared Drive
 * has no owning user to impersonate, so every call here uses an admin-impersonated client
 * (services/googleWorkspaceAuth.ts's getDriveClientAs(adminEmail)), never a per-resource user.
 */

export interface BasicSharedDrive {
  id: string;
  name: string;
}

function isResourceNotFoundError(err: unknown): boolean {
  const code = (err as { code?: number })?.code;
  return code === 404 || code === 400;
}

/**
 * Paginated `drives.list` with `useDomainAdminAccess: true` — verified against Drive API docs:
 * this returns every shared drive in the domain the impersonated admin administers, not just ones
 * they're personally a member of. Requires the caller to actually be a domain admin and the
 * existing `drive` scope (no new scope needed beyond what My Drive already uses).
 */
export async function listAllSharedDrives(drive: drive_v3.Drive): Promise<BasicSharedDrive[]> {
  const drives: BasicSharedDrive[] = [];
  let pageToken: string | undefined;
  do {
    const res = await drive.drives.list({ useDomainAdminAccess: true, pageSize: 100, pageToken });
    for (const d of res.data.drives ?? []) {
      if (!d.id) continue;
      drives.push({ id: d.id, name: d.name ?? "(untitled)" });
    }
    pageToken = res.data.nextPageToken ?? undefined;
  } while (pageToken);
  return drives;
}

/** Direct single-drive lookup — mirrors getGoogleUserById's role for the resource-scoped Sync flow's validation step. */
export async function getSharedDriveById(drive: drive_v3.Drive, driveId: string): Promise<BasicSharedDrive | null> {
  try {
    const res = await drive.drives.get({ driveId, useDomainAdminAccess: true });
    if (!res.data.id) return null;
    return { id: res.data.id, name: res.data.name ?? "(untitled)" };
  } catch (err) {
    if (isResourceNotFoundError(err)) return null;
    throw err;
  }
}

export interface SharedDriveItem {
  id: string;
  name: string;
  sizeBytes: number;
}

/**
 * Top-level items directly inside the shared drive (a Shared Drive's own id also names its root
 * folder, same as a My Drive's root — 'driveId in parents' lists its immediate children). Not
 * recursive, matching the same "flat walk for the summary/cleanup, folder delete cascades the
 * rest" convention already used for My Drive/OneDrive. supportsAllDrives + includeItemsFromAllDrives
 * are required for any files.list call to see shared-drive content at all.
 */
export async function listSharedDriveRootItems(drive: drive_v3.Drive, driveId: string): Promise<SharedDriveItem[]> {
  const items: SharedDriveItem[] = [];
  let pageToken: string | undefined;
  do {
    const res = await drive.files.list({
      q: `'${driveId}' in parents and trashed = false`,
      corpora: "drive",
      driveId,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
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

export interface SharedDriveUsage {
  itemCount: number;
  usedBytes: number;
}

/**
 * No cheap total-storage call exists for a Shared Drive the way `about.get` gives one for a user's
 * My Drive (Shared Drive storage counts against the pooled Workspace org quota, not a per-drive
 * quota field) — this sums the same top-level listing listSharedDriveRootItems already needs for
 * cleanup, so it's one call, not an extra enumeration, but is a non-recursive, best-effort figure
 * like every other "cheap flat walk" total in this app.
 */
export async function getSharedDriveUsage(drive: drive_v3.Drive, driveId: string): Promise<SharedDriveUsage> {
  const items = await listSharedDriveRootItems(drive, driveId);
  return { itemCount: items.length, usedBytes: items.reduce((sum, i) => sum + i.sizeBytes, 0) };
}
