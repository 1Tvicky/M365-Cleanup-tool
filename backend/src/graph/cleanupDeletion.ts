import type { Client } from "@microsoft/microsoft-graph-client";

/**
 * Graph calls specific to the Cleaning module's cleanup/deletion execution phase — kept separate
 * from graph/cloudEnumeration.ts (discovery, read-only) and from the older graph/onedrive.ts /
 * graph/sharepoint.ts / graph/teams.ts (the separate, tenant-role-gated legacy Cleanup pipeline —
 * its `deleteDocumentLibrary` deletes the library container itself, which is not what this module
 * does; see the cleanup-execution plan for why these are deliberately not reused).
 *
 * Only OneDrive and SharePoint have a real, application-permission delete path. Deleting a Teams
 * channel or chat *message* requires a delegated (signed-in user present) permission — Microsoft
 * Graph does not support it for an unattended application-only service like this one — so there
 * are no equivalent functions here for channels/chats; callers must mark those items 'unsupported'
 * without ever calling Graph for them.
 */

export type DriveOwnerKind = "user" | "site";

export interface DriveRootChild {
  id: string;
  name: string;
}

function driveRootPath(kind: DriveOwnerKind, id: string): string {
  return kind === "user" ? `/users/${id}/drive/root/children` : `/sites/${id}/drive/root/children`;
}

function driveItemPath(kind: DriveOwnerKind, id: string, itemId: string): string {
  return kind === "user" ? `/users/${id}/drive/items/${itemId}` : `/sites/${id}/drive/items/${itemId}`;
}

/**
 * Lists the top-level files/folders of a user's OneDrive or a site's default document library.
 * Deleting a folder recursively removes everything inside it, so top-level children are all that's
 * needed to clear the whole drive/library's contents — never the user or the site itself.
 */
export async function listDriveRootChildren(client: Client, kind: DriveOwnerKind, id: string): Promise<DriveRootChild[]> {
  const children: DriveRootChild[] = [];
  let url: string | undefined = `${driveRootPath(kind, id)}?$select=id,name&$top=200`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const item of res.value as any[]) {
      children.push({ id: item.id, name: item.name ?? item.id });
    }
    url = res["@odata.nextLink"];
  }
  return children;
}

export type DriveItemDeleteResult = "deleted" | "already_gone";

/** Deletes one drive item (moves it to the Graph recycle bin). A 404 means it's already gone — treated as success, never a failure, so retries stay idempotent. */
export async function deleteDriveItem(client: Client, kind: DriveOwnerKind, id: string, itemId: string): Promise<DriveItemDeleteResult> {
  try {
    await client.api(driveItemPath(kind, id, itemId)).delete();
    return "deleted";
  } catch (err) {
    if ((err as { statusCode?: number })?.statusCode === 404) return "already_gone";
    throw err;
  }
}

/**
 * Distinguishes a real delete failure from "insufficient permission" (shouldn't happen —
 * Files.ReadWrite.All/Sites.ReadWrite.All are already granted — but must surface clearly rather
 * than being swallowed as a generic error) or a lockable/retryable conflict. Kept in this
 * side-effect-free module (rather than jobs/cleanupExecutionWorker.ts, which instantiates a BullMQ
 * Worker at import time) specifically so it's unit-testable without touching Redis.
 */
export function classifyDeleteError(err: unknown): { code: string; message: string } {
  const status = (err as { statusCode?: number })?.statusCode;
  if (status === 403) {
    return { code: "INSUFFICIENT_PERMISSION", message: "This connection doesn't currently have permission to delete this data." };
  }
  if (status === 409) {
    return { code: "CONFLICT", message: "This item is locked or in use and couldn't be removed. It can be retried." };
  }
  return { code: String(status ?? "UNKNOWN"), message: String((err as { message?: string })?.message ?? err) };
}
