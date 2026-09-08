import type { Client } from "@microsoft/microsoft-graph-client";

/**
 * Graph calls specific to the Cleaning module's cleanup/deletion execution phase — kept separate
 * from graph/cloudEnumeration.ts (discovery, read-only) and from the older graph/onedrive.ts /
 * graph/sharepoint.ts / graph/teams.ts (the separate, tenant-role-gated legacy Cleanup pipeline —
 * its `deleteDocumentLibrary` deletes the library container itself, which is not what this module
 * does; see the cleanup-execution plan for why these are deliberately not reused).
 *
 * OneDrive, SharePoint, and Outlook mail all have a real, application-permission delete path.
 * Deleting a Teams channel or chat *message* requires a delegated (signed-in user present)
 * permission — Microsoft Graph does not support it for an unattended application-only service like
 * this one — so there are no equivalent functions here for channels/chats; callers must mark those
 * items 'unsupported' without ever calling Graph for them.
 */

export type DriveOwnerKind = "user" | "site";

export interface DriveRootChild {
  id: string;
  name: string;
  /** Bytes, per Graph's driveItem.size — captured at listing time since the delete response carries no size. 0 for a folder (Graph doesn't report a meaningful size for containers) or if Graph omits the field. */
  size: number;
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
  let url: string | undefined = `${driveRootPath(kind, id)}?$select=id,name,size&$top=200`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const item of res.value as any[]) {
      children.push({ id: item.id, name: item.name ?? item.id, size: Number(item.size ?? 0) });
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

/**
 * Outlook mail. Deleting a *folder* is not an option here the way it is for a OneDrive/SharePoint
 * top-level item: distinguished/well-known folders (Inbox, Sent Items, Drafts, Deleted Items, Junk
 * Email, etc.) return ErrorDeleteDistinguishedFolder if you try — Exchange protects them from
 * deletion, only their *contents* can be cleared. So cleanup here always means recursing every
 * folder and deleting the messages inside, never the folder itself — this also means it works
 * uniformly for distinguished and custom folders alike, with no need to tell them apart.
 */

/** Every mail folder id in the mailbox, at every depth — BFS over childFolders, starting from the top-level list. */
export async function listMailFoldersRecursive(client: Client, userId: string): Promise<string[]> {
  const folderIds: string[] = [];
  let queue: string[] = [];

  let url: string | undefined = `/users/${userId}/mailFolders?$select=id&$top=100`;
  while (url) {
    const res: any = await client.api(url).get();
    for (const folder of res.value as any[]) queue.push(folder.id);
    url = res["@odata.nextLink"];
  }

  while (queue.length > 0) {
    const folderId = queue.shift()!;
    folderIds.push(folderId);
    let childUrl: string | undefined = `/users/${userId}/mailFolders/${folderId}/childFolders?$select=id&$top=100`;
    while (childUrl) {
      const res: any = await client.api(childUrl).get();
      for (const folder of res.value as any[]) queue.push(folder.id);
      childUrl = res["@odata.nextLink"];
    }
  }
  return folderIds;
}

export interface MailMessage {
  id: string;
  subject: string;
  /** Bytes, per Graph's message.size (body + attachments) — captured at listing time, same reasoning as DriveRootChild.size. */
  size: number;
}

/** Messages directly in one folder (not its child folders — those are separate entries from listMailFoldersRecursive, walked independently). */
export async function listFolderMessages(client: Client, userId: string, folderId: string): Promise<MailMessage[]> {
  const messages: MailMessage[] = [];
  let url: string | undefined = `/users/${userId}/mailFolders/${folderId}/messages?$select=id,subject,size&$top=200`;

  while (url) {
    const res: any = await client.api(url).get();
    for (const m of res.value as any[]) {
      messages.push({ id: m.id, subject: m.subject || "(no subject)", size: Number(m.size ?? 0) });
    }
    url = res["@odata.nextLink"];
  }
  return messages;
}

/** Deletes one message. A 404 means it's already gone — treated as success, never a failure, so retries stay idempotent (same convention as deleteDriveItem). */
export async function deleteMessage(client: Client, userId: string, messageId: string): Promise<DriveItemDeleteResult> {
  try {
    await client.api(`/users/${userId}/messages/${messageId}`).delete();
    return "deleted";
  } catch (err) {
    if ((err as { statusCode?: number })?.statusCode === 404) return "already_gone";
    throw err;
  }
}
