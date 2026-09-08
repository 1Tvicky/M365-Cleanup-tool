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

/**
 * Some Exchange Online backends (confirmed live on at least one real tenant) don't expose `size`
 * on message objects at all — not "it's 0", the property genuinely isn't returned even with no
 * $select restriction — and requesting it in $select makes Graph reject the *entire* query with
 * `400 Could not find a property named 'size' on type 'Microsoft.OutlookServices.Message'`. Left
 * unhandled, that 400 kills mail cleanup for every message in every folder for that mailbox. So
 * this asks for `size` first (the common case — real byte reporting), and only on that exact error
 * falls back to a `size`-free query for the rest of this folder, reporting 0 for its messages
 * rather than failing the whole mailbox.
 */
function isMissingSizePropertyError(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  const message = String((err as { message?: string })?.message ?? "");
  return status === 400 && /property named 'size'/i.test(message);
}

/** Messages directly in one folder (not its child folders — those are separate entries from listMailFoldersRecursive, walked independently). */
export async function listFolderMessages(client: Client, userId: string, folderId: string): Promise<MailMessage[]> {
  const messages: MailMessage[] = [];
  const basePath = `/users/${userId}/mailFolders/${folderId}/messages`;
  let url: string | undefined = `${basePath}?$select=id,subject,size&$top=200`;
  let sizeSupported = true;

  while (url) {
    let res: any;
    try {
      res = await client.api(url).get();
    } catch (err) {
      if (sizeSupported && isMissingSizePropertyError(err)) {
        sizeSupported = false;
        url = `${basePath}?$select=id,subject&$top=200`;
        continue;
      }
      throw err;
    }
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

/**
 * Outlook calendar. Deliberately separate functions from the Mail ones above — own Graph paths,
 * never a shared "which resource" helper (see the Outlook isolation note in
 * docs/azure-ad-app-registration.md).
 *
 * Unlike mail folders, a user's calendars don't nest and are never deleted here — only events
 * inside them. A calendar's events are only reachable by first knowing which calendars exist
 * (GET /users/{id}/events alone only covers the *default* calendar).
 */

/**
 * Every calendar the user can actually delete events from — flat list, no recursion (Graph
 * calendars don't nest under a user). Filtered to canEdit=true: confirmed live that Graph's own
 * auto-generated per-user calendars ("Birthdays", "United States holidays" — owned by the user but
 * canEdit=false) reject event deletion with "Read-only calendars can't be modified." Skipping them
 * here (rather than attempting and always failing) mirrors how Mail never attempts to delete a
 * distinguished folder — see graph/cloudEnumeration.ts's getUserCalendarEventCount for the matching
 * filter on the discovery/summary side, so a selected mailbox's shown event count is only ever
 * events this function can actually reach.
 */
export async function listUserCalendars(client: Client, userId: string): Promise<string[]> {
  const calendarIds: string[] = [];
  let url: string | undefined = `/users/${userId}/calendars?$select=id,canEdit&$top=100`;
  while (url) {
    const res: any = await client.api(url).get();
    for (const cal of res.value as any[]) if (cal.canEdit) calendarIds.push(cal.id);
    url = res["@odata.nextLink"];
  }
  return calendarIds;
}

export interface CalendarEvent {
  id: string;
  subject: string;
}

/** Events in one calendar. Graph's event objects carry no reliable size field (unlike mail/drive items), so there is no size to capture here. */
export async function listCalendarEvents(client: Client, userId: string, calendarId: string): Promise<CalendarEvent[]> {
  const events: CalendarEvent[] = [];
  let url: string | undefined = `/users/${userId}/calendars/${calendarId}/events?$select=id,subject&$top=200`;
  while (url) {
    const res: any = await client.api(url).get();
    for (const e of res.value as any[]) {
      events.push({ id: e.id, subject: e.subject || "(no subject)" });
    }
    url = res["@odata.nextLink"];
  }
  return events;
}

/** Deletes one event. Event ids are addressable directly (no need to repeat the calendar id) — mirrors deleteMessage's 404-is-already-gone convention. Never deletes a calendar itself. */
export async function deleteCalendarEvent(client: Client, userId: string, eventId: string): Promise<DriveItemDeleteResult> {
  try {
    await client.api(`/users/${userId}/events/${eventId}`).delete();
    return "deleted";
  } catch (err) {
    if ((err as { statusCode?: number })?.statusCode === 404) return "already_gone";
    throw err;
  }
}

/**
 * Outlook contacts. Deliberately separate functions again — own Graph paths, own module-level
 * concerns, never shared with the Mail or Calendar functions above via a parameter.
 *
 * Contact folders nest (unlike calendars), so this mirrors listMailFoldersRecursive's BFS shape.
 * Contacts with no folder live directly under /users/{id}/contacts — represented here as a `null`
 * folderId, not a real folder, since there is no folder object to ever protect from deletion there.
 */

/** Every contact folder id, at every depth — BFS over childFolders. Folders are never deleted, only listed. */
export async function listUserContactFoldersRecursive(client: Client, userId: string): Promise<string[]> {
  const folderIds: string[] = [];
  let queue: string[] = [];

  let url: string | undefined = `/users/${userId}/contactFolders?$select=id&$top=100`;
  while (url) {
    const res: any = await client.api(url).get();
    for (const folder of res.value as any[]) queue.push(folder.id);
    url = res["@odata.nextLink"];
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
  return folderIds;
}

export interface Contact {
  id: string;
  displayName: string;
}

/** Contacts directly in one folder, or the root "Contacts" collection when folderId is null. */
export async function listFolderContacts(client: Client, userId: string, folderId: string | null): Promise<Contact[]> {
  const contacts: Contact[] = [];
  let url: string | undefined =
    folderId === null
      ? `/users/${userId}/contacts?$select=id,displayName&$top=200`
      : `/users/${userId}/contactFolders/${folderId}/contacts?$select=id,displayName&$top=200`;
  while (url) {
    const res: any = await client.api(url).get();
    for (const c of res.value as any[]) {
      contacts.push({ id: c.id, displayName: c.displayName || "(no name)" });
    }
    url = res["@odata.nextLink"];
  }
  return contacts;
}

/** Deletes one contact. Same 404-is-already-gone idempotency convention as deleteMessage/deleteCalendarEvent. Never deletes a contact folder. */
export async function deleteContact(client: Client, userId: string, contactId: string): Promise<DriveItemDeleteResult> {
  try {
    await client.api(`/users/${userId}/contacts/${contactId}`).delete();
    return "deleted";
  } catch (err) {
    if ((err as { statusCode?: number })?.statusCode === 404) return "already_gone";
    throw err;
  }
}
