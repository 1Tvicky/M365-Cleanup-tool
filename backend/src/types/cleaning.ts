/** Cleaning module (discovery phase) — reuses connections/tenant_roles from types/connections.ts. */

export type CleaningScanType = "teams_structure" | "message_counts";
export type CleaningScanStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type CountStatus = "pending" | "calculating" | "completed" | "failed";

export interface CleaningScanRow {
  id: string;
  scanType: CleaningScanType;
  status: CleaningScanStatus;
  totalItems: number;
  processedItems: number;
  createdAt: string;
  startedAt: string | null;
  finishedAt: string | null;
}

/** Backs both the OneDrive accounts table and the SharePoint sites table — same underlying connection_users row shape. */
export interface CleaningResourceRow {
  id: string;
  name: string;
  detail: string; // upn for OneDrive, site webUrl for SharePoint
  storageUsedBytes: number;
  itemCount: number;
  status: "pending" | "synced" | "failed";
  lastSyncedAt: string | null;
  /**
   * Set for a while after a permanent-deletion cleanup completed against this resource, to explain
   * why storageUsedBytes/itemCount might still look pre-deletion — not "hasn't synced since the
   * delete." Microsoft's own storage-quota recalculation runs asynchronously on their backend and
   * can lag a real deletion by minutes or longer, so even a sync that ran after the delete can still
   * return the old figure; there's no reliable way to know from here when Graph has actually caught
   * up, so both windows below (see routes/cleaning.ts) are time-based heuristics, not guarantees.
   * - 'recent': within the first 24h — the common case, reads as "still catching up."
   * - 'verify': 24h–7d out — Microsoft publishes no SLA for this recalculation, so rather than
   *   silently dropping the hint (making a still-wrong number look normal) or keeping the same
   *   urgency indefinitely, this tapers to a softer "check the admin center" suggestion for the
   *   rare long tail.
   * - null: no recent permanent deletion, or more than 7 days have passed.
   */
  deletionRecalcHint: "recent" | "verify" | null;
}

export interface CleaningChannelRow {
  id: string;
  teamId: string;
  teamName: string;
  channelId: string;
  channelName: string;
  messageCount: number | null;
  countStatus: CountStatus;
}

export interface CleaningChatParticipant {
  displayName: string | null;
  upn: string | null;
}

export interface CleaningChatRow {
  id: string;
  chatType: "oneOnOne" | "group" | "meeting" | "unknownFutureValue";
  participants: CleaningChatParticipant[];
  messageCount: number | null;
  countStatus: CountStatus;
  lastMessageAt: string | null;
}

export interface CleaningTeamsSummary {
  teamCount: number;
  channelCount: number;
  chatCount: number;
  /** Sum of message_count for channels+chats whose countStatus is 'completed' — never includes pending/calculating/failed, so it's never a fake/partial number presented as final. */
  messagesCountedSoFar: number;
  /** How many channel+chat rows are still pending/calculating — 0 once every row has settled one way or another (completed OR failed). */
  itemsAwaitingCount: number;
  /** How many channel+chat rows gave up (e.g. ChannelMessage.Read.All not yet granted) — kept separate from itemsAwaitingCount so a fully-failed connection reads as "unable to calculate," never as a fake "0 messages." */
  itemsFailedCount: number;
  structureScan: CleaningScanRow | null;
  countScan: CleaningScanRow | null;
}

/**
 * Cleanup (deletion) execution phase — see docs/cleanup-execution plan. onedrive_account/
 * sharepoint_site/outlook_mailbox/outlook_calendar/outlook_contacts are actually executed against
 * Graph; channel/chat always resolve to 'unsupported' (Microsoft Graph has no application-
 * permission path to delete Teams channel/chat messages — delegated-only). For outlook_mailbox,
 * "delete" means delete every message in every folder — distinguished folders (Inbox, Sent Items,
 * etc.) can't be deleted themselves, only emptied. outlook_calendar/outlook_contacts are the same
 * mailbox-level granularity: selecting a user deletes all of *their* calendar events / contacts
 * across every calendar/contact folder they have — never the calendars or folders themselves.
 */
export type CleanupResourceType =
  | "onedrive_account"
  | "sharepoint_site"
  | "outlook_mailbox"
  | "outlook_calendar"
  | "outlook_contacts"
  /** A single Teams channel — DELETE /teams/{teamId}/channels/{channelId} (Channel.Delete.All). Deletes only the channel; the parent Team is untouched. */
  | "channel"
  /** A 1:1/group Teams chat — selecting one always resolves as unsupported (see routes/cleaning.ts's chats block): Graph has no application-permission path to delete chat messages, and this app never deletes the chat itself either. Unrelated to and unaffected by "team"/"channel" support below. */
  | "chat"
  /** A whole Microsoft Team — deletes the M365 Group backing it (DELETE /groups/{id}, Group.ReadWrite.All), which removes the Team and every channel under it per Graph/Entra semantics. Selecting a Team is a distinct manifest slot from selecting its channels — see resolveManifestItems' dedup of channels already covered by a selected Team. */
  | "team"
  /** Google Workspace My Drive — a user's root Drive content, deleted via domain-wide delegation impersonating them. Same "select a user, delete their root-level content, never the account itself" granularity as onedrive_account. */
  | "google_my_drive_account"
  /** Google Shared Drive — the drive's top-level content, deleted via an admin-impersonated client (a Shared Drive has no owning user). Never deletes the Shared Drive itself — see graph/googleSharedDriveDeletion.ts. */
  | "shared_drive"
  /** Gmail — every message in the mailbox, never the mailbox/user/account. Uses aggregate counters only, never one DB row per message (a mailbox can hold millions) — see graph/gmailDeletion.ts. */
  | "gmail_mailbox"
  /** Google Chat Space — the whole Space itself (spaces.delete, useAdminAccess:true — a cascading delete that also removes its messages/memberships). See graph/googleChatDeletion.ts's deleteChatSpace. */
  | "google_chat_space";
export type CleanupOperationStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type CleanupItemStatus = "pending" | "processing" | "completed" | "failed" | "skipped" | "unsupported";
/** Chosen by the operator on the confirmation screen — see migrations/013_cleanup_deletion_mode.sql. Only OneDrive/SharePoint/Outlook-mail items ever consult this; Teams/Calendar/Contacts are unaffected either way. */
export type CleanupDeletionMode = "recycle_bin" | "permanent";

/**
 * One slot per resource family; `ids` reference the same internal row ids already used by the
 * existing selection state (connection_users.id / connection_outlook_calendars.id /
 * connection_outlook_contacts.id / cleaning_channels.id / cleaning_chats.id) — never raw
 * Microsoft Graph ids.
 *
 * `teams` is the one deliberate exception: there is no dedicated per-team row anywhere (a Team is
 * only implicit as the distinct team_id values on cleaning_channels — see migrations/004), so
 * `teams.ids` holds team_id (the Microsoft Graph/Entra group id) directly. This is still never
 * trusted blindly — resolveManifestItems validates each id against cleaning_channels scoped to
 * the given connectionId before treating it as real, the same access-proof every other slot gets
 * via its row lookup.
 */
export interface CleanupManifest {
  oneDrive?: { connectionId: string; ids: string[] };
  sharePoint?: { connectionId: string; ids: string[] };
  outlook?: { connectionId: string; ids: string[] };
  outlookCalendar?: { connectionId: string; ids: string[] };
  outlookContacts?: { connectionId: string; ids: string[] };
  channels?: { connectionId: string; ids: string[] };
  chats?: { connectionId: string; ids: string[] };
  /** Whole-Team deletion — ids are team_id (Graph group ids), not row ids; see the interface comment above. */
  teams?: { connectionId: string; ids: string[] };
  googleMyDrive?: { connectionId: string; ids: string[] };
  sharedDrives?: { connectionId: string; ids: string[] };
  gmail?: { connectionId: string; ids: string[] };
  googleChat?: { connectionId: string; ids: string[] };
}

export interface CleanupValidationResult {
  valid: boolean;
  summary: {
    oneDriveAccounts: number;
    sharePointSites: number;
    outlookMailboxes: number;
    outlookCalendars: number;
    outlookContacts: number;
    channels: number;
    chats: number;
    teams: number;
    googleMyDriveAccounts: number;
    sharedDrives: number;
    gmailMailboxes: number;
    googleChatSpaces: number;
  };
  /** Selected items that can never be executed under this app's Graph permissions — reported here, not in errors, since selecting them isn't invalid, just not actionable yet. */
  unsupported: { resourceType: CleanupResourceType; displayName: string }[];
  errors: string[];
  /** Ids (from the submitted manifest) that resolved successfully, grouped by slot — lets the frontend reconcile a selection against the latest sync (drop ids no longer found) without a separate endpoint. */
  foundIds: {
    oneDrive: string[];
    sharePoint: string[];
    outlook: string[];
    outlookCalendar: string[];
    outlookContacts: string[];
    channels: string[];
    chats: string[];
    teams: string[];
    googleMyDrive: string[];
    sharedDrives: string[];
    gmail: string[];
    googleChat: string[];
  };
}

export interface CleanupOperationRow {
  id: string;
  status: CleanupOperationStatus;
  totalItems: number;
  processedItems: number;
  successfulItems: number;
  failedItems: number;
  skippedItems: number;
  retryOfOperationId: string | null;
  startedAt: string | null;
  completedAt: string | null;
  cancelRequestedAt: string | null;
  createdAt: string;
  errorMessage: string | null;
  /** null defensively (LEFT JOIN) — operators are never hard-deleted in this app, so in practice always present. */
  requestedBy: { email: string; displayName: string } | null;
  /** A connection's own display_name (e.g. "cloudfuze.co") touched by this operation — deliberately not tenants.display_name, which can legitimately differ from what every other screen shows the user. */
  label: string;
  /** What this specific operation actually did/does — chosen once at creation (or inherited by retry), never mixed within one operation. */
  deletionMode: CleanupDeletionMode;
}

export interface CleanupOperationItemRow {
  id: string;
  connectionId: string;
  resourceType: CleanupResourceType;
  displayName: string;
  status: CleanupItemStatus;
  attempts: number;
  startedAt: string | null;
  completedAt: string | null;
  errorCode: string | null;
  errorMessage: string | null;
  /** How many individual files/messages/events/contacts this item covers, and how many have settled (succeeded OR failed) so far — filesCompleted is "settled", not "succeeded"; a genuinely-succeeded count would need a separate per-status query this list route doesn't do. 0/0 until file enumeration for this item has started. */
  filesTotal: number;
  filesCompleted: number;
}

export interface CleanupProgress extends CleanupOperationRow {
  byType: Record<CleanupResourceType, { total: number; completed: number; failed: number; skipped: number; unsupported: number }>;
  /** Sum of cleanup_operation_items.files_total/files_completed across the whole operation — 0/0 until file enumeration for at least one item has happened. */
  filesTotal: number;
  filesCompleted: number;
  /** Sum of cleanup_operation_item_files.file_size_bytes — bytesTotal counts every discovered file (known as soon as it's listed), bytesCleared only those actually removed ('deleted' or 'already_gone', never 'failed'/'pending'). This is the number the client-facing report calls "data cleared". */
  bytesTotal: number;
  bytesCleared: number;
}

/** One row of the live "recently removed" feed on the progress screen — the same data the CSV report is built from. */
export interface CleanupRecentFile {
  fileName: string;
  resourceName: string;
  status: "deleted" | "already_gone" | "failed";
  completedAt: string;
}

/** One row of one item's file list — the third drill-down level (operation → item → file) on the progress/results screens. Unlike CleanupRecentFile (always-completed, operation-wide), this can be 'pending' since it's scoped to one item and may be read while that item is still being processed. */
export interface CleanupItemFileRow {
  id: string;
  fileName: string;
  status: "pending" | "deleted" | "already_gone" | "failed";
  fileSizeBytes: number;
  errorMessage: string | null;
  completedAt: string | null;
}

/**
 * "Sync Now" — a thin tenant-level wrapper around the existing sync_jobs (OneDrive/SharePoint)
 * and cleaning_scans (Teams) mechanisms. Status is never persisted for the operation itself; it's
 * computed live from whichever of the 1-3 sub-resources were actually triggered.
 */
export type CleaningSyncOperationStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed";
export type CleaningSyncResourceStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";

export interface CleaningSyncOperation {
  id: string;
  status: CleaningSyncOperationStatus;
  startedAt: string;
  completedAt: string | null;
  byResource: {
    // unavailableCount: how many accounts/sites have no provisioned drive or hit a Graph-reported
    // access block — "completed_with_errors" almost always means this, not that sync itself broke.
    onedrive?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number; unavailableCount: number };
    sharepoint?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number; unavailableCount: number };
    outlook?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number; unavailableCount: number };
    teams?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number };
  };
}
