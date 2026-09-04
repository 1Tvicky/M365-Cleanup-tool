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

/** Cleanup (deletion) execution phase — see docs/cleanup-execution plan. Only onedrive_account/sharepoint_site are ever actually executed against Graph; channel/chat always resolve to 'unsupported' (Microsoft Graph has no application-permission path to delete Teams channel/chat messages — delegated-only). */
export type CleanupResourceType = "onedrive_account" | "sharepoint_site" | "channel" | "chat";
export type CleanupOperationStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type CleanupItemStatus = "pending" | "processing" | "completed" | "failed" | "skipped" | "unsupported";

/** One slot per resource family; `ids` reference the same internal row ids already used by the existing selection state (connection_users.id / cleaning_channels.id / cleaning_chats.id) — never raw Microsoft Graph ids. */
export interface CleanupManifest {
  oneDrive?: { connectionId: string; ids: string[] };
  sharePoint?: { connectionId: string; ids: string[] };
  channels?: { connectionId: string; ids: string[] };
  chats?: { connectionId: string; ids: string[] };
}

export interface CleanupValidationResult {
  valid: boolean;
  summary: { oneDriveAccounts: number; sharePointSites: number; channels: number; chats: number };
  /** Selected items that can never be executed under this app's Graph permissions — reported here, not in errors, since selecting them isn't invalid, just not actionable yet. */
  unsupported: { resourceType: CleanupResourceType; displayName: string }[];
  errors: string[];
  /** Ids (from the submitted manifest) that resolved successfully, grouped by slot — lets the frontend reconcile a selection against the latest sync (drop ids no longer found) without a separate endpoint. */
  foundIds: { oneDrive: string[]; sharePoint: string[]; channels: string[]; chats: string[] };
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
}

export interface CleanupProgress extends CleanupOperationRow {
  byType: Record<CleanupResourceType, { total: number; completed: number; failed: number; skipped: number; unsupported: number }>;
  /** Sum of cleanup_operation_items.files_total/files_completed across the whole operation — 0/0 until file enumeration for at least one item has happened. */
  filesTotal: number;
  filesCompleted: number;
}

/** One row of the live "recently removed" feed on the progress screen — the same data the CSV report is built from. */
export interface CleanupRecentFile {
  fileName: string;
  resourceName: string;
  status: "deleted" | "already_gone" | "failed";
  completedAt: string;
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
    onedrive?: { status: CleaningSyncResourceStatus; error: string | null };
    sharepoint?: { status: CleaningSyncResourceStatus; error: string | null };
    teams?: { status: CleaningSyncResourceStatus; error: string | null };
  };
}
