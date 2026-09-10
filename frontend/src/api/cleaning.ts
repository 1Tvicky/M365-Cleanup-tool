import { rawFetch } from "./client";
import type { CloudType } from "./clouds";

export interface CleaningConnectionRow {
  id: string;
  cloudType: CloudType;
  displayName: string;
  adminEmail: string;
  adminDisplayName: string | null;
  status: "connecting" | "active" | "error" | "needs_reauth" | "disconnected";
  lastSyncedAt: string | null;
}

export interface CleaningResourceRow {
  id: string;
  name: string;
  detail: string;
  storageUsedBytes: number;
  itemCount: number;
  status: "pending" | "synced" | "failed";
  lastSyncedAt: string | null;
  /**
   * Set for a while after a permanent-deletion cleanup completed against this resource, to explain
   * why storageUsedBytes/itemCount might still look pre-deletion — not "hasn't synced since the
   * delete." Microsoft's own storage-quota recalculation runs asynchronously on their backend and
   * can lag a real deletion by minutes or longer, even past a sync that ran after the delete.
   * - 'recent': within the first 24h — the common case.
   * - 'verify': 24h–7d out — the rare long tail; softer wording, suggests checking independently.
   * - null: no recent permanent deletion, or more than 7 days have passed.
   */
  deletionRecalcHint: "recent" | "verify" | null;
}

export type CountStatus = "pending" | "calculating" | "completed" | "failed";

export interface CleaningScanRow {
  id: string;
  scanType: "teams_structure" | "message_counts";
  status: "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
  totalItems: number;
  processedItems: number;
  createdAt: string;
  startedAt: string | null;
  finishedAt: string | null;
}

export interface CleaningTeamsSummary {
  teamCount: number;
  channelCount: number;
  chatCount: number;
  messagesCountedSoFar: number;
  itemsAwaitingCount: number;
  itemsFailedCount: number;
  structureScan: CleaningScanRow | null;
  countScan: CleaningScanRow | null;
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

export interface CleaningChatRow {
  id: string;
  chatType: "oneOnOne" | "group" | "meeting" | "unknownFutureValue";
  participants: { displayName: string | null; upn: string | null }[];
  messageCount: number | null;
  countStatus: CountStatus;
  lastMessageAt: string | null;
}

export function listCleaningConnections(): Promise<{ connections: CleaningConnectionRow[] }> {
  return rawFetch("/api/cleaning/connections");
}

export interface PageResult<T> {
  total: number;
  page: number;
  pageSize: number;
}

interface ListOpts {
  search?: string;
  sort?: "storage" | "name";
  page?: number;
  pageSize?: number;
}

function toQuery(opts: ListOpts): string {
  const params = new URLSearchParams();
  if (opts.search) params.set("search", opts.search);
  if (opts.sort) params.set("sort", opts.sort);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return params.toString();
}

export function listOneDriveAccounts(connectionId: string, opts: ListOpts = {}): Promise<{ accounts: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/onedrive?${toQuery(opts)}`);
}

export function listGoogleMyDriveAccounts(connectionId: string, opts: ListOpts = {}): Promise<{ accounts: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/google-my-drive?${toQuery(opts)}`);
}

export function listSharedDrives(connectionId: string, opts: ListOpts = {}): Promise<{ drives: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/shared-drives?${toQuery(opts)}`);
}

export function listGmailMailboxes(connectionId: string, opts: ListOpts = {}): Promise<{ mailboxes: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/gmail?${toQuery(opts)}`);
}

export function listGoogleChatSpaces(connectionId: string, opts: ListOpts = {}): Promise<{ spaces: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/google-chat?${toQuery(opts)}`);
}

export function listSharePointSites(connectionId: string, opts: ListOpts = {}): Promise<{ sites: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/sharepoint?${toQuery(opts)}`);
}

export function listOutlookMailboxes(connectionId: string, opts: ListOpts = {}): Promise<{ mailboxes: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/outlook?${toQuery(opts)}`);
}

/** Own function, not a resource-parameterized variant of listOutlookMailboxes — same connection, different Graph resource kind (calendar events, not mail). */
export function listOutlookCalendars(connectionId: string, opts: ListOpts = {}): Promise<{ calendars: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/outlook/calendar?${toQuery(opts)}`);
}

/** Own function, same reasoning as listOutlookCalendars above. */
export function listOutlookContacts(connectionId: string, opts: ListOpts = {}): Promise<{ contacts: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/outlook/contacts?${toQuery(opts)}`);
}

export interface OutlookOverviewSubResource {
  id: string;
  itemCount: number;
  status: "pending" | "synced" | "failed";
}

/** One row per mailbox with all three resources together — backs the unified Mail/Calendar/Contacts selection table. A read-only presentational join, not a resource-parameterized fetch — see the backend route's own comment. */
export interface OutlookMailboxOverviewRow {
  upn: string;
  name: string;
  // mail always exists — connection_users is the base table this view is built from, not a LEFT
  // JOIN. calendar/contacts are null until that resource has synced for this user.
  mail: OutlookOverviewSubResource;
  calendar: OutlookOverviewSubResource | null;
  contacts: OutlookOverviewSubResource | null;
}

export function listOutlookOverview(
  connectionId: string,
  opts: { search?: string; page?: number; pageSize?: number } = {}
): Promise<{ mailboxes: OutlookMailboxOverviewRow[] } & PageResult<OutlookMailboxOverviewRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/outlook/overview?${toQuery(opts)}`);
}

export function getTeamsSummary(connectionId: string): Promise<CleaningTeamsSummary> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/teams/summary`);
}

export function listTeamsChannels(connectionId: string, opts: ListOpts = {}): Promise<{ channels: CleaningChannelRow[] } & PageResult<CleaningChannelRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/teams/channels?${toQuery(opts)}`);
}

export function listTeamsDMs(connectionId: string, opts: ListOpts = {}): Promise<{ chats: CleaningChatRow[] } & PageResult<CleaningChatRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/teams/dms?${toQuery(opts)}`);
}

export function calculateTeamsMessageCounts(connectionId: string): Promise<{ status: "queued" }> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/teams/calculate-counts`, { method: "POST" });
}

/**
 * Cleanup (deletion) execution. 'onedrive_account'/'sharepoint_site'/'outlook_mailbox' items are
 * ever actually removed — Microsoft Graph has no application-permission (unattended) path to
 * delete Teams channel or chat messages, so 'channel'/'chat' items always resolve to 'unsupported',
 * never a faked success. See the cleanup-execution plan for the full rationale.
 */
export type CleanupResourceType =
  | "onedrive_account"
  | "sharepoint_site"
  | "outlook_mailbox"
  | "outlook_calendar"
  | "outlook_contacts"
  | "channel"
  | "chat"
  | "google_my_drive_account"
  | "shared_drive"
  | "gmail_mailbox"
  | "google_chat_space";
export type CleanupOperationStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type CleanupItemStatus = "pending" | "processing" | "completed" | "failed" | "skipped" | "unsupported";

/** ids reference the same internal row ids already used by the existing selection Maps (connection_users.id / connection_outlook_calendars.id / connection_outlook_contacts.id / cleaning_channels.id / cleaning_chats.id) — never raw Microsoft Graph ids. */
export interface CleanupManifest {
  oneDrive?: { connectionId: string; ids: string[] };
  sharePoint?: { connectionId: string; ids: string[] };
  outlook?: { connectionId: string; ids: string[] };
  outlookCalendar?: { connectionId: string; ids: string[] };
  outlookContacts?: { connectionId: string; ids: string[] };
  channels?: { connectionId: string; ids: string[] };
  chats?: { connectionId: string; ids: string[] };
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
    googleMyDriveAccounts: number;
    sharedDrives: number;
    gmailMailboxes: number;
    googleChatSpaces: number;
  };
  unsupported: { resourceType: CleanupResourceType; displayName: string }[];
  errors: string[];
  /** Ids from the submitted manifest that resolved successfully, grouped by slot — used to reconcile a selection after a sync (drop ids no longer found). */
  foundIds: {
    oneDrive: string[];
    sharePoint: string[];
    outlook: string[];
    outlookCalendar: string[];
    outlookContacts: string[];
    channels: string[];
    chats: string[];
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
  /** A connection's own display_name (e.g. "cloudfuze.co") touched by this operation — not tenants.display_name, which can legitimately differ from what every other screen shows. */
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
  /** filesCompleted is "settled" (succeeded or failed), not "succeeded" — 0/0 until file enumeration for this item has started. */
  filesTotal: number;
  filesCompleted: number;
}

export interface CleanupProgress extends CleanupOperationRow {
  byType: Record<CleanupResourceType, { total: number; completed: number; failed: number; skipped: number; unsupported: number }>;
  /** Sum across every OneDrive account / SharePoint site in the operation — 0/0 until file enumeration for at least one item has happened. */
  filesTotal: number;
  filesCompleted: number;
  /** bytesTotal counts every discovered file; bytesCleared only files actually removed ('deleted'/'already_gone', never 'failed'/'pending') — the number to show as "data cleared". */
  bytesTotal: number;
  bytesCleared: number;
}

/** One row of the live "recently removed" feed on the progress screen. */
export interface CleanupRecentFile {
  fileName: string;
  resourceName: string;
  status: "deleted" | "already_gone" | "failed";
  completedAt: string;
}

export function validateCleanup(manifest: CleanupManifest): Promise<CleanupValidationResult> {
  return rawFetch(`/api/cleaning/cleanup/validate`, { method: "POST", body: JSON.stringify(manifest) });
}

export type CleanupDeletionMode = "recycle_bin" | "permanent";

/** deletionMode is a sibling field to the manifest, not part of CleanupManifest's own structure — see routes/cleaning.ts's deletionModeSchema. Omitted/invalid defaults to the safer 'recycle_bin' server-side. */
export function startCleanup(
  manifest: CleanupManifest,
  deletionMode: CleanupDeletionMode
): Promise<{ operationId: string; status: "queued" }> {
  return rawFetch(`/api/cleaning/cleanup`, { method: "POST", body: JSON.stringify({ ...manifest, deletionMode }) });
}

export function getCleanupProgress(operationId: string): Promise<CleanupProgress> {
  return rawFetch(`/api/cleaning/cleanup/${operationId}`);
}

/** Paginated list of this operator's tenant's cleanup operations, newest first — powers the Reports page. */
export function listCleanupOperations(
  opts: { status?: CleanupOperationStatus; search?: string; page?: number; pageSize?: number } = {}
): Promise<{ operations: CleanupOperationRow[] } & PageResult<CleanupOperationRow>> {
  const params = new URLSearchParams();
  if (opts.status) params.set("status", opts.status);
  if (opts.search) params.set("search", opts.search);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/cleaning/cleanup/operations?${params.toString()}`);
}

/** Aggregate totals across every operation this operator can see — backs the Reports page's top stat strip. */
export interface CleanupOperationsSummary {
  totalOperations: number;
  processedItems: number;
  totalItems: number;
  bytesCleared: number;
  bytesTotal: number;
  updatedAt: string;
}

export function getCleanupOperationsSummary(): Promise<CleanupOperationsSummary> {
  return rawFetch(`/api/cleaning/cleanup/operations/summary`);
}

export function getCleanupOperationItems(
  operationId: string,
  opts: { status?: CleanupItemStatus; resourceType?: CleanupResourceType; page?: number; pageSize?: number } = {}
): Promise<{ items: CleanupOperationItemRow[] } & PageResult<CleanupOperationItemRow>> {
  const params = new URLSearchParams();
  if (opts.status) params.set("status", opts.status);
  if (opts.resourceType) params.set("resourceType", opts.resourceType);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/cleaning/cleanup/${operationId}/items?${params.toString()}`);
}

/** One item's file list — the third drill-down level (operation → item → file) on the progress/results screens. */
export interface CleanupItemFileRow {
  id: string;
  fileName: string;
  status: "pending" | "deleted" | "already_gone" | "failed";
  fileSizeBytes: number;
  errorMessage: string | null;
  completedAt: string | null;
}

export function getCleanupOperationItemFiles(
  operationId: string,
  itemId: string,
  opts: { page?: number; pageSize?: number } = {}
): Promise<{ files: CleanupItemFileRow[] } & PageResult<CleanupItemFileRow>> {
  const params = new URLSearchParams();
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/cleaning/cleanup/${operationId}/items/${itemId}/files?${params.toString()}`);
}

export function cancelCleanup(operationId: string): Promise<{ status: "cancel_requested" }> {
  return rawFetch(`/api/cleaning/cleanup/${operationId}/cancel`, { method: "POST" });
}

export function getCleanupRecentFiles(operationId: string, limit = 10): Promise<{ files: CleanupRecentFile[] }> {
  return rawFetch(`/api/cleaning/cleanup/${operationId}/recent-files?limit=${limit}`);
}

/** Not a rawFetch call — the report is a CSV file download (Content-Disposition: attachment), which a plain navigation/anchor click handles natively (cookies included automatically, same origin). */
export function cleanupReportUrl(operationId: string): string {
  return `/api/cleaning/cleanup/${operationId}/report`;
}

export function retryCleanup(operationId: string): Promise<{ operationId: string; status: "queued" }> {
  return rawFetch(`/api/cleaning/cleanup/${operationId}/retry`, { method: "POST" });
}

/**
 * "Sync Now" — a thin tenant-level wrapper around the existing OneDrive/SharePoint sync and Teams
 * discovery scan mechanisms. No new discovery logic on the backend; this just lets the Cleaning
 * page trigger and poll all of a tenant's connections with one id instead of up to three.
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

export function startSync(connectionIds: string[]): Promise<{ operationId: string; status: "queued" }> {
  return rawFetch(`/api/cleaning/sync`, { method: "POST", body: JSON.stringify({ connectionIds }) });
}

export function getSyncOperation(operationId: string): Promise<CleaningSyncOperation> {
  return rawFetch(`/api/cleaning/sync/operations/${operationId}`);
}

/**
 * The most recent sync for this tenant, if any — lets the Dashboard resume tracking a sync after
 * navigating away and back (or reloading), since sync progress otherwise only ever lived in the
 * Dashboard component's local state and was lost the moment it unmounted.
 */
export function getLatestSyncOperation(connectionIds: string[]): Promise<{ operation: CleaningSyncOperation | null }> {
  return rawFetch(`/api/cleaning/sync/latest?connectionIds=${connectionIds.join(",")}`);
}
