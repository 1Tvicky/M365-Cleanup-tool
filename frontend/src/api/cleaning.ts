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

export function listSharePointSites(connectionId: string, opts: ListOpts = {}): Promise<{ sites: CleaningResourceRow[] } & PageResult<CleaningResourceRow>> {
  return rawFetch(`/api/cleaning/connections/${connectionId}/sharepoint?${toQuery(opts)}`);
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
 * Cleanup (deletion) execution. Only 'onedrive_account'/'sharepoint_site' items are ever actually
 * removed — Microsoft Graph has no application-permission (unattended) path to delete Teams
 * channel or chat messages, so 'channel'/'chat' items always resolve to 'unsupported', never a
 * faked success. See the cleanup-execution plan for the full rationale.
 */
export type CleanupResourceType = "onedrive_account" | "sharepoint_site" | "channel" | "chat";
export type CleanupOperationStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type CleanupItemStatus = "pending" | "processing" | "completed" | "failed" | "skipped" | "unsupported";

/** ids reference the same internal row ids already used by the existing selection Maps (connection_users.id / cleaning_channels.id / cleaning_chats.id) — never raw Microsoft Graph ids. */
export interface CleanupManifest {
  oneDrive?: { connectionId: string; ids: string[] };
  sharePoint?: { connectionId: string; ids: string[] };
  channels?: { connectionId: string; ids: string[] };
  chats?: { connectionId: string; ids: string[] };
}

export interface CleanupValidationResult {
  valid: boolean;
  summary: { oneDriveAccounts: number; sharePointSites: number; channels: number; chats: number };
  unsupported: { resourceType: CleanupResourceType; displayName: string }[];
  errors: string[];
  /** Ids from the submitted manifest that resolved successfully, grouped by slot — used to reconcile a selection after a sync (drop ids no longer found). */
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
  /** Sum across every OneDrive account / SharePoint site in the operation — 0/0 until file enumeration for at least one item has happened. */
  filesTotal: number;
  filesCompleted: number;
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

export function startCleanup(manifest: CleanupManifest): Promise<{ operationId: string; status: "queued" }> {
  return rawFetch(`/api/cleaning/cleanup`, { method: "POST", body: JSON.stringify(manifest) });
}

export function getCleanupProgress(operationId: string): Promise<CleanupProgress> {
  return rawFetch(`/api/cleaning/cleanup/${operationId}`);
}

export function getCleanupOperationItems(
  operationId: string,
  opts: { status?: CleanupItemStatus; page?: number; pageSize?: number } = {}
): Promise<{ items: CleanupOperationItemRow[] } & PageResult<CleanupOperationItemRow>> {
  const params = new URLSearchParams();
  if (opts.status) params.set("status", opts.status);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/cleaning/cleanup/${operationId}/items?${params.toString()}`);
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
    onedrive?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number };
    sharepoint?: { status: CleaningSyncResourceStatus; error: string | null; processed: number; total: number };
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
