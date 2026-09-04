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
