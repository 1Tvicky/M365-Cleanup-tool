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
