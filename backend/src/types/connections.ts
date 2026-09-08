export type CloudType = "onedrive" | "sharepoint" | "teams" | "outlook";
export type ConnectionStatus = "connecting" | "active" | "error" | "needs_reauth" | "disconnected";
export type SyncJobStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type ConnectionUserSyncStatus = "pending" | "synced" | "failed";

export const CLOUD_TYPES: readonly CloudType[] = ["onedrive", "sharepoint", "teams", "outlook"];

export function isCloudType(value: string): value is CloudType {
  return (CLOUD_TYPES as readonly string[]).includes(value);
}

export interface ManageCloudsRow {
  id: string;
  cloudType: CloudType;
  iconKey: CloudType;
  displayName: string;
  adminEmail: string;
  adminDisplayName: string | null;
  tenantDomain: string;
  totalUsers: number;
  /** Attempted so far (success + failure) — drives the row's percent, i.e. "how much of the job is done", not "how much succeeded". */
  processedUsers: number;
  /** Succeeded — the "X" in "X out of Y Users" on the row, and "Added Users" in the expand panel. */
  addedUsers: number;
  /** Failed/ineligible (e.g. no OneDrive provisioned) — "Users Not Added" in the expand panel. */
  notAddedUsers: number;
  percent: number;
  status: ConnectionStatus;
  multiUser: true;
  connectedAt: string | null;
  lastSyncedAt: string | null;
  lastError: string | null;
}

export interface ConnectionUserRow {
  id: string;
  graphUserId: string;
  upn: string;
  displayName: string | null;
  storageUsedBytes: number;
  itemCount: number;
  syncStatus: ConnectionUserSyncStatus;
  lastSyncedAt: string | null;
  errorMessage: string | null;
}

/** One resource available to sync for a workload — the id is a live Graph id, not a connection_users row id, since the browse step must work even before anything has ever been synced. */
export interface AvailableResourceRow {
  id: string;
  displayName: string;
  /** Secondary label shown alongside displayName — UPN for onedrive/outlook/teams-members-N/A, webUrl for sharepoint. Omitted where not meaningful. */
  secondary?: string;
}

export type SyncJobResourceStatus = "pending" | "processing" | "completed" | "failed" | "cancelled";

/** One selected resource within a resource-scoped sync_jobs run — mirrors CleanupOperationItemRow's relationship to CleanupOperationRow. */
export interface SyncJobResourceRow {
  id: string;
  graphResourceId: string;
  displayName: string;
  secondary: string | null;
  status: SyncJobResourceStatus;
  errorMessage: string | null;
  startedAt: string | null;
  completedAt: string | null;
}

export interface ConnectInitResponse {
  authorizeUrl: string;
  state: string;
}

export type M365ConnectMessage =
  | { type: "m365-connect-complete"; status: "success"; connectionId: string; cloudType: CloudType }
  | { type: "m365-connect-complete"; status: "error"; cloudType: CloudType | null; reason: string };

/** Signed `state` payload round-tripped through the Microsoft authorize redirect. */
export interface ConnectState {
  cloudType: CloudType;
  operatorId: string;
  nonce: string;
  iat: number;
}
