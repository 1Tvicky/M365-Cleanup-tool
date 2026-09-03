export type TenantStatus = "connected" | "consent_pending" | "token_error" | "disconnected";
export type Workload = "teams" | "onedrive" | "sharepoint";
export type OperatorRole = "viewer" | "cleanup_admin";
export type JobStatus =
  | "export_in_progress"
  | "queued"
  | "running"
  | "completed"
  | "completed_with_errors"
  | "failed"
  | "cancelled";

export interface Operator {
  id: string;
  email: string;
  displayName: string;
}

export interface TenantSummary {
  id: string;
  displayName: string;
  m365TenantId: string;
  status: TenantStatus;
  connectedAt: string | null;
  connectedByAdminUpn: string | null;
  lastTokenRefreshAt: string | null;
  workloads: Workload[];
}

export interface PreviewScope {
  users?: string[];
  teams?: { teamId: string; channelIds?: string[] }[];
  sites?: { siteId: string; libraryIds?: string[] }[];
  removeM365Groups?: boolean;
  fileFilter?: { olderThanCutoff?: string };
}

export interface PreviewResult {
  previewId: string;
  generatedAt: string;
  totals: { itemCount: number; totalSizeBytes: number };
  breakdown: {
    teams: { teamId: string; displayName: string; channelsToDelete: number; groupWillBeRemoved: boolean }[];
    onedrive: { userId: string; upn: string; fileCount: number; sizeBytes: number }[];
    sharepoint: { siteId: string; displayName: string; libraryCount: number; sizeBytes: number }[];
    chats: { userId: string; upn: string; chatCount: number; approxSizeBytes: number; note: string }[];
  };
  warnings: string[];
}

export interface AuditEntry {
  itemType: "channel" | "file" | "site_library" | "group";
  itemId: string;
  itemPath: string;
  sizeBytes: number;
  result: "deleted" | "failed" | "skipped";
  errorCode: string | null;
  timestamp: string;
}

export class ApiError extends Error {
  constructor(
    public status: number,
    public code: string,
    message: string,
    public details?: unknown
  ) {
    super(message);
  }
}
