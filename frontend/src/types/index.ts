export type Workload = "teams" | "onedrive" | "sharepoint";
export type TenantStatus = "connected" | "consent_pending" | "token_error" | "disconnected";
export type JobStatus =
  | "export_in_progress"
  | "queued"
  | "running"
  | "completed"
  | "completed_with_errors"
  | "failed"
  | "cancelled";

export interface CloudConnector {
  id: Workload;
  label: string;
}

export interface Tenant {
  id: string;
  displayName: string;
  m365TenantId: string;
  status: TenantStatus;
  connectedAt: string | null;
  connectedByAdminUpn: string | null;
  lastTokenRefreshAt: string | null;
  workloads: Workload[];
}

export interface PreviewBreakdownRow {
  label: string;
  category: "Teams channel" | "OneDrive files" | "SharePoint library" | "Chat (report-only)";
  itemCount: number;
  sizeBytes: number;
  note?: string;
}

export interface Preview {
  previewId: string;
  generatedAt: string;
  totals: { itemCount: number; totalSizeBytes: number };
  rows: PreviewBreakdownRow[];
  warnings: string[];
}
