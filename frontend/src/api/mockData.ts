import type { CleanupJob, CloudConnector, Preview, Tenant } from "../types";

// Phase 1 (M365 only) — mirrors CloudFuze's "Business Clouds" tile grid, scoped to the three
// workloads this tool cleans up. Broader clouds (Google, Box, Slack, etc.) are out of scope here.
export const CONNECTORS: CloudConnector[] = [
  { id: "onedrive", label: "OneDrive for business" },
  { id: "sharepoint", label: "SharePoint online" },
  { id: "teams", label: "Microsoft Teams" },
];

export const MOCK_TENANTS: Tenant[] = [
  {
    id: "t-1",
    displayName: "Contoso Ltd",
    m365TenantId: "8f14e45f-ceea-4c5c-8f1e-000000000001",
    status: "connected",
    connectedAt: "2026-08-12T14:03:00Z",
    connectedByAdminUpn: "admin@contoso.com",
    lastTokenRefreshAt: "2026-09-03T08:00:00Z",
    workloads: ["teams", "onedrive", "sharepoint"],
  },
  {
    id: "t-2",
    displayName: "Fabrikam Inc",
    m365TenantId: "8f14e45f-ceea-4c5c-8f1e-000000000002",
    status: "connected",
    connectedAt: "2026-07-28T09:15:00Z",
    connectedByAdminUpn: "it-admin@fabrikam.com",
    lastTokenRefreshAt: "2026-09-02T22:10:00Z",
    workloads: ["teams", "onedrive", "sharepoint"],
  },
  {
    id: "t-3",
    displayName: "Northwind Traders",
    m365TenantId: "8f14e45f-ceea-4c5c-8f1e-000000000003",
    status: "consent_pending",
    connectedAt: null,
    connectedByAdminUpn: null,
    lastTokenRefreshAt: null,
    workloads: [],
  },
  {
    id: "t-4",
    displayName: "Adatum Corp",
    m365TenantId: "8f14e45f-ceea-4c5c-8f1e-000000000004",
    status: "token_error",
    connectedAt: "2026-05-01T10:00:00Z",
    connectedByAdminUpn: "admin@adatum.com",
    lastTokenRefreshAt: "2026-08-20T10:00:00Z",
    workloads: ["teams", "onedrive"],
  },
];

export const MOCK_PREVIEW: Preview = {
  previewId: "pv-1001",
  generatedAt: "2026-09-03T10:00:00Z",
  totals: { itemCount: 4821, totalSizeBytes: 812_345_678_912 },
  rows: [
    { label: "Migration-Archive team", category: "Teams channel", itemCount: 12, sizeBytes: 0 },
    { label: "Project-Falcon team", category: "Teams channel", itemCount: 6, sizeBytes: 0 },
    { label: "jane.doe@contoso.com", category: "OneDrive files", itemCount: 2140, sizeBytes: 214_748_364_800 },
    { label: "sam.lee@contoso.com", category: "OneDrive files", itemCount: 1830, sizeBytes: 178_253_611_008 },
    { label: "Finance Document Library", category: "SharePoint library", itemCount: 833, sizeBytes: 419_343_703_104 },
    {
      label: "jane.doe@contoso.com",
      category: "Chat (report-only)",
      itemCount: 58,
      sizeBytes: 0,
      note: "Not deletable in v1 — shown for visibility only",
    },
  ],
  warnings: [
    "Project-Falcon team includes 1 private channel — deletion is not restorable via Graph once confirmed.",
    "Estimated batch time is ~14 minutes at current Graph throttling limits for this tenant.",
  ],
};

export const MOCK_JOBS: CleanupJob[] = [
  {
    jobId: "job-9001",
    tenantName: "Contoso Ltd",
    status: "completed",
    progress: { total: 3210, completed: 3210, failed: 0 },
    startedAt: "2026-09-01T09:00:00Z",
    finishedAt: "2026-09-01T09:22:00Z",
    bytesReclaimed: 512_450_000_000,
    confirmedByEmail: "vignesh.t@cloudfuze.com",
  },
  {
    jobId: "job-9002",
    tenantName: "Fabrikam Inc",
    status: "completed_with_errors",
    progress: { total: 1890, completed: 1864, failed: 26 },
    startedAt: "2026-08-29T13:00:00Z",
    finishedAt: "2026-08-29T13:31:00Z",
    bytesReclaimed: 288_900_000_000,
    confirmedByEmail: "ops@cloudfuze.com",
  },
  {
    jobId: "job-9003",
    tenantName: "Contoso Ltd",
    status: "running",
    progress: { total: 4821, completed: 1560, failed: 2 },
    startedAt: "2026-09-03T09:58:00Z",
    finishedAt: null,
    bytesReclaimed: 0,
    confirmedByEmail: "vignesh.t@cloudfuze.com",
  },
];
