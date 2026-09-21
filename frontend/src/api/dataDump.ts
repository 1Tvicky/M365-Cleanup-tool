import { rawFetch } from "./client";

/**
 * Data Dump module client — its own file, its own /api/data-dump namespace, deliberately never
 * importing from api/cleaning.ts beyond the shared rawFetch helper. Mirrors that file's naming
 * conventions (listX/getX/startX) without sharing any of its types, so a Data Dump operation can
 * never be passed where a Cleanup operation is expected, or vice versa, at the type-checker level.
 */

export type DataDumpWorkload = "onedrive" | "sharepoint" | "teams" | "outlook";
export const DATA_DUMP_WORKLOADS: DataDumpWorkload[] = ["onedrive", "sharepoint", "teams", "outlook"];

export type DataDumpProfile = "basic_demo" | "migration_demo" | "enterprise_demo" | "performance_test" | "custom";

export type DataDumpOperationStatus = "queued" | "running" | "paused" | "completed" | "completed_with_errors" | "failed" | "cancelled";

export type FileTypeKey = "docx" | "xlsx" | "pptx" | "pdf" | "txt" | "csv" | "jpg" | "png" | "zip";

export interface DateRangeConfig {
  mode: "last_30_days" | "last_6_months" | "last_1_year" | "last_3_years" | "custom";
  customStartDate?: string;
  customEndDate?: string;
}

export interface OneDriveGenConfig {
  selectedUserIds?: string[];
  userCount: number;
  /** Total number of root folders created directly under the user's OneDrive. */
  rootFolders: number;
  /** How many sub-folders each folder gets, uniformly, at every level below the root. */
  subFoldersPerFolder: number;
  /** How many levels deep the folder tree goes; 1 means root folders only, no sub-folders at all. */
  maxFolderDepth: number;
  /** Generated in EVERY folder in the tree — root, sub-folder, and nested folder alike. */
  filesPerFolder: number;
  targetTotalSizeBytes: number;
  minFileSizeBytes: number;
  maxFileSizeBytes: number;
  /** Omitted (or absent after clearing every checkbox in the File Types picker) means "use the profile's default mixed distribution" — only set this to restrict generation to specific extensions. */
  fileTypeDistribution?: Partial<Record<FileTypeKey, number>>;
  namingStyle: "professional" | "synthetic";
}

export interface NewSiteConfig {
  displayName: string;
  urlSlug: string;
  template: "teamSiteWithoutMicrosoft365Group" | "communicationSite";
  description?: string;
}

export interface SharePointGenConfig {
  selectedSiteIds?: string[];
  newSite?: NewSiteConfig;
  siteCount: number;
  librariesPerSite: number;
  /** Total number of root folders created directly under each document library. */
  foldersPerLibrary: number;
  /** How many sub-folders each folder gets, uniformly, at every level below the root. */
  subFoldersPerFolder: number;
  /** How many levels deep the folder tree goes; 1 means root folders only, no sub-folders at all. */
  maxFolderDepth: number;
  /** Generated in EVERY folder in the tree — root, sub-folder, and nested folder alike. */
  filesPerFolder: number;
  targetTotalSizeBytes: number;
  minFileSizeBytes: number;
  maxFileSizeBytes: number;
  /** Omitted (or absent after clearing every checkbox in the File Types picker) means "use the profile's default mixed distribution" — only set this to restrict generation to specific extensions. */
  fileTypeDistribution?: Partial<Record<FileTypeKey, number>>;
  generatePermissions: boolean;
  namingStyle: "professional" | "synthetic";
}

export interface ChannelConfig {
  name: string;
  memberUserIds?: string[];
  messages: number;
  replies: number;
}

export interface NewTeamConfig {
  displayName: string;
  description?: string;
  visibility: "private" | "public";
  memberUserIds: string[];
}

export interface TeamsGenConfig {
  selectedTeamIds?: string[];
  newTeam?: NewTeamConfig;
  channels?: ChannelConfig[];
  teamCount: number;
  membersPerTeam: number;
  channelsPerTeam: number;
  messagesPerChannel: number;
  repliesPerMessage: number;
  namingStyle: "professional" | "synthetic";
}

export interface OutlookGenConfig {
  selectedUserIds?: string[];
  userCount: number;
  emailsPerUser: number;
  attachmentsPerEmail: number;
  averageAttachmentSizeBytes: number;
  calendarEventCount: number;
  includeRecurringEvents: boolean;
  includeAttendees: boolean;
  contactCount: number;
  namingStyle: "professional" | "synthetic";
}

export interface DataDumpConfig {
  namingPrefix: string;
  seed?: number;
  dateRange: DateRangeConfig;
  onedrive?: OneDriveGenConfig;
  sharepoint?: SharePointGenConfig;
  teams?: TeamsGenConfig;
  outlook?: OutlookGenConfig;
}

export interface DataDumpTenant {
  tenantId: string;
  displayName: string;
  workloads: DataDumpWorkload[];
  connectionsByWorkload: Partial<Record<DataDumpWorkload, string>>;
  adminUpn: string;
  adminDisplayName: string | null;
  lastUpdatedAt: string | null;
}

export function listDataDumpTenants(): Promise<{ tenants: DataDumpTenant[] }> {
  return rawFetch(`/api/data-dump/tenants`);
}

/** Tenant-wide real-member user listing — powers the Teams member picker (spec §15), which isn't tied to any single workload connectionId the way OneDrive/Outlook selection is. Same shape as api/clouds.ts's listAvailableResources so it plugs into the same ResourceSelectionStep UI. */
export function listTenantUsers(
  tenantId: string,
  opts: { search?: string; page?: number; pageSize?: number } = {}
): Promise<{ resources: import("./clouds").AvailableResourceRow[]; total: number; page: number; pageSize: number }> {
  const params = new URLSearchParams();
  if (opts.search) params.set("search", opts.search);
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  return rawFetch(`/api/data-dump/tenants/${tenantId}/users?${params.toString()}`);
}

export function getDataDumpProfileDefaults(profile: DataDumpProfile, workloads: DataDumpWorkload[]): Promise<{ config: DataDumpConfig }> {
  const params = new URLSearchParams({ profile, workloads: workloads.join(",") });
  return rawFetch(`/api/data-dump/profile-defaults?${params.toString()}`);
}

export interface DataDumpCreateRequest {
  tenantId: string;
  profile: DataDumpProfile;
  workloads: DataDumpWorkload[];
  namingPrefix?: string;
  seed?: number;
  dateRange?: DateRangeConfig;
  onedrive?: Partial<OneDriveGenConfig>;
  sharepoint?: Partial<SharePointGenConfig>;
  teams?: Partial<TeamsGenConfig>;
  outlook?: Partial<OutlookGenConfig>;
  parentOperationId?: string;
}

export interface DataDumpPreviewWorkload {
  workload: DataDumpWorkload;
  requestedObjects: number;
  estimatedSizeBytes: number;
  breakdown: Record<string, number>;
  warnings: string[];
}

export interface DataDumpPreviewResult {
  workloads: DataDumpPreviewWorkload[];
  totalRequestedObjects: number;
  totalEstimatedSizeBytes: number;
}

export function previewDataDump(body: DataDumpCreateRequest): Promise<{ preview: DataDumpPreviewResult; config: DataDumpConfig }> {
  return rawFetch(`/api/data-dump/preview`, { method: "POST", body: JSON.stringify(body) });
}

export function startDataDump(body: DataDumpCreateRequest): Promise<{ operationId: string; status: "queued"; label: string }> {
  return rawFetch(`/api/data-dump`, { method: "POST", body: JSON.stringify(body) });
}

export interface SubcountEntry {
  requested: number;
  created: number;
  failed: number;
  skipped: number;
}

export interface DataDumpWorkloadTaskSummary {
  workload: DataDumpWorkload;
  status: string;
  requestedItems: number;
  createdItems: number;
  failedItems: number;
  skippedItems: number;
  totalSizeBytes: number;
  subcounts?: Record<string, SubcountEntry>;
  errorMessage: string | null;
}

export interface DataDumpOperationRow {
  id: string;
  /** Azure AD directory (tenant) ID for the M365 tenant this operation ran against — for correlating a run against Entra ID/Graph audit logs when troubleshooting. */
  m365TenantId: string | null;
  label: string;
  profile: DataDumpProfile;
  workloads: DataDumpWorkload[];
  status: DataDumpOperationStatus;
  requestedItems: number;
  createdItems: number;
  failedItems: number;
  skippedItems: number;
  totalSizeBytes: number;
  parentOperationId: string | null;
  startedAt: string | null;
  completedAt: string | null;
  errorMessage: string | null;
  createdAt: string;
  tasks: DataDumpWorkloadTaskSummary[];
}

export function getDataDumpOperation(operationId: string): Promise<DataDumpOperationRow> {
  return rawFetch(`/api/data-dump/${operationId}`);
}

export function listDataDumpHistory(opts: { page?: number; pageSize?: number; search?: string } = {}): Promise<{ operations: DataDumpOperationRow[]; total: number }> {
  const params = new URLSearchParams();
  params.set("page", String(opts.page ?? 1));
  params.set("pageSize", String(opts.pageSize ?? 20));
  if (opts.search) params.set("search", opts.search);
  return rawFetch(`/api/data-dump/history?${params.toString()}`);
}

export interface DataDumpSummary {
  total: number;
  running: number;
  completed: number;
  failed: number;
}

export function getDataDumpSummary(): Promise<DataDumpSummary> {
  return rawFetch(`/api/data-dump/summary`);
}

export function pauseDataDump(operationId: string): Promise<{ status: string }> {
  return rawFetch(`/api/data-dump/${operationId}/pause`, { method: "POST" });
}

export function resumeDataDump(operationId: string): Promise<{ status: string }> {
  return rawFetch(`/api/data-dump/${operationId}/resume`, { method: "POST" });
}

export function cancelDataDump(operationId: string): Promise<{ status: string }> {
  return rawFetch(`/api/data-dump/${operationId}/cancel`, { method: "POST" });
}

export function dataDumpReportUrl(operationId: string): string {
  return `/api/data-dump/${operationId}/report`;
}

export interface DataDumpContainerRow {
  id: string;
  kind: string;
  graphId: string;
  displayName: string;
  parentContainerId: string | null;
  workload: DataDumpWorkload;
  createdCount: number;
  failedCount: number;
  sizeBytes: number;
}

/** Every container (user's OneDrive, SharePoint site/library, Team/channel, Outlook mailbox) for this operation, with leaf-object counts already aggregated DB-side — see routes/dataDump.ts. The frontend builds the tree from parentContainerId. */
export function getDataDumpResources(operationId: string): Promise<{ containers: DataDumpContainerRow[] }> {
  return rawFetch(`/api/data-dump/${operationId}/resources`);
}
