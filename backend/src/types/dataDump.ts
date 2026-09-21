/**
 * Data Dump module types. Deliberately its own file, its own table set (see
 * db/migrations/017_data_dump.sql), and its own operation/status vocabulary — never shares a type
 * with Cleaning/Cleanup (types/cleaning.ts) even where the shape rhymes (both have an "operation"
 * with items/progress/status), so the two business functions can never be confused for one another
 * at the type-checker level either.
 */

export type DataDumpProfile = "basic_demo" | "migration_demo" | "enterprise_demo" | "performance_test" | "custom";
export const DATA_DUMP_PROFILES: readonly DataDumpProfile[] = [
  "basic_demo",
  "migration_demo",
  "enterprise_demo",
  "performance_test",
  "custom",
];

export type DataDumpWorkload = "onedrive" | "sharepoint" | "teams" | "outlook";
export const DATA_DUMP_WORKLOADS: readonly DataDumpWorkload[] = ["onedrive", "sharepoint", "teams", "outlook"];

export type DataDumpOperationStatus = "queued" | "running" | "paused" | "completed" | "completed_with_errors" | "failed" | "cancelled";
export type DataDumpWorkloadTaskStatus = "pending" | "running" | "paused" | "completed" | "completed_with_errors" | "failed" | "cancelled" | "skipped";

export type NamingStyle = "professional" | "synthetic";
export type FileTypeKey = "docx" | "xlsx" | "pptx" | "pdf" | "txt" | "csv" | "jpg" | "png" | "zip";
export const FILE_TYPE_KEYS: readonly FileTypeKey[] = ["docx", "xlsx", "pptx", "pdf", "txt", "csv", "jpg", "png", "zip"];

/** Percentage weights keyed by FileTypeKey — need not include every key; missing keys are treated as 0. Need not sum to exactly 100 (normalized at generation time), but must be non-negative. */
export type FileTypeDistribution = Partial<Record<FileTypeKey, number>>;

export type DateRangeMode = "last_30_days" | "last_6_months" | "last_1_year" | "last_3_years" | "custom";

export interface DateRangeConfig {
  mode: DateRangeMode;
  /** Required, ISO date strings, only when mode === "custom". */
  customStartDate?: string;
  customEndDate?: string;
}

export interface OneDriveGenConfig {
  /** Explicit resource selection from the Select Users step (DiscoveryTable) — when present, this is exactly who gets generated content, no more no less. Falls back to userCount below only for the profile-driven quick-start path that has no selection step (Basic/Migration/Enterprise/Performance without manual picking). */
  selectedUserIds?: string[];
  userCount: number;
  /** Total number of root folders created directly under the user's OneDrive. */
  rootFolders: number;
  /** How many sub-folders each folder gets, uniformly, at every level below the root — e.g. 20 root folders + 3 here means every one of those 20 gets exactly 3 children (see buildFolderTree). */
  subFoldersPerFolder: number;
  /** How many levels deep the uniform tree goes; 1 means root folders only, no sub-folders at all regardless of subFoldersPerFolder. */
  maxFolderDepth: number;
  /** Applied at EVERY folder in the tree — root, sub-folder, and nested folder alike — not leaf folders only. */
  filesPerFolder: number;
  targetTotalSizeBytes: number;
  minFileSizeBytes: number;
  maxFileSizeBytes: number;
  fileTypeDistribution: FileTypeDistribution;
  namingStyle: NamingStyle;
}

export interface NewSiteConfig {
  displayName: string;
  /** SharePoint URL alias — only meaningful for a Communication site's own URL segment; Graph derives the rest. */
  urlSlug: string;
  /** Graph's `/beta/sites` template values — see graph/dataDump/sharePointCreation.ts's createSharePointSite; never a value that endpoint doesn't document. */
  template: "teamSiteWithoutMicrosoft365Group" | "communicationSite";
  description?: string;
  ownerUserIds?: string[];
}

export interface SharePointGenConfig {
  /** Existing, already-connected sites selected from the Select Sites step. */
  selectedSiteIds?: string[];
  /** When set, a brand-new site is created first (see graph/dataDump/sharePointCreation.ts — requires the beta API + Sites.Create.All, which this app may or may not have consent for; see docs/azure-ad-app-registration.md) and generation continues into it. selectedSiteIds and newSite may both be set — an operation can target existing sites AND create one new site in the same run. */
  newSite?: NewSiteConfig;
  siteCount: number;
  librariesPerSite: number;
  /** Total number of root folders created directly under each document library. */
  foldersPerLibrary: number;
  /** Same uniform-tree semantics as OneDriveGenConfig.subFoldersPerFolder. */
  subFoldersPerFolder: number;
  maxFolderDepth: number;
  /** Applied at EVERY folder in the tree — root, sub-folder, and nested folder alike — not leaf folders only. */
  filesPerFolder: number;
  targetTotalSizeBytes: number;
  minFileSizeBytes: number;
  maxFileSizeBytes: number;
  fileTypeDistribution: FileTypeDistribution;
  generatePermissions: boolean;
  namingStyle: NamingStyle;
}

export interface ChannelConfig {
  name: string;
  /** Explicit channel membership — need not equal the team's own membership (spec: "do not assume Team membership and channel membership are always identical"). Empty/omitted means every team member. */
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
  /** Existing teams selected from the Select Teams step — generation adds the configured channels (and, best-effort, members) to each of these. */
  selectedTeamIds?: string[];
  /** When set, a brand-new Team is created first (graph/dataDump/teamsCreation.ts) and generation continues into it. */
  newTeam?: NewTeamConfig;
  /** Per-channel configuration — replaces the old flat channelsPerTeam/messagesPerChannel pair once a selection/creation flow is in play. Falls back to the legacy count-based fields below only for the profile-driven quick-start path. */
  channels?: ChannelConfig[];
  teamCount: number;
  membersPerTeam: number;
  channelsPerTeam: number;
  /** Configured/previewed, but never executed — see docs/data-dump-api.md "Known limitations": app-only channel message posting is a Microsoft Graph platform restriction, not an app-imposed one. */
  messagesPerChannel: number;
  repliesPerMessage: number;
  namingStyle: NamingStyle;
}

export interface OutlookGenConfig {
  selectedUserIds?: string[];
  userCount: number;
  emailsPerUser: number;
  attachmentsPerEmail: number;
  averageAttachmentSizeBytes: number;
  /** Independent from email generation (spec §21/§22) — 0 disables that sub-generator entirely. */
  calendarEventCount: number;
  includeRecurringEvents: boolean;
  includeAttendees: boolean;
  contactCount: number;
  namingStyle: NamingStyle;
}

export interface DataDumpConfig {
  namingPrefix: string;
  /** Optional deterministic seed — same config + seed reproduces the same names/content pattern (services/dataDump/seededRandom.ts). */
  seed?: number;
  dateRange: DateRangeConfig;
  onedrive?: OneDriveGenConfig;
  sharepoint?: SharePointGenConfig;
  teams?: TeamsGenConfig;
  outlook?: OutlookGenConfig;
}

export interface DataDumpOperationRow {
  id: string;
  tenantId: string;
  requestedBy: string;
  label: string;
  profile: DataDumpProfile;
  workloads: DataDumpWorkload[];
  config: DataDumpConfig;
  status: DataDumpOperationStatus;
  requestedItems: number;
  createdItems: number;
  failedItems: number;
  skippedItems: number;
  totalSizeBytes: number;
  parentOperationId: string | null;
  pauseRequestedAt: string | null;
  cancelRequestedAt: string | null;
  startedAt: string | null;
  completedAt: string | null;
  errorMessage: string | null;
  createdAt: string;
}

/**
 * Per-resource-kind accounting (spec §27: "requested = created + failed + skipped" for EVERY kind,
 * not just a workload's overall total) — a OneDrive task's folders and files each reconcile
 * independently, same for a Teams task's teams/channels/members/messages/replies, etc.
 */
export type ResourceSubKind =
  | "user"
  | "folder"
  | "file"
  | "site"
  | "library"
  | "team"
  | "channel"
  | "channel_member"
  | "message"
  | "reply"
  | "email"
  | "attachment"
  | "calendar_event"
  | "contact";

export interface SubcountEntry {
  requested: number;
  created: number;
  failed: number;
  skipped: number;
}

export type DataDumpSubcounts = Partial<Record<ResourceSubKind, SubcountEntry>>;

export interface DataDumpWorkloadTaskRow {
  id: string;
  operationId: string;
  connectionId: string;
  workload: DataDumpWorkload;
  status: DataDumpWorkloadTaskStatus;
  requestedItems: number;
  createdItems: number;
  failedItems: number;
  skippedItems: number;
  totalSizeBytes: number;
  subcounts: DataDumpSubcounts;
  checkpoint: Record<string, unknown>;
  errorMessage: string | null;
  startedAt: string | null;
  completedAt: string | null;
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
