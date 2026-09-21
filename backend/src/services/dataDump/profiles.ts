import type {
  DataDumpConfig,
  DataDumpProfile,
  DataDumpWorkload,
  OneDriveGenConfig,
  OutlookGenConfig,
  SharePointGenConfig,
  TeamsGenConfig,
} from "../../types/dataDump.js";

/**
 * Profile defaults (spec §7/§8): "these are defaults only, not limits" — every value here is a
 * starting point a Custom-profile operator (or a Basic/Migration/Enterprise/Performance operator
 * via Advanced Options) can override upward or downward with no code-enforced ceiling. The only
 * enforcement is the Zod schema in routes/dataDump.ts, which rejects non-positive/non-finite values
 * — never a maximum.
 */

const DEFAULT_FILE_TYPE_DISTRIBUTION = { docx: 25, xlsx: 20, pdf: 20, pptx: 15, jpg: 10, csv: 5, png: 5 };

function onedriveDefaults(profile: DataDumpProfile): OneDriveGenConfig {
  const base: OneDriveGenConfig = {
    userCount: 3,
    rootFolders: 8,
    subFoldersPerFolder: 3,
    maxFolderDepth: 2,
    filesPerFolder: 3,
    targetTotalSizeBytes: 500 * 1024 * 1024,
    minFileSizeBytes: 10 * 1024,
    maxFileSizeBytes: 5 * 1024 * 1024,
    fileTypeDistribution: { ...DEFAULT_FILE_TYPE_DISTRIBUTION },
    namingStyle: "professional",
  };
  switch (profile) {
    case "migration_demo":
      return { ...base, userCount: 8, rootFolders: 10, subFoldersPerFolder: 4, maxFolderDepth: 3, filesPerFolder: 5, targetTotalSizeBytes: 5 * 1024 * 1024 * 1024 };
    case "enterprise_demo":
      return {
        ...base,
        userCount: 25,
        rootFolders: 10,
        subFoldersPerFolder: 4,
        maxFolderDepth: 3,
        filesPerFolder: 8,
        targetTotalSizeBytes: 50 * 1024 * 1024 * 1024,
        maxFileSizeBytes: 50 * 1024 * 1024,
      };
    case "performance_test":
      return {
        ...base,
        userCount: 50,
        rootFolders: 12,
        subFoldersPerFolder: 5,
        maxFolderDepth: 3,
        filesPerFolder: 10,
        targetTotalSizeBytes: 500 * 1024 * 1024 * 1024,
        maxFileSizeBytes: 200 * 1024 * 1024,
      };
    default:
      return base;
  }
}

function sharepointDefaults(profile: DataDumpProfile): SharePointGenConfig {
  const base: SharePointGenConfig = {
    siteCount: 1,
    librariesPerSite: 2,
    foldersPerLibrary: 5,
    subFoldersPerFolder: 3,
    maxFolderDepth: 2,
    filesPerFolder: 3,
    targetTotalSizeBytes: 500 * 1024 * 1024,
    minFileSizeBytes: 10 * 1024,
    maxFileSizeBytes: 5 * 1024 * 1024,
    fileTypeDistribution: { ...DEFAULT_FILE_TYPE_DISTRIBUTION },
    generatePermissions: false,
    namingStyle: "professional",
  };
  switch (profile) {
    case "migration_demo":
      return { ...base, siteCount: 2, librariesPerSite: 3, foldersPerLibrary: 8, subFoldersPerFolder: 4, filesPerFolder: 5, targetTotalSizeBytes: 5 * 1024 * 1024 * 1024, generatePermissions: true };
    case "enterprise_demo":
      return {
        ...base,
        siteCount: 3,
        librariesPerSite: 4,
        foldersPerLibrary: 10,
        subFoldersPerFolder: 4,
        filesPerFolder: 8,
        targetTotalSizeBytes: 50 * 1024 * 1024 * 1024,
        maxFileSizeBytes: 50 * 1024 * 1024,
        generatePermissions: true,
      };
    case "performance_test":
      return {
        ...base,
        siteCount: 5,
        librariesPerSite: 5,
        foldersPerLibrary: 15,
        subFoldersPerFolder: 5,
        filesPerFolder: 10,
        targetTotalSizeBytes: 200 * 1024 * 1024 * 1024,
        maxFileSizeBytes: 200 * 1024 * 1024,
      };
    default:
      return base;
  }
}

function teamsDefaults(profile: DataDumpProfile): TeamsGenConfig {
  const base: TeamsGenConfig = {
    teamCount: 1,
    membersPerTeam: 3,
    channelsPerTeam: 4,
    messagesPerChannel: 10,
    repliesPerMessage: 2,
    namingStyle: "professional",
  };
  switch (profile) {
    case "migration_demo":
      return { ...base, teamCount: 3, membersPerTeam: 6, channelsPerTeam: 4, messagesPerChannel: 25, repliesPerMessage: 3 };
    case "enterprise_demo":
      return { ...base, teamCount: 8, membersPerTeam: 12, channelsPerTeam: 5, messagesPerChannel: 50, repliesPerMessage: 4 };
    case "performance_test":
      return { ...base, teamCount: 15, membersPerTeam: 20, channelsPerTeam: 6, messagesPerChannel: 200, repliesPerMessage: 5 };
    default:
      return base;
  }
}

function outlookDefaults(profile: DataDumpProfile): OutlookGenConfig {
  const base: OutlookGenConfig = {
    userCount: 3,
    emailsPerUser: 30,
    attachmentsPerEmail: 0.3,
    averageAttachmentSizeBytes: 200 * 1024,
    calendarEventCount: 10,
    includeRecurringEvents: false,
    includeAttendees: true,
    contactCount: 20,
    namingStyle: "professional",
  };
  switch (profile) {
    case "migration_demo":
      return { ...base, userCount: 8, emailsPerUser: 150, attachmentsPerEmail: 0.4, calendarEventCount: 40, contactCount: 60, includeRecurringEvents: true };
    case "enterprise_demo":
      return { ...base, userCount: 25, emailsPerUser: 500, attachmentsPerEmail: 0.5, averageAttachmentSizeBytes: 500 * 1024, calendarEventCount: 100, contactCount: 150, includeRecurringEvents: true };
    case "performance_test":
      return { ...base, userCount: 50, emailsPerUser: 5000, attachmentsPerEmail: 0.3, averageAttachmentSizeBytes: 1024 * 1024, calendarEventCount: 500, contactCount: 500 };
    default:
      return base;
  }
}

/** Builds a default DataDumpConfig for the given profile, scoped to only the selected workloads. `overrides` is deep-merged on top (Advanced Options / Custom profile). */
export function buildDefaultConfig(profile: DataDumpProfile, workloads: DataDumpWorkload[], overrides?: Partial<DataDumpConfig>): DataDumpConfig {
  const config: DataDumpConfig = {
    namingPrefix: overrides?.namingPrefix ?? defaultPrefixForProfile(profile),
    seed: overrides?.seed,
    dateRange: overrides?.dateRange ?? { mode: "last_30_days" },
  };
  if (workloads.includes("onedrive")) config.onedrive = { ...onedriveDefaults(profile), ...overrides?.onedrive };
  if (workloads.includes("sharepoint")) config.sharepoint = { ...sharepointDefaults(profile), ...overrides?.sharepoint };
  if (workloads.includes("teams")) config.teams = { ...teamsDefaults(profile), ...overrides?.teams };
  if (workloads.includes("outlook")) config.outlook = { ...outlookDefaults(profile), ...overrides?.outlook };
  return config;
}

function defaultPrefixForProfile(profile: DataDumpProfile): string {
  switch (profile) {
    case "basic_demo":
      return "CF-Demo";
    case "migration_demo":
      return "CF-Migration";
    case "enterprise_demo":
      return "CF-Demo";
    case "performance_test":
      return "CF-Performance";
    default:
      return "CF-Test";
  }
}

export const NAMING_PREFIX_PRESETS = ["CF-Demo", "CF-QA", "CF-Performance", "CF-Migration", "CF-Test"] as const;
