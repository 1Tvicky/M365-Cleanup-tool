import type { DataDumpConfig, DataDumpPreviewResult, DataDumpPreviewWorkload, DataDumpWorkload } from "../../types/dataDump.js";
import { buildFolderTree, countFolderNodes } from "./businessContent.js";
import { createSeededSource, forkSeed } from "./seededRandom.js";

/**
 * Pure preview computation — spec §21/§24: "Preview must not create real objects." No Graph calls,
 * no database writes; everything here is arithmetic + the same folder-tree builder the real
 * generator uses, so the preview's folder count matches what generation will actually produce for
 * the same config/seed. Explicit resource selection (selectedUserIds/selectedSiteIds/
 * selectedTeamIds/newSite/newTeam — the Select Resources step) is reflected exactly: the preview
 * counts precisely who/what was picked, never a generic "userCount" placeholder once a real
 * selection exists.
 */

function averageOf(min: number, max: number): number {
  return Math.max(1, Math.round((min + max) / 2));
}

/** Every root, sub-folder, and nested folder in the uniform tree — filesPerFolder is generated in EVERY one of these, not just leaves, so this count feeds totalFiles = totalFolders * filesPerFolder directly. */
function countFolders(rootFolders: number, subFoldersPerFolder: number, maxFolderDepth: number, seed: number, namingStyle: "professional" | "synthetic"): number {
  const tree = buildFolderTree(createSeededSource(seed), rootFolders, subFoldersPerFolder, maxFolderDepth, namingStyle);
  return countFolderNodes(tree);
}

export function computeDataDumpPreview(operationSeed: number, workloads: DataDumpWorkload[], config: DataDumpConfig): DataDumpPreviewResult {
  const results: DataDumpPreviewWorkload[] = [];

  if (workloads.includes("onedrive") && config.onedrive) {
    const c = config.onedrive;
    const userCount = c.selectedUserIds?.length ?? c.userCount;
    const folderSeed = forkSeed(operationSeed, "preview:onedrive:folders");
    const foldersPerUser = countFolders(c.rootFolders, c.subFoldersPerFolder, c.maxFolderDepth, folderSeed, c.namingStyle);
    const totalFolders = foldersPerUser * userCount;
    const totalFiles = totalFolders * c.filesPerFolder;
    const avgSize = averageOf(c.minFileSizeBytes, c.maxFileSizeBytes);
    const estimatedSizeBytes = c.targetTotalSizeBytes > 0 ? c.targetTotalSizeBytes : totalFiles * avgSize;
    results.push({
      workload: "onedrive",
      requestedObjects: totalFolders + totalFiles,
      estimatedSizeBytes,
      breakdown: { users: userCount, folders: totalFolders, files: totalFiles },
      warnings: [],
    });
  }

  if (workloads.includes("sharepoint") && config.sharepoint) {
    const c = config.sharepoint;
    const existingSiteCount = c.selectedSiteIds?.length ?? c.siteCount;
    const newSiteCount = c.newSite ? 1 : 0;
    const siteCount = existingSiteCount + newSiteCount;
    const folderSeed = forkSeed(operationSeed, "preview:sharepoint:folders");
    const foldersPerLibrary = countFolders(c.foldersPerLibrary, c.subFoldersPerFolder, c.maxFolderDepth, folderSeed, c.namingStyle);
    const totalLibraries = c.librariesPerSite * siteCount;
    const totalFolders = foldersPerLibrary * totalLibraries;
    const totalFiles = totalFolders * c.filesPerFolder;
    const avgSize = averageOf(c.minFileSizeBytes, c.maxFileSizeBytes);
    const estimatedSizeBytes = c.targetTotalSizeBytes > 0 ? c.targetTotalSizeBytes : totalFiles * avgSize;
    const warnings: string[] = [];
    if (newSiteCount > 0) {
      warnings.push(
        "Creating a new SharePoint site requires the Microsoft Graph beta API and the Sites.Create.All application permission. If this app's Azure AD registration hasn't been granted that permission, site creation will fail with a clear authorization error rather than silently succeeding — see docs/azure-ad-app-registration.md."
      );
    }
    if (c.generatePermissions) warnings.push("Permission assignment is best-effort: a failure to grant a specific role is reported per-user, not treated as a fatal error for the rest of the workload.");
    results.push({
      workload: "sharepoint",
      // Only the (at most one) new site is a "create" attempt — existing selected sites are
      // targeted, not created, same reasoning as Teams' existingTeams/newTeams split below.
      requestedObjects: newSiteCount + totalLibraries + totalFolders + totalFiles,
      estimatedSizeBytes,
      breakdown: { existingSites: existingSiteCount, newSites: newSiteCount, libraries: totalLibraries, folders: totalFolders, files: totalFiles },
      warnings,
    });
  }

  if (workloads.includes("teams") && config.teams) {
    const c = config.teams;
    const existingTeamCount = c.selectedTeamIds?.length ?? 0;
    const newTeamCount = c.newTeam ? 1 : 0;
    const usingExplicit = existingTeamCount > 0 || newTeamCount > 0;
    // Legacy quick-start (no manual selection): teamCount brand-new teams are created, same as
    // before explicit selection existed. Once a selection/newTeam exists, only the (at most one)
    // newTeam is ever "created" here — existing selected teams are targeted, not created.
    const teamCount = usingExplicit ? existingTeamCount + newTeamCount : c.teamCount;
    const createdTeamCount = usingExplicit ? newTeamCount : c.teamCount;
    const channelsPerTeam = c.channels && c.channels.length > 0 ? c.channels.length : c.channelsPerTeam;
    const totalChannels = teamCount * channelsPerTeam;
    const totalMessages = c.channels && c.channels.length > 0 ? teamCount * c.channels.reduce((s, ch) => s + ch.messages, 0) : totalChannels * c.messagesPerChannel;
    const totalReplies = c.channels && c.channels.length > 0 ? teamCount * c.channels.reduce((s, ch) => s + ch.messages * ch.replies, 0) : totalMessages * c.repliesPerMessage;
    results.push({
      workload: "teams",
      requestedObjects: createdTeamCount + totalChannels,
      estimatedSizeBytes: 0,
      breakdown: { existingTeams: existingTeamCount, newTeams: newTeamCount, teams: usingExplicit ? existingTeamCount + newTeamCount : c.teamCount, channels: totalChannels, configuredMessages: totalMessages, configuredReplies: totalReplies },
      warnings: [
        "Channel messages/replies are configured here for planning purposes but are NOT created: Microsoft Graph only allows application-permission channel-message posting in a special migration-mode team state, not for normal teams. Teams and channels themselves ARE created for real.",
      ],
    });
  }

  if (workloads.includes("outlook") && config.outlook) {
    const c = config.outlook;
    const userCount = c.selectedUserIds?.length ?? c.userCount;
    const totalEmails = userCount * c.emailsPerUser;
    const totalAttachments = Math.round(totalEmails * c.attachmentsPerEmail);
    const totalEvents = userCount * c.calendarEventCount;
    const totalContacts = userCount * c.contactCount;
    const estimatedSizeBytes = totalAttachments * c.averageAttachmentSizeBytes;
    const warnings: string[] = [];
    if (c.emailsPerUser > 0) {
      warnings.push(
        "Historical received/sent timestamps on generated mail are a Microsoft Graph v1.0 platform limitation: custom receivedDateTime/sentDateTime is only available on the unsupported beta endpoint. Generated mail is timestamped at creation time."
      );
    }
    results.push({
      workload: "outlook",
      // Attachments are deliberately NOT added into requestedObjects: an attachment isn't its own
      // independently-tracked create attempt the way an email/event/contact is — it's a probabilistic
      // sub-property of one email's creation (services/dataDump/runners/outlookRunner.ts folds a
      // failed attachment into that email's own outcome). Same treatment SharePoint gives permission
      // grants (informational breakdown figure, never counted toward the total) — keeps requested =
      // created + failed + skipped reconciling exactly against what the runner actually accounts for.
      requestedObjects: totalEmails + totalEvents + totalContacts,
      estimatedSizeBytes,
      breakdown: { users: userCount, emails: totalEmails, attachments: totalAttachments, calendarEvents: totalEvents, contacts: totalContacts },
      warnings,
    });
  }

  return {
    workloads: results,
    totalRequestedObjects: results.reduce((sum, w) => sum + w.requestedObjects, 0),
    totalEstimatedSizeBytes: results.reduce((sum, w) => sum + w.estimatedSizeBytes, 0),
  };
}
