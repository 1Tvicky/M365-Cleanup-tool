import { describe, expect, it } from "vitest";
import { computeDataDumpPreview } from "./preview.js";
import { buildDefaultConfig } from "./profiles.js";
import type { DataDumpConfig } from "../../types/dataDump.js";

describe("computeDataDumpPreview", () => {
  it("never mutates or creates anything — pure arithmetic, same result every call (spec: preview must not create real objects)", () => {
    const config = buildDefaultConfig("basic_demo", ["onedrive"]);
    const a = computeDataDumpPreview(1, ["onedrive"], config);
    const b = computeDataDumpPreview(1, ["onedrive"], config);
    expect(a).toEqual(b);
  });

  it("onedrive requestedObjects = folders + files, matching folder count from the same tree builder generation uses", () => {
    const config: DataDumpConfig = {
      namingPrefix: "CF-Test",
      dateRange: { mode: "last_30_days" },
      onedrive: {
        userCount: 2,
        rootFolders: 3,
        subFoldersPerFolder: 4, // irrelevant at maxFolderDepth: 1 — no children regardless
        maxFolderDepth: 1, // no children -> folder count per user == rootFolders
        filesPerFolder: 5,
        targetTotalSizeBytes: 0,
        minFileSizeBytes: 1000,
        maxFileSizeBytes: 2000,
        fileTypeDistribution: { txt: 100 },
        namingStyle: "synthetic",
      },
    };
    const preview = computeDataDumpPreview(1, ["onedrive"], config);
    const workload = preview.workloads[0]!;
    expect(workload.breakdown.folders).toBe(3 * 2); // 3 folders/user * 2 users
    expect(workload.breakdown.files).toBe(3 * 2 * 5); // + 5 files per folder
    expect(workload.requestedObjects).toBe(workload.breakdown.folders! + workload.breakdown.files!);
  });

  it("matches the user's own worked example: 20 root folders + 3 sub-folders each, files generated at every level (not leaves only)", () => {
    const config: DataDumpConfig = {
      namingPrefix: "CF-Test",
      dateRange: { mode: "last_30_days" },
      onedrive: {
        userCount: 1,
        rootFolders: 20,
        subFoldersPerFolder: 3,
        maxFolderDepth: 2,
        filesPerFolder: 3,
        targetTotalSizeBytes: 0,
        minFileSizeBytes: 1000,
        maxFileSizeBytes: 2000,
        fileTypeDistribution: { txt: 100 },
        namingStyle: "synthetic",
      },
    };
    const preview = computeDataDumpPreview(1, ["onedrive"], config);
    const workload = preview.workloads[0]!;
    const expectedFolders = 20 + 20 * 3; // 20 roots + 3 sub-folders under each
    expect(workload.breakdown.folders).toBe(expectedFolders);
    expect(workload.breakdown.files).toBe(expectedFolders * 3); // 3 files in EVERY folder, root included
  });

  it("uses targetTotalSizeBytes as the estimate when it is set, instead of deriving from average file size", () => {
    const config: DataDumpConfig = {
      namingPrefix: "CF-Test",
      dateRange: { mode: "last_30_days" },
      onedrive: {
        userCount: 1,
        rootFolders: 1,
        subFoldersPerFolder: 0,
        maxFolderDepth: 1,
        filesPerFolder: 1,
        targetTotalSizeBytes: 999_999,
        minFileSizeBytes: 1,
        maxFileSizeBytes: 2,
        fileTypeDistribution: { txt: 100 },
        namingStyle: "synthetic",
      },
    };
    const preview = computeDataDumpPreview(1, ["onedrive"], config);
    expect(preview.workloads[0]!.estimatedSizeBytes).toBe(999_999);
  });

  it("teams workload never counts configured messages/replies toward requestedObjects (they are not created)", () => {
    const config: DataDumpConfig = {
      namingPrefix: "CF-Test",
      dateRange: { mode: "last_30_days" },
      teams: { teamCount: 2, membersPerTeam: 3, channelsPerTeam: 4, messagesPerChannel: 1000, repliesPerMessage: 10, namingStyle: "professional" },
    };
    const preview = computeDataDumpPreview(1, ["teams"], config);
    const workload = preview.workloads[0]!;
    expect(workload.requestedObjects).toBe(2 + 2 * 4); // teams + channels only
    expect(workload.warnings.join(" ")).toMatch(/migration-mode/i);
  });

  it("sharepoint workload warns about the Sites.Create.All requirement only when a new site is actually configured", () => {
    const withoutNewSite = buildDefaultConfig("basic_demo", ["sharepoint"]);
    const previewWithout = computeDataDumpPreview(1, ["sharepoint"], withoutNewSite);
    expect(previewWithout.workloads[0]!.warnings.join(" ")).not.toMatch(/Sites\.Create\.All/i);

    const withNewSite = buildDefaultConfig("basic_demo", ["sharepoint"], {
      sharepoint: { ...withoutNewSite.sharepoint!, newSite: { displayName: "New Site", urlSlug: "new-site", template: "communicationSite" } },
    });
    const previewWith = computeDataDumpPreview(1, ["sharepoint"], withNewSite);
    expect(previewWith.workloads[0]!.warnings.join(" ")).toMatch(/Sites\.Create\.All/i);
    expect(previewWith.workloads[0]!.breakdown.newSites).toBe(1);
  });

  it("onedrive uses selectedUserIds.length (explicit selection) over the generic userCount once a real selection exists", () => {
    const base = buildDefaultConfig("basic_demo", ["onedrive"]);
    const config: DataDumpConfig = { ...base, onedrive: { ...base.onedrive!, userCount: 999, selectedUserIds: ["u1", "u2", "u3"] } };
    const preview = computeDataDumpPreview(1, ["onedrive"], config);
    expect(preview.workloads[0]!.breakdown.users).toBe(3);
  });

  it("outlook breakdown includes calendar events and contacts independently of email count", () => {
    const base = buildDefaultConfig("basic_demo", ["outlook"]);
    const config: DataDumpConfig = { ...base, outlook: { ...base.outlook!, userCount: 2, emailsPerUser: 0, calendarEventCount: 5, contactCount: 10 } };
    const preview = computeDataDumpPreview(1, ["outlook"], config);
    const b = preview.workloads[0]!.breakdown;
    expect(b.emails).toBe(0);
    expect(b.calendarEvents).toBe(10); // 2 users * 5 events
    expect(b.contacts).toBe(20); // 2 users * 10 contacts
  });

  it("aggregates totals across multiple selected workloads", () => {
    const config = buildDefaultConfig("basic_demo", ["onedrive", "outlook"]);
    const preview = computeDataDumpPreview(1, ["onedrive", "outlook"], config);
    expect(preview.workloads).toHaveLength(2);
    expect(preview.totalRequestedObjects).toBe(preview.workloads.reduce((sum, w) => sum + w.requestedObjects, 0));
  });
});
