import type { SharePointGenConfig } from "../../../types/dataDump.js";
import { createDocumentLibrary, createSharePointFolder, createSharePointSite, grantSitePermission, uploadSharePointFile } from "../../../graph/dataDump/sharePointCreation.js";
import { runThrottled } from "../../rateLimiter.js";
import { buildFolderTree, countFolderNodes, type FolderNode } from "../businessContent.js";
import { createSeededSource, forkSeed, pickWeighted } from "../seededRandom.js";
import { generateFileContent } from "../fileContent.js";
import { pickFileSizeBytes } from "../sizing.js";
import { resolveDateRange, randomTimestampInRange } from "../dateRange.js";
import { getOrCreateContainer } from "../containers.js";
import { BatchAccumulator, nextBatchIndexFor } from "../batching.js";
import { applyTaskProgress, bumpSubcount, setTaskCheckpoint } from "../progress.js";
import { FILE_TYPE_KEYS } from "../../../types/dataDump.js";
import type { RunnerContext, RunnerOutcome } from "./oneDriveRunner.js";

export interface SiteTarget {
  id: string;
  displayName: string;
}

interface ContainerRef {
  id: string;
  graphId: string;
}

interface FairShareTally {
  folders: (n: number) => void;
  files: (n: number) => void;
}

/**
 * Generates real SharePoint document libraries + folders + files. Supports both target shapes from
 * the Select Sites / Create New Site step (spec §6): `sites` are existing, already-connected sites
 * to generate into as-is; `config.newSite`, when set, creates a brand-new site first (see
 * graph/dataDump/sharePointCreation.ts — requires the beta API + Sites.Create.All, reported clearly
 * if not granted rather than faked) and folds it into the same run. Each library's folder tree is the
 * same simple, uniform shape as oneDriveRunner.ts's: `foldersPerLibrary` root folders, each
 * recursively given `subFoldersPerFolder` children down to `maxFolderDepth` levels, with
 * `filesPerFolder` files generated in EVERY one of those folders (root, sub-folder, and nested
 * folder alike). Same resumability/memory-bounding shape as oneDriveRunner.ts.
 */
export async function runSharePointWorkload(
  ctx: RunnerContext,
  sites: SiteTarget[],
  tenantUsers: { id: string; upn: string; displayName: string }[],
  config: SharePointGenConfig
): Promise<RunnerOutcome> {
  const dateRangeWindow = resolveDateRange(ctx.dateRange);
  let hadFailures = false;
  let globalFileIndex = 0;
  let batchIndex = await nextBatchIndexFor(ctx.taskId);
  const targetSites = [...sites];

  // Mirrors preview.ts's own per-library folder/file math exactly, so a library (or an entire site)
  // that fails before its folders/files are ever attempted still has that full prospective share
  // credited as "failed" instead of leaving Requested permanently ahead of Created + Failed +
  // Skipped — same reasoning as oneDriveRunner.ts's foldersPerUser/filesPerUser.
  const previewFolderSeed = forkSeed(ctx.operationSeed, "preview:sharepoint:folders");
  const foldersPerLibrary = countFolderNodes(
    buildFolderTree(createSeededSource(previewFolderSeed), config.foldersPerLibrary, config.subFoldersPerFolder, config.maxFolderDepth, config.namingStyle)
  );
  const filesPerLibrary = foldersPerLibrary * config.filesPerFolder;
  const foldersPerSite = config.librariesPerSite * foldersPerLibrary;
  const filesPerSite = config.librariesPerSite * filesPerLibrary;

  if (config.newSite) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";
    const owner = tenantUsers[0];
    await bumpSubcount(ctx.taskId, "site", { requested: 1 });
    if (!owner) {
      hadFailures = true;
      await bumpSubcount(ctx.taskId, "site", { failed: 1 });
      await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
    } else {
      try {
        const site = await createSharePointSite(ctx.client, {
          displayName: config.newSite.displayName,
          urlSlug: config.newSite.urlSlug,
          template: config.newSite.template === "communicationSite" ? "communicationSite" : "teamSiteWithoutMicrosoft365Group",
          description: config.newSite.description,
          ownerEmail: owner.upn,
        });
        await getOrCreateContainer(ctx.taskId, "site", site.displayName, null, async () => ({ graphId: site.siteId }));
        targetSites.push({ id: site.siteId, displayName: site.displayName });
        await bumpSubcount(ctx.taskId, "site", { created: 1 });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 1, failed: 0, skipped: 0, sizeBytes: 0 });
      } catch (err) {
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "site", { failed: 1 });
        const message = err instanceof Error ? err.message : String(err);
        console.error(`[data-dump] SharePoint: new site ${config.newSite.displayName} failed`, message);
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
      }
    }
  }

  /** Uploads config.filesPerFolder files directly into one folder container. Returns how many were attempted (created + failed). */
  async function generateFilesInFolder(site: SiteTarget, driveId: string, folder: ContainerRef): Promise<number> {
    if (config.filesPerFolder <= 0) return 0;

    const fileDescriptors = Array.from({ length: config.filesPerFolder }, () => globalFileIndex++);
    const settled: { ok: boolean; name: string; graphId: string | null; sizeBytes: number }[] = [];

    await runThrottled(
      fileDescriptors,
      async (fileIndex) => {
        const fileSeed = createSeededSource(forkSeed(ctx.operationSeed, `sharepoint:file:${fileIndex}`));
        const fileType = pickWeighted(fileSeed.rng, config.fileTypeDistribution, FILE_TYPE_KEYS);
        const generated = await generateFileContent(ctx.operationSeed, fileIndex, fileType, config.namingStyle);
        pickFileSizeBytes(fileSeed.rng, config.minFileSizeBytes, config.maxFileSizeBytes, Math.round((config.minFileSizeBytes + config.maxFileSizeBytes) / 2));
        const timestamp = randomTimestampInRange(fileSeed.rng, dateRangeWindow);
        return uploadSharePointFile(ctx.client, site.id, driveId, folder.graphId, generated.fileName, generated.buffer, generated.mimeType, {
          createdDateTime: timestamp,
          lastModifiedDateTime: timestamp,
        });
      },
      {
        label: "DataDump-SharePoint",
        batchSize: 8,
        isCancelled: ctx.isCancelled,
        onItemSettled: (_fileIndex, result) => {
          if (result.ok) settled.push({ ok: true, name: result.value.name, graphId: result.value.id, sizeBytes: result.value.sizeBytes });
          else settled.push({ ok: false, name: "upload-failed", graphId: null, sizeBytes: 0 });
        },
      }
    );

    const accumulator = new BatchAccumulator(ctx.taskId, folder.id, batchIndex);
    let createdCount = 0;
    let failedCount = 0;
    let sizeBytesSum = 0;
    for (const item of settled) {
      await accumulator.add({ name: item.name, graphId: item.graphId, sizeBytes: item.sizeBytes, status: item.ok ? "created" : "failed" });
      if (item.ok) {
        createdCount++;
        sizeBytesSum += item.sizeBytes;
      } else {
        failedCount++;
        hadFailures = true;
      }
    }
    await accumulator.flush();
    batchIndex = accumulator.currentBatchIndex;
    await bumpSubcount(ctx.taskId, "file", { requested: settled.length, created: createdCount, failed: failedCount });
    await applyTaskProgress(ctx.taskId, ctx.operationId, { created: createdCount, failed: failedCount, skipped: 0, sizeBytes: sizeBytesSum });
    await setTaskCheckpoint(ctx.taskId, { phase: "files", nextBatchIndex: batchIndex, lastSiteId: site.id });

    return settled.length;
  }

  /**
   * Creates one folder node (at ANY depth) and its files, then recurses into its children — same
   * one-function-walks-the-whole-tree shape as oneDriveRunner.ts's processFolderNode, INCLUDING its
   * per-node folder-creation isolation: a deep/wide tree makes a single folder create realistically
   * likely to exhaust Graph's retries, and without isolating it, that one failure used to abort this
   * site's ENTIRE remaining tree rather than just this node's own subtree.
   */
  async function processFolderNode(site: SiteTarget, driveId: string, node: FolderNode, parent: ContainerRef, tally: FairShareTally): Promise<"cancelled" | "paused" | undefined> {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    let folderContainer: ContainerRef;
    try {
      folderContainer = await getOrCreateContainer(ctx.taskId, "folder", node.name, parent.id, async () => {
        const folder = await createSharePointFolder(ctx.client, site.id, driveId, parent.graphId, node.name);
        return { graphId: folder.id };
      });
    } catch (err) {
      // Every descendant of this node needs THIS folder as its parent, so if this folder itself
      // can't be created, its whole subtree (itself + every descendant) can never be created either
      // — credit that whole subtree as failed and move on to this node's siblings.
      const subtreeFolders = countFolderNodes([node]);
      const subtreeFiles = subtreeFolders * config.filesPerFolder;
      hadFailures = true;
      const message = err instanceof Error ? err.message : String(err);
      console.error(`[data-dump] SharePoint: folder "${node.name}" failed`, message);
      await bumpSubcount(ctx.taskId, "folder", { requested: subtreeFolders, created: 0, failed: subtreeFolders });
      if (subtreeFiles > 0) await bumpSubcount(ctx.taskId, "file", { requested: subtreeFiles, created: 0, failed: subtreeFiles });
      await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: subtreeFolders + subtreeFiles, skipped: 0, sizeBytes: 0 });
      tally.folders(subtreeFolders);
      tally.files(subtreeFiles);
      return undefined;
    }

    await bumpSubcount(ctx.taskId, "folder", { requested: 1, created: 1 });
    await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 1, failed: 0, skipped: 0, sizeBytes: 0 });
    tally.folders(1);

    const filesAttempted = await generateFilesInFolder(site, driveId, folderContainer);
    tally.files(filesAttempted);

    for (const child of node.children) {
      const signal = await processFolderNode(site, driveId, child, folderContainer, tally);
      if (signal) return signal;
    }
    return undefined;
  }

  /**
   * Isolated per-site, same reasoning as oneDriveRunner.ts's per-user isolation: one site's failure
   * (e.g. a library-creation permission gap on that specific site) must not abort every other site
   * in the same run. `tally` reports back exactly how many of this site's folders/files were fully
   * accounted for (created or failed) before any exception, mirroring oneDriveRunner.ts.
   */
  async function processSite(site: SiteTarget, tally: FairShareTally): Promise<"cancelled" | "paused" | undefined> {
    // Every targeted site gets its own container record — for an EXISTING site this is pure local
    // bookkeeping (the site already exists, nothing is created in M365), and for the just-created
    // new site above this reuses that same record (same kind/displayName/null-parent, found by
    // getOrCreateContainer's own lookup) rather than inserting a duplicate. Without this, libraries
    // were being recorded with `parentContainerId: null` and a name scoped only by index ("Library
    // 1"), which collided across DIFFERENT sites the same way OneDrive's per-user root once did —
    // real bug, fixed alongside that one.
    const siteContainer = await getOrCreateContainer(ctx.taskId, "site", site.displayName, null, async () => ({ graphId: site.id }));

    for (let libIndex = 0; libIndex < config.librariesPerSite; libIndex++) {
      if (await ctx.isCancelled()) return "cancelled";
      if (await ctx.isPaused()) return "paused";

      const libraryName = `${ctx.namingPrefix} Library ${libIndex + 1}`;
      await bumpSubcount(ctx.taskId, "library", { requested: 1 });
      let libraryContainer;
      try {
        libraryContainer = await getOrCreateContainer(ctx.taskId, "site_library", libraryName, siteContainer.id, async () => {
          const library = await createDocumentLibrary(ctx.client, site.id, libraryName);
          return { graphId: library.driveId };
        });
        await bumpSubcount(ctx.taskId, "library", { created: 1 });
      } catch (err) {
        // This library's own folders/files never get a chance to run — credit that whole share as
        // failed too (not just the library itself), or Requested would permanently outpace Created +
        // Failed + Skipped by exactly this library's foldersPerLibrary/filesPerLibrary share.
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "library", { failed: 1 });
        await bumpSubcount(ctx.taskId, "folder", { requested: foldersPerLibrary, created: 0, failed: foldersPerLibrary });
        await bumpSubcount(ctx.taskId, "file", { requested: filesPerLibrary, created: 0, failed: filesPerLibrary });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1 + foldersPerLibrary + filesPerLibrary, skipped: 0, sizeBytes: 0 });
        tally.folders(foldersPerLibrary);
        tally.files(filesPerLibrary);
        continue;
      }
      const driveId = libraryContainer.graphId;

      if (config.generatePermissions && tenantUsers.length > 0) {
        const rolePlan: ("owner" | "write" | "read")[] = ["owner", "write", "read"];
        for (let i = 0; i < Math.min(3, tenantUsers.length); i++) {
          await grantSitePermission(ctx.client, site.id, tenantUsers[i]!.upn, rolePlan[i]!);
        }
      }

      const folderSeed = forkSeed(ctx.operationSeed, `sharepoint:${site.id}:${libIndex}:folders`);
      const tree = buildFolderTree(createSeededSource(folderSeed), config.foldersPerLibrary, config.subFoldersPerFolder, config.maxFolderDepth, config.namingStyle);

      for (const root of tree) {
        const signal = await processFolderNode(site, driveId, root, libraryContainer, tally);
        if (signal) return signal;
      }
    }
    return undefined;
  }

  for (const site of targetSites) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    let accountedFolders = 0;
    let accountedFiles = 0;
    try {
      const signal = await processSite(site, {
        folders: (n) => (accountedFolders += n),
        files: (n) => (accountedFiles += n),
      });
      if (signal) return signal;
    } catch (err) {
      // Whatever share of this site's folders/files never got a chance to run is credited as failed
      // here — same reasoning as oneDriveRunner.ts's own outer per-user catch.
      hadFailures = true;
      const message = err instanceof Error ? err.message : String(err);
      console.error(`[data-dump] SharePoint: site ${site.displayName} failed`, message);
      const remainingFolders = Math.max(foldersPerSite - accountedFolders, 0);
      const remainingFiles = Math.max(filesPerSite - accountedFiles, 0);
      if (remainingFolders > 0) await bumpSubcount(ctx.taskId, "folder", { requested: remainingFolders, created: 0, failed: remainingFolders });
      if (remainingFiles > 0) await bumpSubcount(ctx.taskId, "file", { requested: remainingFiles, created: 0, failed: remainingFiles });
      await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: remainingFolders + remainingFiles, skipped: 0, sizeBytes: 0 });
    }
  }

  return hadFailures ? "completed_with_errors" : "completed";
}
