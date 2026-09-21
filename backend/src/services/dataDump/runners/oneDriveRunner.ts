import type { Client } from "@microsoft/microsoft-graph-client";
import type { OneDriveGenConfig } from "../../../types/dataDump.js";
import { createOneDriveFolder, uploadOneDriveFile } from "../../../graph/dataDump/oneDriveCreation.js";
import { runThrottled } from "../../rateLimiter.js";
import { buildFolderTree, countFolderNodes, type FolderNode } from "../businessContent.js";
import { createSeededSource, forkSeed, pickWeighted } from "../seededRandom.js";
import { generateFileContent } from "../fileContent.js";
import { pickFileSizeBytes } from "../sizing.js";
import { resolveDateRange, randomTimestampInRange } from "../dateRange.js";
import { getOrCreateContainer } from "../containers.js";
import { BatchAccumulator, nextBatchIndexFor, type BatchItemSummary } from "../batching.js";
import { applyTaskProgress, bumpSubcount, setTaskCheckpoint } from "../progress.js";
import { FILE_TYPE_KEYS } from "../../../types/dataDump.js";
import type { DateRangeConfig } from "../../../types/dataDump.js";

export interface RunnerContext {
  client: Client;
  operationId: string;
  taskId: string;
  operationSeed: number;
  namingPrefix: string;
  dateRange: DateRangeConfig;
  isCancelled: () => Promise<boolean>;
  isPaused: () => Promise<boolean>;
}

export type RunnerOutcome = "completed" | "completed_with_errors" | "paused" | "cancelled";

interface ContainerRef {
  id: string;
  graphId: string;
}

/** Reports back how many of a user's/site's fair-share folders and files were actually accounted for (created or failed) before any exception, so an outer catch can credit the untouched remainder as failed without double-counting — see the per-user loop at the bottom of this file. */
interface FairShareTally {
  folders: (n: number) => void;
  files: (n: number) => void;
}

/**
 * Generates real OneDrive folders + files for `users` (existing tenant users — spec never creates
 * new M365 users). The folder tree is a simple, uniform shape the operator controls directly:
 * `rootFolders` folders at the top, each recursively given `subFoldersPerFolder` children down to
 * `maxFolderDepth` levels deep — and `filesPerFolder` files are generated in EVERY one of those
 * folders (root, sub-folder, and nested folder alike), not leaf folders only.
 *
 * Resumable at folder-batch granularity: containers (every folder at every depth) are looked up in
 * data_dump_containers before creation, and file batches resume from nextBatchIndexFor — a restart
 * never recreates an already-recorded container, and at most re-does one partial batch of files,
 * never the whole workload.
 *
 * Memory-bounded (spec §23): only ONE folder's worth of file descriptors (lightweight {index}
 * objects, not file content) is ever materialized at a time — the tree is walked node by node, never
 * flattened into one array of every file in the whole operation.
 */
export async function runOneDriveWorkload(ctx: RunnerContext, users: { id: string; upn: string; displayName: string }[], config: OneDriveGenConfig): Promise<RunnerOutcome> {
  const dateRangeWindow = resolveDateRange(ctx.dateRange);
  let hadFailures = false;
  let globalFileIndex = 0;
  let batchIndex = await nextBatchIndexFor(ctx.taskId);

  // Mirrors preview.ts's own per-user folder/file math exactly (same seed, same formula) so that a
  // user who fails outright — e.g. before a single folder is created — still has their FULL
  // prospective share of "requested" credited as "failed" rather than left permanently unaccounted
  // for. Without this, one user with no provisioned OneDrive ("mysite not found") would leave the
  // operation's own Requested = Created + Failed + Skipped invariant (spec §27) broken forever: the
  // preview already counted this user's folders/files toward Requested, so something must eventually
  // account for them.
  const previewFolderSeed = forkSeed(ctx.operationSeed, "preview:onedrive:folders");
  const foldersPerUser = countFolderNodes(
    buildFolderTree(createSeededSource(previewFolderSeed), config.rootFolders, config.subFoldersPerFolder, config.maxFolderDepth, config.namingStyle)
  );
  const filesPerUser = foldersPerUser * config.filesPerFolder;

  /** Uploads config.filesPerFolder files directly into one folder container. Returns how many were attempted (created + failed), which IS the number now accounted for toward that folder's/user's fair share. */
  async function generateFilesInFolder(userId: string, folder: ContainerRef): Promise<number> {
    if (config.filesPerFolder <= 0) return 0;

    const fileDescriptors = Array.from({ length: config.filesPerFolder }, () => globalFileIndex++);
    const settled: { ok: boolean; name: string; graphId: string | null; sizeBytes: number }[] = [];

    await runThrottled(
      fileDescriptors,
      async (fileIndex) => {
        const fileSeed = createSeededSource(forkSeed(ctx.operationSeed, `onedrive:file:${fileIndex}`));
        const fileType = pickWeighted(fileSeed.rng, config.fileTypeDistribution, FILE_TYPE_KEYS);
        const generated = await generateFileContent(ctx.operationSeed, fileIndex, fileType, config.namingStyle);
        const targetSize = pickFileSizeBytes(fileSeed.rng, config.minFileSizeBytes, config.maxFileSizeBytes, Math.round((config.minFileSizeBytes + config.maxFileSizeBytes) / 2));
        void targetSize; // structured file formats keep their natural generated size — see fileContent.ts's padToSize note; target size guides file COUNT (services/dataDump/sizing.ts), not per-file padding, for these types
        const timestamp = randomTimestampInRange(fileSeed.rng, dateRangeWindow);
        return uploadOneDriveFile(ctx.client, userId, folder.graphId, generated.fileName, generated.buffer, generated.mimeType, {
          createdDateTime: timestamp,
          lastModifiedDateTime: timestamp,
        });
      },
      {
        label: "DataDump-OneDrive",
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
      const summary: BatchItemSummary = { name: item.name, graphId: item.graphId, sizeBytes: item.sizeBytes, status: item.ok ? "created" : "failed" };
      await accumulator.add(summary);
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
    await setTaskCheckpoint(ctx.taskId, { phase: "files", nextBatchIndex: batchIndex, lastUserId: userId });

    return settled.length;
  }

  /**
   * Creates one folder node (at ANY depth) and its files, then recurses into its children — the
   * whole tree, root through the deepest nested folder, is walked by this one function so every
   * node is treated identically: create the folder, generate filesPerFolder files inside it, then
   * do the same for each child. `node.name` already doubles as its lookup key AND display name
   * (services/dataDump/containers.ts's getOrCreateContainer matches on kind+displayName+parent) —
   * safe here because businessContent.ts's buildFolderTree guarantees every node in a tree has a
   * unique name, so two different nodes never collide under the same parent.
   *
   * The folder-creation Graph call is isolated in its own try/catch, deliberately separate from
   * file generation and recursion: a deep, wide tree (e.g. depth 6 × 2 sub-folders = 63 folders
   * under just one root) makes it realistic for a SINGLE folder create to exhaust Graph's own
   * SDK-level retry (throttling, a transient 5xx, ...) — without this isolation, that one failure
   * used to propagate all the way up through every recursive call and abort the user's ENTIRE
   * remaining tree, not just this node's own branch (caught live: a 630-folder/1890-file run where
   * one early failure left only 22 folders/62 files created before the rest was bulk-marked
   * failed). Now it aborts only this node's own subtree — its siblings, and every OTHER root, still
   * get a fair chance to run.
   */
  async function processFolderNode(userId: string, node: FolderNode, parent: ContainerRef, tally: FairShareTally): Promise<"cancelled" | "paused" | undefined> {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    let folderContainer: ContainerRef;
    try {
      folderContainer = await getOrCreateContainer(ctx.taskId, "folder", node.name, parent.id, async () => {
        const folder = await createOneDriveFolder(ctx.client, userId, parent.graphId, node.name);
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
      console.error(`[data-dump] OneDrive: folder "${node.name}" failed`, message);
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

    const filesAttempted = await generateFilesInFolder(userId, folderContainer);
    tally.files(filesAttempted);

    for (const child of node.children) {
      const signal = await processFolderNode(userId, child, folderContainer, tally);
      if (signal) return signal;
    }
    return undefined;
  }

  /**
   * Per-user body, isolated in its own function specifically so one user's failure can never abort
   * the whole workload — confirmed live against a real tenant: a user with no OneDrive provisioned
   * yet (Graph: 404 "User's mysite not found", the same real, expected condition the Cleaning
   * module's getUserDriveQuota already treats as "not provisioned," not an error) must not stop
   * every OTHER user in the same run from being processed (spec §27/§41: one failure isn't allowed
   * to silently drop the rest of a batch).
   */
  async function processUser(user: { id: string; upn: string; displayName: string }, tally: FairShareTally): Promise<"cancelled" | "paused" | undefined> {
    // The container's display_name doubles as its lookup key (services/dataDump/containers.ts's
    // getOrCreateContainer matches on kind+displayName+parent) AND as what the Report/Results view
    // shows the operator — it MUST include the user's own identity: with a generic name shared by
    // every user, a second user's container would collide with the first user's (same kind, same
    // name, same null parent) and reuse THEIR drive-scoped folder id, which Graph would then reject
    // for every subsequent user. Real bug, caught by re-reading this after the report-drilldown
    // request below rather than by a run that happened to exercise 2+ successful users at once.
    const rootLabel = `${user.displayName} (${user.upn})`;
    const folderCreationLabel = `${ctx.namingPrefix} - OneDrive`;
    const rootContainer = await getOrCreateContainer(ctx.taskId, "root_folder", rootLabel, null, async () => {
      const folder = await createOneDriveFolder(ctx.client, user.id, null, folderCreationLabel);
      return { graphId: folder.id };
    });

    const folderSeed = forkSeed(ctx.operationSeed, `onedrive:${user.id}:folders`);
    const tree = buildFolderTree(createSeededSource(folderSeed), config.rootFolders, config.subFoldersPerFolder, config.maxFolderDepth, config.namingStyle);

    for (const root of tree) {
      const signal = await processFolderNode(user.id, root, rootContainer, tally);
      if (signal) return signal;
    }
    return undefined;
  }

  for (const user of users) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    let accountedFolders = 0;
    let accountedFiles = 0;
    try {
      const signal = await processUser(user, {
        folders: (n) => (accountedFolders += n),
        files: (n) => (accountedFiles += n),
      });
      if (signal) return signal;
    } catch (err) {
      // This one user's OneDrive is unavailable (no provisioned drive, access blocked, etc.) or some
      // other per-user Graph failure — record it and move on to the next user, never abort the whole
      // workload for a single user's problem. Whatever share of this user's folders/files never got
      // a chance to run (most commonly all of it, when the failure happens at root-folder creation)
      // is credited as failed here so Requested still equals Created + Failed + Skipped once the
      // operation finishes, instead of leaving a permanent, unexplained "Remaining" gap.
      hadFailures = true;
      const message = err instanceof Error ? err.message : String(err);
      console.error(`[data-dump] OneDrive: user ${user.upn} failed`, message);
      const remainingFolders = Math.max(foldersPerUser - accountedFolders, 0);
      const remainingFiles = Math.max(filesPerUser - accountedFiles, 0);
      if (remainingFolders > 0) await bumpSubcount(ctx.taskId, "folder", { requested: remainingFolders, created: 0, failed: remainingFolders });
      if (remainingFiles > 0) await bumpSubcount(ctx.taskId, "file", { requested: remainingFiles, created: 0, failed: remainingFiles });
      await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: remainingFolders + remainingFiles, skipped: 0, sizeBytes: 0 });
    }
  }

  return hadFailures ? "completed_with_errors" : "completed";
}
