import type { CleaningSyncOperationStatus } from "../types/cleaning.js";

/**
 * Combines the 1-3 sub-resource statuses of one "Sync Now" operation (its OneDrive sync_jobs row,
 * SharePoint sync_jobs row, and/or Teams cleaning_scans row — whichever were actually triggered)
 * into a single overall status. Kept in this side-effect-free module (rather than routes/cleaning.ts,
 * which transitively imports jobs/queue.ts and its Redis connections) specifically so it's
 * unit-testable without touching Redis — same reasoning as graph/cleanupDeletion.ts's classifyDeleteError.
 */
export function computeSyncStatus(subStatuses: string[]): CleaningSyncOperationStatus {
  if (subStatuses.length === 0) return "completed";
  if (subStatuses.some((s) => s === "queued" || s === "running")) {
    return subStatuses.every((s) => s === "queued") ? "queued" : "running";
  }
  const unsuccessful = subStatuses.filter((s) => s === "failed" || s === "cancelled").length;
  if (unsuccessful === subStatuses.length) return "failed";
  if (unsuccessful > 0) return "completed_with_errors";
  return "completed";
}
