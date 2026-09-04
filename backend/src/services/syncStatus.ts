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
  // A sub-resource ending "completed_with_errors" (e.g. some OneDrive accounts have no provisioned
  // drive) still means the overall sync wasn't fully clean — it must count toward the "some issues"
  // bucket below, not be silently treated the same as a clean "completed".
  const fullyFailed = subStatuses.filter((s) => s === "failed" || s === "cancelled").length;
  if (fullyFailed === subStatuses.length) return "failed";
  const anyIssues = subStatuses.some((s) => s === "failed" || s === "cancelled" || s === "completed_with_errors");
  if (anyIssues) return "completed_with_errors";
  return "completed";
}
