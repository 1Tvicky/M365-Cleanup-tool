/**
 * File-size distribution helpers (spec §10: target total size / average / min / max, "no artificial
 * application-level maximum"). A triangular distribution peaked at `averageBytes` is realistic
 * enough for demo/test data variety without needing a full histogram config, and is cheap to sample.
 */

export function pickFileSizeBytes(rng: () => number, minBytes: number, maxBytes: number, averageBytes: number): number {
  const lo = Math.max(1, Math.min(minBytes, maxBytes));
  const hi = Math.max(lo, maxBytes);
  const mode = Math.min(Math.max(averageBytes, lo), hi);
  // Triangular distribution via inverse CDF.
  const u = rng();
  const f = (mode - lo) / (hi - lo || 1);
  if (u < f) return Math.round(lo + Math.sqrt(u * (hi - lo) * (mode - lo)));
  return Math.round(hi - Math.sqrt((1 - u) * (hi - lo) * (hi - mode)));
}

/** How many files at ~averageBytes it takes to approximate targetTotalBytes — used when a caller configures a target size rather than an explicit file count. Always at least 1 if a target is set at all. */
export function estimateFileCountForTargetSize(targetTotalBytes: number, averageBytes: number): number {
  if (targetTotalBytes <= 0) return 0;
  return Math.max(1, Math.round(targetTotalBytes / Math.max(1, averageBytes)));
}

/** Graph's small-file-upload threshold — above this, an upload session (chunked PUT) is required instead of a single PUT :content call. Not an app-imposed limit; this is Microsoft's own documented boundary (services/dataDump/uploadHelper.ts uses this to pick the upload strategy). */
export const GRAPH_SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024;

/** Graph upload-session chunk size must be a multiple of 320 KiB except for the final chunk — this is Microsoft's own documented requirement, not an app choice. Using a larger multiple (a few MB) reduces round-trips for big files while staying inside Graph's per-chunk size guidance. */
export const GRAPH_UPLOAD_CHUNK_BYTES = 320 * 1024 * 10; // 3,200 KiB
