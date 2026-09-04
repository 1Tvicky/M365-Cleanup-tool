import { useCallback, useEffect, useState } from "react";
import {
  cancelCleanup,
  getCleanupProgress,
  getCleanupRecentFiles,
  type CleanupProgress as CleanupProgressData,
  type CleanupRecentFile,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { formatBytes, formatDate } from "../../utils/format";

const TERMINAL_STATUSES = new Set(["completed", "completed_with_errors", "failed", "cancelled"]);

/** Two-segment green/red bar over a slate track — same convention as ManageCloudsRowView's sync-progress bar. */
function ProgressBar({ completed, failed, total }: { completed: number; failed: number; total: number }) {
  const safeTotal = Math.max(total, 1);
  return (
    <div className="flex h-1.5 w-full overflow-hidden rounded-full bg-slate-100">
      <div className="h-full bg-emerald-500" style={{ width: `${(completed / safeTotal) * 100}%` }} />
      <div className="h-full bg-rose-500" style={{ width: `${(failed / safeTotal) * 100}%` }} />
    </div>
  );
}

const CATEGORY_LABEL: Record<string, string> = {
  onedrive_account: "OneDrive",
  sharepoint_site: "SharePoint",
  channel: "Teams Channels",
  chat: "Direct Messages",
};

/** Polls the operation every 3s (same self-rearming setTimeout convention used by TeamsView/Dashboard) until it reaches a terminal status, then hands off to the Results screen. */
export function CleanupProgressView({ operationId, onFinished }: { operationId: string; onFinished: () => void }) {
  const [progress, setProgress] = useState<CleanupProgressData | null>(null);
  const [recentFiles, setRecentFiles] = useState<CleanupRecentFile[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [cancelling, setCancelling] = useState(false);

  const refresh = useCallback(async () => {
    try {
      const [data, { files }] = await Promise.all([getCleanupProgress(operationId), getCleanupRecentFiles(operationId, 8)]);
      setProgress(data);
      setRecentFiles(files);
      return data;
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't load cleanup progress.");
      return null;
    }
  }, [operationId]);

  useEffect(() => {
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const data = await refresh();
      if (cancelled) return;
      if (data && TERMINAL_STATUSES.has(data.status)) {
        onFinished();
        return;
      }
      timer = setTimeout(poll, 3000);
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refresh]);

  async function handleCancel() {
    setCancelling(true);
    try {
      await cancelCleanup(operationId);
      await refresh();
    } catch {
      // A 409 here just means it already finished — the next poll will pick up the real status.
    } finally {
      setCancelling(false);
    }
  }

  if (error) {
    return <p className="mx-auto max-w-2xl px-8 py-10 text-sm text-rose-600">{error}</p>;
  }
  if (!progress) {
    return <p className="mx-auto max-w-2xl px-8 py-10 text-sm text-slate-500">Starting cleanup…</p>;
  }

  const remaining = progress.totalItems - progress.processedItems;
  const isRunning = progress.status === "queued" || progress.status === "running";

  return (
    <div className="mx-auto max-w-2xl px-8 py-10">
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Cleaning Microsoft 365</h2>
      <p className="mb-6 text-sm text-slate-500">{isRunning ? "This runs in the background — you can leave this page and come back." : "Finishing up…"}</p>

      <div className="rounded-xl border border-slate-200 bg-white p-6">
        <ProgressBar completed={progress.successfulItems} failed={progress.failedItems} total={progress.totalItems} />
        <div className="mt-3 grid grid-cols-2 gap-x-6 gap-y-1 text-sm text-slate-600 sm:grid-cols-3">
          <div>Total items: <span className="font-medium text-slate-800">{progress.totalItems.toLocaleString()}</span></div>
          <div>Completed: <span className="font-medium text-emerald-600">{progress.successfulItems.toLocaleString()}</span></div>
          <div>Failed: <span className="font-medium text-rose-600">{progress.failedItems.toLocaleString()}</span></div>
          <div>Skipped: <span className="font-medium text-slate-500">{progress.skippedItems.toLocaleString()}</span></div>
          <div>Remaining: <span className="font-medium text-slate-800">{Math.max(remaining, 0).toLocaleString()}</span></div>
        </div>

        <div className="mt-5 space-y-2 border-t border-slate-100 pt-4">
          {Object.entries(progress.byType)
            .filter(([, v]) => v.total > 0)
            .map(([type, v]) => (
              <div key={type} className="flex items-center justify-between text-sm text-slate-600">
                <span>{CATEGORY_LABEL[type] ?? type}</span>
                <span>
                  {(v.completed + v.failed + v.skipped + v.unsupported).toLocaleString()} / {v.total.toLocaleString()}
                </span>
              </div>
            ))}
          {progress.filesTotal > 0 && (
            <div className="flex items-center justify-between border-t border-slate-100 pt-2 text-sm text-slate-600">
              <span>Files removed</span>
              <span>
                {progress.filesCompleted.toLocaleString()} / {progress.filesTotal.toLocaleString()}
              </span>
            </div>
          )}
          {progress.bytesTotal > 0 && (
            <div className="flex items-center justify-between text-sm text-slate-600">
              <span>Data cleared</span>
              <span>
                {formatBytes(progress.bytesCleared)} / {formatBytes(progress.bytesTotal)}
              </span>
            </div>
          )}
        </div>
      </div>

      {recentFiles.length > 0 && (
        <div className="mt-4 rounded-xl border border-slate-200 bg-white p-4">
          <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Recently removed</p>
          <ul className="space-y-1.5 text-sm">
            {recentFiles.map((f, i) => (
              <li key={i} className="flex items-center justify-between gap-3">
                <span className={`truncate ${f.status === "failed" ? "text-rose-500" : "text-slate-700"}`}>
                  {f.fileName} <span className="text-slate-400">· {f.resourceName}</span>
                </span>
                <span className="shrink-0 text-xs text-slate-400">{formatDate(f.completedAt)}</span>
              </li>
            ))}
          </ul>
        </div>
      )}

      {isRunning && (
        <button
          onClick={handleCancel}
          disabled={cancelling || progress.cancelRequestedAt != null}
          className="mt-6 rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {progress.cancelRequestedAt ? "Cancelling…" : "Cancel Cleanup"}
        </button>
      )}
    </div>
  );
}
