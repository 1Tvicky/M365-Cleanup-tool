import { useCallback, useEffect, useState } from "react";
import {
  cleanupReportUrl,
  getCleanupOperationItems,
  getCleanupProgress,
  retryCleanup,
  type CleanupItemStatus,
  type CleanupOperationItemRow,
  type CleanupProgress,
  type CleanupResourceType,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { PageFooter } from "./DiscoveryTable";
import { formatBytes, formatDate } from "../../utils/format";

const RESOURCE_LABEL: Record<CleanupResourceType, string> = {
  onedrive_account: "OneDrive account",
  sharepoint_site: "SharePoint site",
  channel: "Teams channel",
  chat: "Direct message",
};

const STATUS_STYLE: Record<CleanupItemStatus, { label: string; className: string }> = {
  pending: { label: "Pending", className: "text-slate-400" },
  processing: { label: "In progress", className: "text-blue-600" },
  completed: { label: "Removed", className: "text-emerald-600" },
  failed: { label: "Failed", className: "text-rose-600" },
  skipped: { label: "Skipped", className: "text-slate-500" },
  unsupported: { label: "Not supported", className: "text-amber-600" },
};

const FILTERS: { value: CleanupItemStatus | "all"; label: string }[] = [
  { value: "all", label: "All" },
  { value: "completed", label: "Successful" },
  { value: "failed", label: "Failed" },
  { value: "skipped", label: "Skipped" },
  { value: "unsupported", label: "Unsupported" },
];

const PAGE_SIZE = 20;

/** Never shows raw Graph error bodies/stack traces — only the friendly message the backend already classified (routes/cleaning.ts / graph/cleanupDeletion.ts's classifyDeleteError). */
function friendlyError(item: CleanupOperationItemRow): string | null {
  if (item.status !== "failed") return null;
  return item.errorMessage ?? "Couldn't be removed. It can be retried.";
}

export function CleanupResultsView({ operationId, onDone, onRetried }: { operationId: string; onDone: () => void; onRetried: (newOperationId: string) => void }) {
  const [summary, setSummary] = useState<CleanupProgress | null>(null);
  const [filter, setFilter] = useState<CleanupItemStatus | "all">("all");
  const [page, setPage] = useState(1);
  const [items, setItems] = useState<CleanupOperationItemRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [retrying, setRetrying] = useState(false);
  const [retryError, setRetryError] = useState<string | null>(null);

  useEffect(() => {
    getCleanupProgress(operationId)
      .then(setSummary)
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load cleanup results."));
  }, [operationId]);

  const loadItems = useCallback(
    (targetPage: number, targetFilter: CleanupItemStatus | "all") => {
      setLoading(true);
      getCleanupOperationItems(operationId, { status: targetFilter === "all" ? undefined : targetFilter, page: targetPage, pageSize: PAGE_SIZE })
        .then((res) => {
          setItems(res.items);
          setTotal(res.total);
        })
        .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load cleanup results."))
        .finally(() => setLoading(false));
    },
    [operationId]
  );

  useEffect(() => {
    setPage(1);
    loadItems(1, filter);
  }, [filter, loadItems]);

  async function handleRetry() {
    setRetrying(true);
    setRetryError(null);
    try {
      const { operationId: newOperationId } = await retryCleanup(operationId);
      onRetried(newOperationId);
    } catch (err) {
      setRetryError(err instanceof ApiClientError ? err.message : "Couldn't retry failed items.");
    } finally {
      setRetrying(false);
    }
  }

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));

  return (
    <div className="mx-auto max-w-3xl px-8 py-10">
      <div className="mb-6 flex items-start justify-between gap-4">
        <div>
          <h2 className="mb-1 text-lg font-semibold text-slate-800">Cleanup Completed</h2>
          {summary && (
            <p className="text-sm text-slate-500">
              Successfully cleaned: <span className="font-medium text-emerald-600">{summary.successfulItems.toLocaleString()}</span> · Failed:{" "}
              <span className="font-medium text-rose-600">{summary.failedItems.toLocaleString()}</span> · Skipped:{" "}
              <span className="font-medium text-slate-600">{summary.skippedItems.toLocaleString()}</span>
              {summary.filesTotal > 0 && (
                <>
                  {" "}
                  · Files removed: <span className="font-medium text-slate-800">{summary.filesCompleted.toLocaleString()}</span>
                </>
              )}
              {summary.bytesCleared > 0 && (
                <>
                  {" "}
                  · Data cleared: <span className="font-medium text-slate-800">{formatBytes(summary.bytesCleared)}</span>
                </>
              )}
            </p>
          )}
        </div>
        <a
          href={cleanupReportUrl(operationId)}
          className="shrink-0 rounded-md border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50"
        >
          Download Report (CSV)
        </a>
      </div>

      {summary && (
        <div className="mb-4 rounded-xl border border-slate-200 bg-white p-4">
          <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">By resource type</p>
          <div className="grid grid-cols-1 gap-x-6 gap-y-2 sm:grid-cols-2">
            {Object.entries(summary.byType)
              .filter(([, v]) => v.total > 0)
              .map(([type, v]) => (
                <div key={type} className="flex items-center justify-between text-sm">
                  <span className="text-slate-600">{RESOURCE_LABEL[type as CleanupResourceType]}</span>
                  <span className="flex gap-3 text-xs">
                    <span className="text-slate-500">Total {v.total.toLocaleString()}</span>
                    <span className="text-emerald-600">Success {v.completed.toLocaleString()}</span>
                    <span className="text-rose-600">Failed {v.failed.toLocaleString()}</span>
                    {v.skipped > 0 && <span className="text-slate-500">Skipped {v.skipped.toLocaleString()}</span>}
                    {v.unsupported > 0 && <span className="text-amber-600">Not supported {v.unsupported.toLocaleString()}</span>}
                  </span>
                </div>
              ))}
          </div>
        </div>
      )}

      <div className="mb-4 flex gap-1 rounded-lg bg-slate-100 p-1">
        {FILTERS.map((f) => (
          <button
            key={f.value}
            onClick={() => setFilter(f.value)}
            className={`rounded-md px-3 py-1.5 text-sm font-medium ${filter === f.value ? "bg-white text-slate-800 shadow-sm" : "text-slate-500"}`}
          >
            {f.label}
          </button>
        ))}
      </div>

      <div className="flex max-h-[50vh] flex-col overflow-hidden rounded-xl border border-slate-200 bg-white">
        {loading ? (
          <p className="px-4 py-8 text-center text-sm text-slate-500">Loading…</p>
        ) : error ? (
          <p className="px-4 py-8 text-center text-sm text-rose-600">{error}</p>
        ) : items.length === 0 ? (
          <p className="px-4 py-8 text-center text-sm text-slate-500">No items to show for this filter.</p>
        ) : (
          <>
            <div className="overflow-y-auto overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                  <tr>
                    <th className="px-4 py-3 font-medium">Resource</th>
                    <th className="px-4 py-3 font-medium">Type</th>
                    <th className="px-4 py-3 font-medium">Status</th>
                    <th className="px-4 py-3 font-medium">Completed</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {items.map((item) => (
                    <tr key={item.id} className="hover:bg-slate-50/60">
                      <td className="px-4 py-3 text-slate-700">
                        {item.displayName}
                        {friendlyError(item) && <div className="text-xs text-rose-500">{friendlyError(item)}</div>}
                      </td>
                      <td className="px-4 py-3 text-slate-500">{RESOURCE_LABEL[item.resourceType]}</td>
                      <td className={`px-4 py-3 font-medium ${STATUS_STYLE[item.status].className}`}>{STATUS_STYLE[item.status].label}</td>
                      <td className="px-4 py-3 text-slate-500">{formatDate(item.completedAt)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <PageFooter page={page} totalPages={totalPages} total={total} disabled={loading} onGoToPage={(p) => { setPage(p); loadItems(p, filter); }} />
          </>
        )}
      </div>

      {retryError && <p className="mt-3 text-sm text-rose-600">{retryError}</p>}

      <div className="mt-6 flex items-center gap-3">
        <button onClick={onDone} className="rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
          Done
        </button>
        {summary && summary.failedItems > 0 && (
          <button
            onClick={handleRetry}
            disabled={retrying}
            className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {retrying ? "Retrying…" : `Retry Failed Items (${summary.failedItems.toLocaleString()})`}
          </button>
        )}
      </div>
    </div>
  );
}
