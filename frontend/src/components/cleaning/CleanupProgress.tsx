import { useCallback, useEffect, useState } from "react";
import {
  cancelCleanup,
  getCleanupOperationItems,
  getCleanupProgress,
  getCleanupRecentFiles,
  type CleanupItemStatus,
  type CleanupOperationItemRow,
  type CleanupProgress as CleanupProgressData,
  type CleanupRecentFile,
  type CleanupResourceType,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { ItemFilesDrilldown, Spinner } from "./ItemFilesDrilldown";
import { PageFooter } from "./DiscoveryTable";
import { formatBytes, formatDate } from "../../utils/format";

const TERMINAL_STATUSES = new Set(["completed", "completed_with_errors", "failed", "cancelled"]);
const ITEMS_PAGE_SIZE = 20;

/** Two-segment green/red bar over a slate track — same convention as ManageCloudsRowView's sync-progress bar. Pulses while the operation is still running, so "it's alive" is visible even between polls, not just a static bar. */
function ProgressBar({ completed, failed, total, isRunning }: { completed: number; failed: number; total: number; isRunning: boolean }) {
  const safeTotal = Math.max(total, 1);
  return (
    <div className={`flex h-1.5 w-full overflow-hidden rounded-full bg-slate-100 ${isRunning ? "animate-pulse" : ""}`}>
      <div className="h-full bg-emerald-500" style={{ width: `${(completed / safeTotal) * 100}%` }} />
      <div className="h-full bg-rose-500" style={{ width: `${(failed / safeTotal) * 100}%` }} />
    </div>
  );
}

const CATEGORY_LABEL: Record<string, string> = {
  onedrive_account: "OneDrive",
  sharepoint_site: "SharePoint",
  outlook_mailbox: "Outlook Mail",
  outlook_calendar: "Outlook Calendar",
  outlook_contacts: "Outlook Contacts",
  channel: "Teams Channels",
  chat: "Direct Messages",
};

const ITEM_STATUS_STYLE: Record<CleanupOperationItemRow["status"], { label: string; className: string }> = {
  pending: { label: "Pending", className: "text-slate-400" },
  processing: { label: "In progress", className: "text-blue-600" },
  completed: { label: "Removed", className: "text-emerald-600" },
  failed: { label: "Failed", className: "text-rose-600" },
  skipped: { label: "Skipped", className: "text-slate-500" },
  unsupported: { label: "Not supported", className: "text-amber-600" },
};

const ITEM_STATUS_FILTERS: { value: CleanupItemStatus | "all"; label: string }[] = [
  { value: "all", label: "All" },
  { value: "pending", label: "Pending" },
  { value: "processing", label: "Processing" },
  { value: "completed", label: "Completed" },
  { value: "failed", label: "Failed" },
  { value: "skipped", label: "Skipped" },
  { value: "unsupported", label: "Unsupported" },
];

/** One resource type's aggregate counts — the value shape of CleanupProgress["byType"][string], named here so ReportsTable.tsx can reuse CategoryRow without re-deriving it. */
export interface CategoryTotals {
  total: number;
  completed: number;
  failed: number;
  skipped: number;
  unsupported: number;
}

/**
 * Level 2 of the drill-down: one resource type's real items (mailboxes/accounts/sites), paginated,
 * fetched only once its category header is expanded — mirrors getCleanupOperationItems's existing
 * status filter, now with the resourceType filter added alongside it (routes/cleaning.ts). Each
 * item row has its own expand chevron down to level 3 (ItemFilesDrilldown).
 */
function CategoryItemsList({ operationId, resourceType, isRunning }: { operationId: string; resourceType: CleanupResourceType; isRunning: boolean }) {
  const [page, setPage] = useState(1);
  const [statusFilter, setStatusFilter] = useState<CleanupItemStatus | "all">("all");
  const [items, setItems] = useState<CleanupOperationItemRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [expandedItemId, setExpandedItemId] = useState<string | null>(null);

  const refresh = useCallback(() => {
    setLoading(true);
    getCleanupOperationItems(operationId, {
      resourceType,
      status: statusFilter === "all" ? undefined : statusFilter,
      page,
      pageSize: ITEMS_PAGE_SIZE,
    })
      .then((res) => {
        setItems(res.items);
        setTotal(res.total);
      })
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load items for this category."))
      .finally(() => setLoading(false));
  }, [operationId, resourceType, statusFilter, page]);

  // Changing the filter always starts back at page 1 — a stale page number from a differently-filtered list wouldn't make sense against the new one.
  useEffect(() => {
    setPage(1);
  }, [statusFilter]);

  useEffect(() => {
    refresh();
    // While running, re-fetch this category's items on the same cadence as the rest of the page's
    // poll (3s) so item statuses visibly update without the user re-expanding anything.
    if (!isRunning) return;
    const timer = setInterval(refresh, 3000);
    return () => clearInterval(timer);
  }, [refresh, isRunning]);

  const totalPages = Math.max(1, Math.ceil(total / ITEMS_PAGE_SIZE));

  return (
    <div>
      <div className="flex items-center justify-end border-b border-slate-100 px-4 py-2">
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as CleanupItemStatus | "all")}
          className="rounded-md border border-slate-200 px-2 py-1 text-xs text-slate-600 focus:border-[#1b2fc4] focus:outline-none"
          aria-label="Filter items by status"
        >
          {ITEM_STATUS_FILTERS.map((f) => (
            <option key={f.value} value={f.value}>
              Status: {f.label}
            </option>
          ))}
        </select>
      </div>

      {loading && items.length === 0 ? (
        <p className="px-4 py-3 text-sm text-slate-500">Loading items…</p>
      ) : error ? (
        <p className="px-4 py-3 text-sm text-rose-600">{error}</p>
      ) : items.length === 0 ? (
        <p className="px-4 py-3 text-sm text-slate-400">No items found for this filter.</p>
      ) : (
        <>
          <table className="w-full text-sm">
            <tbody className="divide-y divide-slate-100">
              {items.map((item) => (
                <tr key={item.id} className="align-top">
                  <td className="w-full p-0">
                    <div
                      role="button"
                      tabIndex={0}
                      aria-expanded={expandedItemId === item.id}
                      className="flex cursor-pointer items-center justify-between gap-3 px-4 py-2 hover:bg-slate-50/60"
                      onClick={() => setExpandedItemId((prev) => (prev === item.id ? null : item.id))}
                    >
                      <span className="flex items-center gap-2 text-slate-700">
                        <span className={`inline-block text-slate-400 transition-transform ${expandedItemId === item.id ? "rotate-90" : ""}`}>▸</span>
                        {item.displayName}
                      </span>
                      <span className={`flex items-center gap-1.5 text-xs font-medium ${ITEM_STATUS_STYLE[item.status].className}`}>
                        {(item.status === "pending" || item.status === "processing") && <Spinner />}
                        {ITEM_STATUS_STYLE[item.status].label}
                      </span>
                    </div>
                    {expandedItemId === item.id && (
                      <>
                        {item.filesTotal > 0 && (
                          <div className="flex items-center gap-4 border-t border-slate-100 bg-slate-50/60 px-4 py-1.5 text-xs text-slate-500">
                            <span>
                              Total Files/Folders <span className="font-medium text-slate-700">{item.filesTotal.toLocaleString()}</span>
                            </span>
                            <span>
                              Processed <span className="font-medium text-slate-700">{item.filesCompleted.toLocaleString()}</span>
                            </span>
                          </div>
                        )}
                        <ItemFilesDrilldown operationId={operationId} itemId={item.id} />
                      </>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {totalPages > 1 && <PageFooter page={page} totalPages={totalPages} total={total} onGoToPage={setPage} disabled={loading} />}
        </>
      )}
    </div>
  );
}

/**
 * Level 1 of the drill-down: one resource type's aggregate count, expandable to CategoryItemsList.
 * Shows the spinner while this category still has unsettled items. Exported so ReportsTable.tsx's
 * list-level row expand can reuse this exact component instead of a parallel implementation —
 * it already has everything it needs (operationId, byType's per-category totals, isRunning) from
 * its own existing per-row poll.
 */
export function CategoryRow({ operationId, resourceType, v, isRunning }: {
  operationId: string;
  resourceType: CleanupResourceType;
  v: CategoryTotals;
  isRunning: boolean;
}) {
  const [expanded, setExpanded] = useState(false);
  const settled = v.completed + v.failed + v.skipped + v.unsupported;
  const stillWorking = isRunning && settled < v.total;
  // Includes operationId, not just resourceType — ReportsTable.tsx can render this same category
  // (e.g. "outlook_mailbox") for multiple different operations expanded at once on one page, and
  // DOM ids must stay unique across all of them.
  const panelId = `category-panel-${operationId}-${resourceType}`;

  return (
    <div className="border-b border-slate-100 last:border-b-0">
      <div
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
        aria-controls={panelId}
        className="flex cursor-pointer items-center justify-between py-2 text-sm text-slate-600 hover:text-slate-800"
        onClick={() => setExpanded((e) => !e)}
      >
        <span className="flex items-center gap-2">
          <span className={`inline-block text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`}>▸</span>
          {CATEGORY_LABEL[resourceType] ?? resourceType}
          {stillWorking && <Spinner />}
        </span>
        <span>
          {settled.toLocaleString()} / {v.total.toLocaleString()}
        </span>
      </div>
      {expanded && (
        <div id={panelId} className="mb-2 rounded-lg border border-slate-200 bg-white">
          <CategoryItemsList operationId={operationId} resourceType={resourceType} isRunning={isRunning} />
        </div>
      )}
    </div>
  );
}

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
      <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold text-slate-800">
        Cleaning Microsoft 365
        {isRunning && <Spinner className="h-4 w-4" />}
      </h2>
      <p className="mb-6 text-sm text-slate-500">{isRunning ? "This runs in the background — you can leave this page and come back." : "Finishing up…"}</p>

      <div className="rounded-xl border border-slate-200 bg-white p-6">
        <ProgressBar completed={progress.successfulItems} failed={progress.failedItems} total={progress.totalItems} isRunning={isRunning} />
        <div className="mt-3 grid grid-cols-2 gap-x-6 gap-y-1 text-sm text-slate-600 sm:grid-cols-3">
          <div>Total items: <span className="font-medium text-slate-800">{progress.totalItems.toLocaleString()}</span></div>
          <div>Completed: <span className="font-medium text-emerald-600">{progress.successfulItems.toLocaleString()}</span></div>
          <div>Failed: <span className="font-medium text-rose-600">{progress.failedItems.toLocaleString()}</span></div>
          <div>Skipped: <span className="font-medium text-slate-500">{progress.skippedItems.toLocaleString()}</span></div>
          <div>Remaining: <span className="font-medium text-slate-800">{Math.max(remaining, 0).toLocaleString()}</span></div>
        </div>

        <div className="mt-5 border-t border-slate-100 pt-2">
          {Object.entries(progress.byType)
            .filter(([, v]) => v.total > 0)
            .map(([type, v]) => (
              <CategoryRow key={type} operationId={operationId} resourceType={type as CleanupResourceType} v={v} isRunning={isRunning} />
            ))}
        </div>

        {(progress.filesTotal > 0 || progress.bytesTotal > 0) && (
          <div className="mt-2 space-y-1 border-t border-slate-100 pt-3">
            {progress.filesTotal > 0 && (
              <div className="flex items-center justify-between text-sm text-slate-600">
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
        )}
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
