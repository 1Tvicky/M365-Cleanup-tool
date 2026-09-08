import { useEffect, useState } from "react";
import {
  cleanupReportUrl,
  getCleanupProgress,
  retryCleanup,
  type CleanupOperationRow,
  type CleanupOperationStatus,
  type CleanupProgress,
  type CleanupResourceType,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { CategoryRow } from "../cleaning/CleanupProgress";
import { Spinner } from "../cleaning/ItemFilesDrilldown";
import { PageFooter, useDebouncedValue } from "../cleaning/DiscoveryTable";
import { formatBytes, formatDate } from "../../utils/format";

const TERMINAL_STATUSES = new Set<CleanupOperationStatus>(["completed", "completed_with_errors", "failed", "cancelled"]);

const STATUS_STYLES: Record<CleanupOperationStatus, string> = {
  queued: "bg-slate-100 text-slate-600",
  running: "bg-blue-50 text-blue-700",
  completed: "bg-emerald-50 text-emerald-700",
  completed_with_errors: "bg-amber-50 text-amber-700",
  failed: "bg-rose-50 text-rose-700",
  cancelled: "bg-slate-100 text-slate-500",
};

/** Plain-language status vocabulary (not the raw enum) shown in the badge. */
const STATUS_LABEL: Record<CleanupOperationStatus, string> = {
  queued: "Pending",
  running: "Running",
  completed: "Completed",
  completed_with_errors: "Partially Completed",
  failed: "Failed",
  cancelled: "Cancelled",
};

/** Two-segment green/red bar over a slate track — same convention as CleanupProgressView's bar. */
function ProgressBar({ completed, failed, total }: { completed: number; failed: number; total: number }) {
  const safeTotal = Math.max(total, 1);
  return (
    <div className="flex h-1.5 w-32 overflow-hidden rounded-full bg-slate-100">
      <div className="h-full bg-emerald-500" style={{ width: `${(completed / safeTotal) * 100}%` }} />
      <div className="h-full bg-rose-500" style={{ width: `${(failed / safeTotal) * 100}%` }} />
    </div>
  );
}

/** Single-segment bar for bytes — there's no "failed bytes" concept (a failed item just never contributes its bytes to bytesCleared), so this is cleared-vs-total only. */
function BytesBar({ cleared, total }: { cleared: number; total: number }) {
  const safeTotal = Math.max(total, 1);
  return (
    <div className="flex h-1.5 w-32 overflow-hidden rounded-full bg-slate-100">
      <div className="h-full bg-[#1b2fc4]" style={{ width: `${(cleared / safeTotal) * 100}%` }} />
    </div>
  );
}

/** Small ⟳ icon button — re-triggers a manual out-of-band fetch alongside the row's own 3s poll (a plain idempotent GET, so an occasional overlap with the timer is harmless; `spinning` just guards against double-click spam). */
function RefreshIconButton({ spinning, onClick }: { spinning: boolean; onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={spinning}
      aria-label="Refresh"
      title="Refresh"
      className="rounded-md p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600 disabled:cursor-not-allowed"
    >
      <span className={`inline-block ${spinning ? "animate-spin" : ""}`}>⟳</span>
    </button>
  );
}

function OperationRow({
  operation,
  onViewDetails,
  onRetried,
}: {
  operation: CleanupOperationRow;
  onViewDetails: (operationId: string) => void;
  onRetried: (newOperationId: string) => void;
}) {
  // live starts as `operation` plus zeroed placeholders for the fields only getCleanupProgress
  // (not the list endpoint) returns — byType/files/bytes — so this can be typed as the richer
  // CleanupProgress throughout rather than juggling two types; the placeholder values are replaced
  // within moments by the poll below (it fetches immediately on mount, not just after 3s).
  const [live, setLive] = useState<CleanupProgress>(() => ({
    ...operation,
    byType: {} as CleanupProgress["byType"],
    filesTotal: 0,
    filesCompleted: 0,
    bytesTotal: 0,
    bytesCleared: 0,
  }));
  const [retrying, setRetrying] = useState(false);
  const [retryError, setRetryError] = useState<string | null>(null);
  const [expanded, setExpanded] = useState(false);
  const [manualRefreshing, setManualRefreshing] = useState(false);

  // Merges rather than replaces — `operation` (the list endpoint's row) never carries byType/files/
  // bytes, so a plain overwrite would erase whatever the richer poll below already fetched.
  useEffect(() => setLive((prev) => ({ ...prev, ...operation })), [operation]);

  useEffect(() => {
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const data = await getCleanupProgress(operation.id).catch(() => null);
      if (cancelled || !data) return;
      setLive(data);
      // Always fetch once — even for an already-terminal operation, this is the only call that
      // ever populates byType/bytes (the list endpoint's row never carries them), which the new
      // Data Processed column and inline workload expand both need. Only the *recurring* poll
      // stops once terminal, matching the original behavior for everything else on this row.
      if (!TERMINAL_STATUSES.has(data.status)) timer = setTimeout(poll, 3000);
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [operation.id, live.status]);

  async function handleManualRefresh() {
    setManualRefreshing(true);
    const data = await getCleanupProgress(operation.id).catch(() => null);
    if (data) setLive(data);
    setManualRefreshing(false);
  }

  async function handleRetry() {
    setRetrying(true);
    setRetryError(null);
    try {
      const { operationId: newOperationId } = await retryCleanup(operation.id);
      onRetried(newOperationId);
    } catch (err) {
      setRetryError(err instanceof ApiClientError ? err.message : "Couldn't retry failed items.");
    } finally {
      setRetrying(false);
    }
  }

  const isTerminal = TERMINAL_STATUSES.has(live.status);
  const isRunning = !isTerminal;
  const categories = Object.entries(live.byType).filter(([, v]) => v.total > 0);
  const panelId = `operation-panel-${operation.id}`;

  return (
    <>
      <tr className="hover:bg-slate-50/60">
        <td className="px-4 py-3">
          <button
            type="button"
            aria-expanded={expanded}
            aria-controls={panelId}
            aria-label={expanded ? "Collapse" : "Expand"}
            onClick={() => setExpanded((e) => !e)}
            className="text-slate-400 hover:text-slate-600"
          >
            <span className={`inline-block transition-transform ${expanded ? "rotate-90" : ""}`}>▸</span>
          </button>
        </td>
        <td className="px-4 py-3 font-medium text-slate-800">
          <button onClick={() => onViewDetails(operation.id)} className="text-left hover:text-[#1b2fc4] hover:underline">
            {live.label}
          </button>
          {live.retryOfOperationId && <div className="text-xs text-slate-400">Retry</div>}
        </td>
        <td className="px-4 py-3">
          <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${STATUS_STYLES[live.status]}`}>
            <span className="inline-flex items-center gap-1.5">
              {isRunning && <Spinner className="h-3 w-3" />}
              {STATUS_LABEL[live.status]}
            </span>
          </span>
        </td>
        <td className="px-4 py-3">
          <ProgressBar completed={live.successfulItems} failed={live.failedItems} total={live.totalItems} />
        </td>
        <td className="px-4 py-3 text-right text-slate-600">
          {live.processedItems}/{live.totalItems}
          {live.failedItems > 0 && <span className="text-rose-600"> ({live.failedItems} failed)</span>}
        </td>
        <td className="px-4 py-3 text-right text-slate-600">
          {live.bytesTotal > 0 ? (
            <div className="flex flex-col items-end gap-1">
              <BytesBar cleared={live.bytesCleared} total={live.bytesTotal} />
              <span>{formatBytes(live.bytesCleared)} / {formatBytes(live.bytesTotal)}</span>
            </div>
          ) : (
            "—"
          )}
        </td>
        <td className="px-4 py-3 text-slate-600">{live.requestedBy?.displayName ?? live.requestedBy?.email ?? "—"}</td>
        <td className="px-4 py-3 text-slate-600">{formatDate(live.createdAt)}</td>
        <td className="px-4 py-3">
          <div className="flex flex-wrap items-center justify-end gap-2">
            <RefreshIconButton spinning={manualRefreshing} onClick={handleManualRefresh} />
            {isTerminal && live.processedItems > 0 && (
              <a
                href={cleanupReportUrl(operation.id)}
                className="rounded-md border border-slate-200 px-3 py-1 text-xs font-medium text-slate-600 hover:border-[#1b2fc4]/40 hover:text-[#1b2fc4]"
              >
                Download
              </a>
            )}
            {isTerminal && live.failedItems > 0 && (
              <button
                onClick={handleRetry}
                disabled={retrying}
                className="rounded-md bg-[#1b2fc4] px-3 py-1 text-xs font-medium text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {retrying ? "Retrying…" : "Retry Failed"}
              </button>
            )}
          </div>
          {retryError && <p className="mt-1 text-right text-xs text-rose-600">{retryError}</p>}
        </td>
      </tr>
      {expanded && (
        <tr id={panelId}>
          <td colSpan={9} className="border-b border-slate-100 bg-slate-50/60 p-0">
            {categories.length === 0 ? (
              <p className="px-4 py-3 text-sm text-slate-400">No workload details yet.</p>
            ) : (
              <div className="px-4 py-2">
                {categories.map(([type, v]) => (
                  <CategoryRow key={type} operationId={operation.id} resourceType={type as CleanupResourceType} v={v} isRunning={isRunning} />
                ))}
              </div>
            )}
          </td>
        </tr>
      )}
    </>
  );
}

export function ReportsTable({
  operations,
  onViewDetails,
  onRetried,
  page,
  totalPages,
  total,
  onGoToPage,
  loading,
  search,
  onSearchChange,
}: {
  operations: CleanupOperationRow[];
  onViewDetails: (operationId: string) => void;
  onRetried: (newOperationId: string) => void;
  page: number;
  totalPages: number;
  total: number;
  onGoToPage: (page: number) => void;
  loading: boolean;
  search: string;
  onSearchChange: (search: string) => void;
}) {
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="border-b border-slate-100 px-4 py-3">
        <input
          type="search"
          value={search}
          onChange={(e) => onSearchChange(e.target.value)}
          placeholder="Search by cleanup name…"
          className="w-full max-w-xs rounded-md border border-slate-200 px-3 py-1.5 text-sm text-slate-700 focus:border-[#1b2fc4] focus:outline-none"
        />
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
            <tr>
              <th className="w-8 px-4 py-3" />
              <th className="px-4 py-3 font-medium">Cleanup Name</th>
              <th className="px-4 py-3 font-medium">Status</th>
              <th className="px-4 py-3 font-medium">Progress</th>
              <th className="px-4 py-3 font-medium text-right">Items</th>
              <th className="px-4 py-3 font-medium text-right">Data Processed</th>
              <th className="px-4 py-3 font-medium">Requested by</th>
              <th className="px-4 py-3 font-medium">Started</th>
              <th className="px-4 py-3 font-medium"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {operations.length === 0 && !loading && (
              <tr>
                <td colSpan={9} className="px-4 py-8 text-center text-sm text-slate-400">
                  {search ? `No cleanup operations match "${search}".` : "No cleanup operations yet."}
                </td>
              </tr>
            )}
            {operations.map((op) => (
              <OperationRow key={op.id} operation={op} onViewDetails={onViewDetails} onRetried={onRetried} />
            ))}
          </tbody>
        </table>
      </div>
      <PageFooter page={page} totalPages={totalPages} total={total} disabled={loading} onGoToPage={onGoToPage} />
    </div>
  );
}
