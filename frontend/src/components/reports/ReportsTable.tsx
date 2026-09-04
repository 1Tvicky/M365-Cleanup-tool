import { useEffect, useState } from "react";
import {
  cleanupReportUrl,
  getCleanupProgress,
  retryCleanup,
  type CleanupOperationRow,
  type CleanupOperationStatus,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { PageFooter } from "../cleaning/DiscoveryTable";
import { formatDate } from "../../utils/format";

const TERMINAL_STATUSES = new Set<CleanupOperationStatus>(["completed", "completed_with_errors", "failed", "cancelled"]);

const STATUS_STYLES: Record<CleanupOperationStatus, string> = {
  queued: "bg-slate-100 text-slate-600",
  running: "bg-blue-50 text-blue-700",
  completed: "bg-emerald-50 text-emerald-700",
  completed_with_errors: "bg-amber-50 text-amber-700",
  failed: "bg-rose-50 text-rose-700",
  cancelled: "bg-slate-100 text-slate-500",
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

function OperationRow({
  operation,
  onViewDetails,
  onRetried,
}: {
  operation: CleanupOperationRow;
  onViewDetails: (operationId: string) => void;
  onRetried: (newOperationId: string) => void;
}) {
  const [live, setLive] = useState<CleanupOperationRow>(operation);
  const [retrying, setRetrying] = useState(false);
  const [retryError, setRetryError] = useState<string | null>(null);

  useEffect(() => setLive(operation), [operation]);

  useEffect(() => {
    if (TERMINAL_STATUSES.has(live.status)) return;
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const data = await getCleanupProgress(operation.id).catch(() => null);
      if (cancelled || !data) return;
      setLive(data);
      if (!TERMINAL_STATUSES.has(data.status)) timer = setTimeout(poll, 3000);
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [operation.id, live.status]);

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

  return (
    <tr className="hover:bg-slate-50/60">
      <td className="px-4 py-3 font-medium text-slate-800">
        {live.label}
        {live.retryOfOperationId && <div className="text-xs text-slate-400">Retry</div>}
      </td>
      <td className="px-4 py-3">
        <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${STATUS_STYLES[live.status]}`}>{live.status.replace(/_/g, " ")}</span>
      </td>
      <td className="px-4 py-3">
        <ProgressBar completed={live.successfulItems} failed={live.failedItems} total={live.totalItems} />
      </td>
      <td className="px-4 py-3 text-right text-slate-600">
        {live.processedItems}/{live.totalItems}
        {live.failedItems > 0 && <span className="text-rose-600"> ({live.failedItems} failed)</span>}
      </td>
      <td className="px-4 py-3 text-slate-600">{live.requestedBy?.displayName ?? live.requestedBy?.email ?? "—"}</td>
      <td className="px-4 py-3 text-slate-600">{formatDate(live.createdAt)}</td>
      <td className="px-4 py-3">
        <div className="flex flex-wrap items-center justify-end gap-2">
          <button
            onClick={() => onViewDetails(operation.id)}
            className="rounded-md border border-slate-200 px-3 py-1 text-xs font-medium text-slate-600 hover:border-[#1b2fc4]/40 hover:text-[#1b2fc4]"
          >
            View Details
          </button>
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
}: {
  operations: CleanupOperationRow[];
  onViewDetails: (operationId: string) => void;
  onRetried: (newOperationId: string) => void;
  page: number;
  totalPages: number;
  total: number;
  onGoToPage: (page: number) => void;
  loading: boolean;
}) {
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
          <tr>
            <th className="px-4 py-3 font-medium">Cloud</th>
            <th className="px-4 py-3 font-medium">Status</th>
            <th className="px-4 py-3 font-medium">Progress</th>
            <th className="px-4 py-3 font-medium text-right">Items</th>
            <th className="px-4 py-3 font-medium">Requested by</th>
            <th className="px-4 py-3 font-medium">Started</th>
            <th className="px-4 py-3 font-medium"></th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {operations.map((op) => (
            <OperationRow key={op.id} operation={op} onViewDetails={onViewDetails} onRetried={onRetried} />
          ))}
        </tbody>
      </table>
      <PageFooter page={page} totalPages={totalPages} total={total} disabled={loading} onGoToPage={onGoToPage} />
    </div>
  );
}
