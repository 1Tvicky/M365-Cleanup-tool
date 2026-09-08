import { useEffect, useRef, useState } from "react";
import { ReportsTable } from "../components/reports/ReportsTable";
import { CleanupProgressView } from "../components/cleaning/CleanupProgress";
import { CleanupResultsView } from "../components/cleaning/CleanupResults";
import { useDebouncedValue } from "../components/cleaning/DiscoveryTable";
import {
  getCleanupOperationsSummary,
  getCleanupProgress,
  listCleanupOperations,
  type CleanupOperationRow,
  type CleanupOperationStatus,
  type CleanupOperationsSummary,
} from "../api/cleaning";
import { ApiClientError } from "../api/client";
import { formatBytes, formatDate } from "../utils/format";

const PAGE_SIZE = 20;
const TERMINAL_STATUSES = new Set<CleanupOperationStatus>(["completed", "completed_with_errors", "failed", "cancelled"]);

/** Reads `?operationId=&page=` back out — used on mount (deep link / reload) and on Back/Forward. */
function reportsStateFromUrl(): { operationId: string | null; page: number } {
  const params = new URLSearchParams(window.location.search);
  return { operationId: params.get("operationId"), page: Math.max(1, Number(params.get("page")) || 1) };
}

/**
 * Lists real cleanup_operations (no separate "job" model) and drills into one via a plain internal
 * list/detail switch — reuses CleanupProgressView/CleanupResultsView unchanged, exactly as they're
 * used from CleaningPage's own view switch. Owns its own URL sync (like CleaningPage does for its
 * tenant/view): `?operationId=` names the open detail (and doubles as the "Start Cleanup" redirect
 * target, set by App.tsx before this component even mounts), `?page=` names the list's current page.
 */
export function ReportsPage() {
  const [initial] = useState(reportsStateFromUrl);
  const [operations, setOperations] = useState<CleanupOperationRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(initial.page);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [detailOperationId, setDetailOperationId] = useState<string | null>(initial.operationId);
  const [detailIsRunning, setDetailIsRunning] = useState(true);
  // Tracks the previous detailOperationId so the URL-sync effect below can tell "opened/closed a
  // detail" (worth a Back stop) apart from "just changed list page" (shouldn't be — see that effect).
  const prevDetailIdRef = useRef(initial.operationId);
  const [searchInput, setSearchInput] = useState("");
  const search = useDebouncedValue(searchInput, 400);
  const [summary, setSummary] = useState<CleanupOperationsSummary | null>(null);
  // Distinguishes "still loading the very first page" (worth a full-panel "Loading…") from any
  // later load (a page change, a search) where the table itself should stay mounted and just show
  // its own empty/row state — never regresses back to false once the first load settles.
  const [hasLoadedOnce, setHasLoadedOnce] = useState(false);

  function load(targetPage: number, targetSearch: string) {
    setLoading(true);
    listCleanupOperations({ page: targetPage, pageSize: PAGE_SIZE, search: targetSearch || undefined })
      .then((res) => {
        setOperations(res.operations);
        setTotal(res.total);
        setError(null);
      })
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load cleanup reports."))
      .finally(() => {
        setLoading(false);
        setHasLoadedOnce(true);
      });
    getCleanupOperationsSummary()
      .then(setSummary)
      .catch(() => {});
  }

  // A search change always starts back at page 1 — a stale page number from a differently-filtered
  // list wouldn't make sense against the new one.
  useEffect(() => {
    setPage(1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [search]);

  useEffect(() => {
    load(page, search);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, search]);

  // Deep-link / reload with ?operationId= already in the URL — resolve whether it's still running
  // so we land on the right sub-view without waiting for the list to load first. Only ever needs to
  // run once for whatever the URL said at mount; openDetails/the popstate handler cover the rest.
  useEffect(() => {
    if (!initial.operationId) return;
    getCleanupProgress(initial.operationId)
      .then((p) => setDetailIsRunning(!TERMINAL_STATUSES.has(p.status)))
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Keeps the address bar naming the operation (or list page) actually on screen. Opening/closing a
  // detail is a real navigation (pushState — worth a Back stop); paging through the list by itself
  // is a within-page adjustment (replaceState — Back shouldn't have to click through every page).
  useEffect(() => {
    const url = detailOperationId ? `/reports?operationId=${detailOperationId}` : page > 1 ? `/reports?page=${page}` : "/reports";
    const current = `${window.location.pathname}${window.location.search}`;
    const detailChanged = prevDetailIdRef.current !== detailOperationId;
    prevDetailIdRef.current = detailOperationId;
    if (current === url) return;
    if (detailChanged) window.history.pushState({}, "", url);
    else window.history.replaceState({}, "", url);
  }, [detailOperationId, page]);

  useEffect(() => {
    function onPopState() {
      const restored = reportsStateFromUrl();
      setDetailOperationId(restored.operationId);
      setPage(restored.page);
      if (restored.operationId) {
        const op = operations.find((o) => o.id === restored.operationId);
        setDetailIsRunning(op ? !TERMINAL_STATUSES.has(op.status) : true);
      }
    }
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, [operations]);

  function openDetails(operationId: string) {
    const op = operations.find((o) => o.id === operationId);
    setDetailIsRunning(op ? !TERMINAL_STATUSES.has(op.status) : true);
    setDetailOperationId(operationId);
  }

  function closeDetails() {
    setDetailOperationId(null);
    load(page, search);
  }

  if (detailOperationId) {
    return (
      <div className="mx-auto max-w-3xl px-8 py-6">
        <button onClick={closeDetails} className="mb-2 text-sm font-medium text-[#1b2fc4] hover:underline">
          ← Back to Reports
        </button>
        {detailIsRunning ? (
          <CleanupProgressView operationId={detailOperationId} onFinished={() => setDetailIsRunning(false)} />
        ) : (
          <CleanupResultsView
            operationId={detailOperationId}
            onDone={closeDetails}
            onRetried={(newOperationId) => {
              setDetailOperationId(newOperationId);
              setDetailIsRunning(true);
            }}
          />
        )}
      </div>
    );
  }

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));

  return (
    <div className="mx-auto max-w-5xl px-8 py-8">
      <h1 className="mb-1 text-xl font-semibold text-slate-900">Reports</h1>
      <p className="mb-6 text-sm text-slate-500">Cleanup job history and live progress.</p>
      {summary && (
        <div className="mb-6 grid grid-cols-2 gap-4 rounded-xl border border-slate-200 bg-white p-5 sm:grid-cols-4">
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Total Migrations</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.totalOperations.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Processed Migrations</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.processedItems.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Total Items</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.totalItems.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Processed Data Size</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{formatBytes(summary.bytesCleared)}</div>
          </div>
          <div className="col-span-2 text-xs text-slate-400 sm:col-span-4">Updated {formatDate(summary.updatedAt)}</div>
        </div>
      )}
      {error ? (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-rose-600">{error}</p>
      ) : !hasLoadedOnce ? (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-slate-500">Loading…</p>
      ) : (
        <ReportsTable
          operations={operations}
          onViewDetails={openDetails}
          onRetried={(newOperationId) => openDetails(newOperationId)}
          page={page}
          totalPages={totalPages}
          total={total}
          onGoToPage={setPage}
          loading={loading}
          search={searchInput}
          onSearchChange={setSearchInput}
        />
      )}
    </div>
  );
}
