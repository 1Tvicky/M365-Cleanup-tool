import { useEffect, useRef, useState } from "react";
import { ReportsTable } from "../components/reports/ReportsTable";
import { CleanupProgressView } from "../components/cleaning/CleanupProgress";
import { CleanupResultsView } from "../components/cleaning/CleanupResults";
import { getCleanupProgress, listCleanupOperations, type CleanupOperationRow, type CleanupOperationStatus } from "../api/cleaning";
import { ApiClientError } from "../api/client";

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

  function load(targetPage: number) {
    setLoading(true);
    listCleanupOperations({ page: targetPage, pageSize: PAGE_SIZE })
      .then((res) => {
        setOperations(res.operations);
        setTotal(res.total);
        setError(null);
      })
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load cleanup reports."))
      .finally(() => setLoading(false));
  }

  useEffect(() => {
    load(page);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page]);

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
    load(page);
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
      {loading && operations.length === 0 ? (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-slate-500">Loading…</p>
      ) : error ? (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-rose-600">{error}</p>
      ) : operations.length === 0 ? (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-slate-500">
          No cleanup operations yet. Start a cleanup from the Cleaning page to see it here.
        </p>
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
        />
      )}
    </div>
  );
}
