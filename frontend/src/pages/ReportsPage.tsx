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
import { DataDumpHistoryTable } from "../components/dataDump/DataDumpHistoryTable";
import { DataDumpOperationView } from "../components/dataDump/DataDumpOperationView";
import { getDataDumpSummary, listDataDumpHistory, type DataDumpOperationRow, type DataDumpSummary } from "../api/dataDump";

const PAGE_SIZE = 20;
const TERMINAL_STATUSES = new Set<CleanupOperationStatus>(["completed", "completed_with_errors", "failed", "cancelled"]);

type ReportsTab = "cleanup" | "dataDump";

function readReportsTab(): ReportsTab {
  return new URLSearchParams(window.location.search).get("reportType") === "dataDump" ? "dataDump" : "cleanup";
}

/** Reads `?operationId=&page=` back out — used on mount (deep link / reload) and on Back/Forward. */
function reportsStateFromUrl(): { operationId: string | null; page: number } {
  const params = new URLSearchParams(window.location.search);
  return { operationId: params.get("operationId"), page: Math.max(1, Number(params.get("page")) || 1) };
}

/**
 * Reports — split into two entirely separate tabs, Cleanup Reports and Data Dump Reports (spec
 * §31: "NON-NEGOTIABLE... do NOT show both operation types in one undifferentiated list"). Each
 * tab renders its own operation type exclusively: this file never merges a data_dump_operations row
 * into the same list/table component as a cleanup_operations row, and the two summary stat strips
 * below are deliberately separate cards, never a combined "Total Jobs" figure (spec §39).
 */
export function ReportsPage() {
  const [tab, setTab] = useState<ReportsTab>(readReportsTab);
  const [cleanupSummary, setCleanupSummary] = useState<CleanupOperationsSummary | null>(null);
  const [dataDumpSummary, setDataDumpSummary] = useState<DataDumpSummary | null>(null);

  useEffect(() => {
    getCleanupOperationsSummary().then(setCleanupSummary).catch(() => {});
    getDataDumpSummary().then(setDataDumpSummary).catch(() => {});
  }, []);

  useEffect(() => {
    const url = tab === "dataDump" ? "/reports?reportType=dataDump" : "/reports";
    if (`${window.location.pathname}${window.location.search}` !== url && !window.location.search.includes("operationId")) {
      window.history.replaceState({}, "", url);
    }
  }, [tab]);

  return (
    <div className="mx-auto max-w-5xl px-8 py-8">
      <h1 className="mb-1 text-xl font-semibold text-slate-900">Reports</h1>
      <p className="mb-6 text-sm text-slate-500">Cleanup job history and Data Dump generation history — tracked separately.</p>

      <div className="mb-6 grid grid-cols-1 gap-4 sm:grid-cols-2">
        <button
          onClick={() => setTab("cleanup")}
          className={`rounded-xl border p-4 text-left transition-colors ${tab === "cleanup" ? "border-[#1b2fc4] bg-[#1b2fc4]/5" : "border-slate-200 bg-white hover:border-slate-300"}`}
        >
          <div className="text-xs uppercase tracking-wide text-slate-400">Cleanup Jobs</div>
          <div className="mt-1 text-2xl font-semibold text-slate-900">{cleanupSummary?.totalOperations.toLocaleString() ?? "—"}</div>
          <div className="text-xs text-slate-400">Deletion operations</div>
        </button>
        <button
          onClick={() => setTab("dataDump")}
          className={`rounded-xl border p-4 text-left transition-colors ${tab === "dataDump" ? "border-[#1b2fc4] bg-[#1b2fc4]/5" : "border-slate-200 bg-white hover:border-slate-300"}`}
        >
          <div className="text-xs uppercase tracking-wide text-slate-400">Data Dump Jobs</div>
          <div className="mt-1 text-2xl font-semibold text-slate-900">{dataDumpSummary?.total.toLocaleString() ?? "—"}</div>
          <div className="text-xs text-slate-400">Generation operations</div>
        </button>
      </div>

      <div className="mb-6 flex gap-1 border-b border-slate-200">
        <button
          onClick={() => setTab("cleanup")}
          className={`-mb-px border-b-2 px-4 py-2 text-sm font-medium ${tab === "cleanup" ? "border-[#1b2fc4] text-[#1b2fc4]" : "border-transparent text-slate-500 hover:text-slate-700"}`}
        >
          Cleanup Reports
        </button>
        <button
          onClick={() => setTab("dataDump")}
          className={`-mb-px border-b-2 px-4 py-2 text-sm font-medium ${tab === "dataDump" ? "border-[#1b2fc4] text-[#1b2fc4]" : "border-transparent text-slate-500 hover:text-slate-700"}`}
        >
          Data Dump Reports
        </button>
      </div>

      {tab === "cleanup" ? <CleanupReportsTab summary={cleanupSummary} /> : <DataDumpReportsTab />}
    </div>
  );
}

/** Everything below is the pre-existing Cleanup Reports implementation, unchanged in behavior — only lifted the summary fetch up to the shared parent above and relabeled it "cleanup-only" so it can never be confused with Data Dump's own summary. */
function CleanupReportsTab({ summary }: { summary: CleanupOperationsSummary | null }) {
  const [initial] = useState(reportsStateFromUrl);
  const [operations, setOperations] = useState<CleanupOperationRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(initial.page);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [detailOperationId, setDetailOperationId] = useState<string | null>(initial.operationId);
  const [detailIsRunning, setDetailIsRunning] = useState(true);
  const prevDetailIdRef = useRef(initial.operationId);
  const [searchInput, setSearchInput] = useState("");
  const search = useDebouncedValue(searchInput, 400);
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
  }

  useEffect(() => {
    setPage(1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [search]);

  useEffect(() => {
    load(page, search);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, search]);

  useEffect(() => {
    if (!initial.operationId) return;
    getCleanupProgress(initial.operationId)
      .then((p) => setDetailIsRunning(!TERMINAL_STATUSES.has(p.status)))
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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
      <div>
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
    <div>
      {summary && (
        <div className="mb-6 grid grid-cols-2 gap-4 rounded-xl border border-slate-200 bg-white p-5 sm:grid-cols-4">
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Total Cleanups</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.totalOperations.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Processed Items</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.processedItems.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Total Items</div>
            <div className="mt-1 text-lg font-semibold text-slate-900">{summary.totalItems.toLocaleString()}</div>
          </div>
          <div>
            <div className="text-xs uppercase tracking-wide text-slate-400">Deleted Data Size</div>
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

/** Data Dump Reports — its own list/detail, its own summary, never mixed with Cleanup Reports above (spec §31/§32/§33). */
function DataDumpReportsTab() {
  const [operations, setOperations] = useState<DataDumpOperationRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [detailId, setDetailId] = useState<string | null>(null);
  const [searchInput, setSearchInput] = useState("");
  const search = useDebouncedValue(searchInput, 400);

  function load() {
    setLoading(true);
    listDataDumpHistory({ page, pageSize: PAGE_SIZE, search: search || undefined })
      .then((res) => {
        setOperations(res.operations);
        setTotal(res.total);
        setError(null);
      })
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load Data Dump reports."))
      .finally(() => setLoading(false));
  }

  useEffect(() => {
    setPage(1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [search]);

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, search]);

  if (detailId) {
    return (
      <div>
        <DataDumpOperationView
          operationId={detailId}
          onClose={() => {
            setDetailId(null);
            load();
          }}
        />
      </div>
    );
  }

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));

  if (error) return <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-rose-600">{error}</p>;

  return (
    <DataDumpHistoryTable
      operations={operations}
      onViewDetails={setDetailId}
      page={page}
      totalPages={totalPages}
      total={total}
      onGoToPage={setPage}
      loading={loading}
      search={searchInput}
      onSearchChange={setSearchInput}
    />
  );
}
