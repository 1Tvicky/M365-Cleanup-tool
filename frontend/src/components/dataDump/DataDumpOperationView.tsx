import { useEffect, useRef, useState } from "react";
import {
  cancelDataDump,
  dataDumpReportUrl,
  getDataDumpOperation,
  pauseDataDump,
  resumeDataDump,
  type DataDumpOperationRow,
  type DataDumpWorkloadTaskSummary,
} from "../../api/dataDump";
import { ApiClientError } from "../../api/client";
import { formatBytes, formatDate } from "../../utils/format";
import { GeneratedResourcesView } from "./GeneratedResourcesView";

const TERMINAL_STATUSES = new Set(["completed", "completed_with_errors", "failed", "cancelled"]);

const STATUS_LABEL: Record<string, string> = {
  queued: "Queued",
  running: "Generating…",
  paused: "Paused",
  completed: "Completed",
  completed_with_errors: "Completed with errors",
  failed: "Failed",
  cancelled: "Cancelled",
  pending: "Pending",
  skipped: "Skipped",
};

const STATUS_COLOR: Record<string, string> = {
  queued: "text-slate-500",
  running: "text-[#1b2fc4]",
  paused: "text-amber-600",
  completed: "text-emerald-600",
  completed_with_errors: "text-amber-600",
  failed: "text-rose-600",
  cancelled: "text-slate-500",
  pending: "text-slate-400",
  skipped: "text-slate-400",
};

const SUBCOUNT_LABEL: Record<string, string> = {
  user: "Users",
  folder: "Folders",
  file: "Files",
  site: "Sites",
  library: "Libraries",
  team: "Teams",
  channel: "Channels",
  channel_member: "Channel Members",
  message: "Messages",
  reply: "Replies",
  email: "Emails",
  attachment: "Attachments",
  calendar_event: "Calendar Events",
  contact: "Contacts",
};

/** Two-segment green/red bar over a slate track — same convention as CleanupProgressView's ProgressBar (spec: "must look and behave like Cleanup"). */
function ProgressBar({ completed, failed, total, isRunning }: { completed: number; failed: number; total: number; isRunning: boolean }) {
  const safeTotal = Math.max(total, 1);
  return (
    <div className={`flex h-1.5 w-full overflow-hidden rounded-full bg-slate-100 ${isRunning ? "animate-pulse" : ""}`}>
      <div className="h-full bg-emerald-500" style={{ width: `${(completed / safeTotal) * 100}%` }} />
      <div className="h-full bg-rose-500" style={{ width: `${(failed / safeTotal) * 100}%` }} />
    </div>
  );
}

/** One workload's subcount breakdown, expandable — mirrors CleanupProgressView's CategoryRow (per-resource-kind requested/created/failed/skipped, spec §27 accounting made visible). */
function WorkloadCategoryRow({ task, isRunning }: { task: DataDumpWorkloadTaskSummary; isRunning: boolean }) {
  const [expanded, setExpanded] = useState(false);
  const settled = task.createdItems + task.failedItems + task.skippedItems;
  const stillWorking = isRunning && (task.status === "running" || task.status === "pending") && settled < task.requestedItems;
  const subEntries = Object.entries(task.subcounts ?? {}).filter(([, v]) => v.requested > 0);

  return (
    <div className="border-b border-slate-100 last:border-b-0">
      <div
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
        className="flex cursor-pointer items-center justify-between py-2 text-sm text-slate-600 hover:text-slate-800"
        onClick={() => setExpanded((e) => !e)}
      >
        <span className="flex items-center gap-2">
          <span className={`inline-block text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`}>▸</span>
          <span className="font-medium capitalize text-slate-700">{task.workload}</span>
          {stillWorking && <span className="h-2 w-2 animate-pulse rounded-full bg-[#1b2fc4]" />}
        </span>
        <span className={STATUS_COLOR[task.status] ?? "text-slate-500"}>{STATUS_LABEL[task.status] ?? task.status}</span>
        <span>
          {settled.toLocaleString()} / {task.requestedItems.toLocaleString()}
        </span>
      </div>
      {expanded && (
        <div className="mb-2 rounded-lg border border-slate-200 bg-white p-3">
          {subEntries.length === 0 ? (
            <p className="text-xs text-slate-400">No sub-resource detail recorded yet.</p>
          ) : (
            <div className="space-y-1.5">
              {subEntries.map(([kind, v]) => (
                <div key={kind} className="flex items-center justify-between text-xs">
                  <span className="text-slate-600">{SUBCOUNT_LABEL[kind] ?? kind}</span>
                  <span className="text-slate-500">
                    <span className="text-emerald-600">{v.created.toLocaleString()} created</span>
                    {v.failed > 0 && <span className="text-rose-600"> · {v.failed.toLocaleString()} failed</span>}
                    {v.skipped > 0 && <span className="text-slate-400"> · {v.skipped.toLocaleString()} skipped</span>}
                    <span> / {v.requested.toLocaleString()} requested</span>
                  </span>
                </div>
              ))}
            </div>
          )}
          {task.errorMessage && <p className="mt-2 rounded bg-rose-50 px-2 py-1 text-xs text-rose-700">{task.errorMessage}</p>}
        </div>
      )}
    </div>
  );
}

/**
 * Shows one Data Dump operation's live progress (polling while not terminal) or its finished
 * results — same visual/structural pattern as components/cleaning/CleanupProgress.tsx's
 * CleanupProgressView (spec: "reuse Cleanup progress architecture"), with Data Dump's own creation
 * vocabulary (Created/Requested, never Deleted) and its own subcount breakdown instead of Cleanup's
 * per-item file drill-down (Data Dump's leaf objects are tracked in batches, not one row per file —
 * see db/migrations/017_data_dump.sql).
 */
export function DataDumpOperationView({ operationId, onClose }: { operationId: string; onClose: () => void }) {
  const [operation, setOperation] = useState<DataDumpOperationRow | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [showResources, setShowResources] = useState(false);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  function load() {
    getDataDumpOperation(operationId)
      .then(setOperation)
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load this Data Dump operation."));
  }

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [operationId]);

  useEffect(() => {
    if (pollRef.current) clearInterval(pollRef.current);
    if (operation && !TERMINAL_STATUSES.has(operation.status)) {
      pollRef.current = setInterval(load, 2000);
    }
    return () => {
      if (pollRef.current) clearInterval(pollRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [operation?.status]);

  async function handlePause() {
    setBusy(true);
    try {
      await pauseDataDump(operationId);
      load();
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't pause this operation.");
    } finally {
      setBusy(false);
    }
  }

  async function handleResume() {
    setBusy(true);
    try {
      await resumeDataDump(operationId);
      load();
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't resume this operation.");
    } finally {
      setBusy(false);
    }
  }

  async function handleCancel() {
    setBusy(true);
    try {
      await cancelDataDump(operationId);
      load();
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't cancel this operation.");
    } finally {
      setBusy(false);
    }
  }

  if (error) {
    return (
      <div className="mx-auto max-w-2xl px-8 py-10">
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-8 text-center text-sm text-rose-600">{error}</p>
        <button onClick={onClose} className="mt-3 text-sm font-medium text-[#1b2fc4] hover:underline">
          ← Back
        </button>
      </div>
    );
  }

  if (!operation) {
    return <p className="mx-auto max-w-2xl px-8 py-10 text-sm text-slate-500">Loading…</p>;
  }

  const isTerminal = TERMINAL_STATUSES.has(operation.status);
  const isRunning = !isTerminal;
  const settled = operation.createdItems + operation.failedItems + operation.skippedItems;

  return (
    <div className="mx-auto max-w-2xl px-8 py-10">
      <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold text-slate-800">
        {operation.label}
        {isRunning && <span className="h-2 w-2 animate-pulse rounded-full bg-[#1b2fc4]" />}
      </h2>
      <p className="mb-1 text-sm text-slate-500">
        {operation.profile.replace(/_/g, " ")} · {operation.workloads.join(" + ")}
        {isRunning && " — this runs in the background, you can leave this page and come back."}
      </p>
      {operation.m365TenantId && <p className="mb-6 text-xs text-slate-400">M365 Tenant ID: {operation.m365TenantId}</p>}
      {!operation.m365TenantId && <div className="mb-6" />}

      <div className="rounded-xl border border-slate-200 bg-white p-6">
        <ProgressBar completed={operation.createdItems} failed={operation.failedItems} total={operation.requestedItems} isRunning={isRunning} />
        <div className="mt-3 grid grid-cols-2 gap-x-6 gap-y-1 text-sm text-slate-600 sm:grid-cols-3">
          <div>
            Requested: <span className="font-medium text-slate-800">{operation.requestedItems.toLocaleString()}</span>
          </div>
          <div>
            Created: <span className="font-medium text-emerald-600">{operation.createdItems.toLocaleString()}</span>
          </div>
          <div>
            Failed: <span className="font-medium text-rose-600">{operation.failedItems.toLocaleString()}</span>
          </div>
          <div>
            Skipped: <span className="font-medium text-slate-500">{operation.skippedItems.toLocaleString()}</span>
          </div>
          <div>
            Remaining: <span className="font-medium text-slate-800">{Math.max(operation.requestedItems - settled, 0).toLocaleString()}</span>
          </div>
        </div>

        <div className="mt-5 border-t border-slate-100 pt-2">
          {operation.tasks.map((task) => (
            <WorkloadCategoryRow key={task.workload} task={task} isRunning={isRunning} />
          ))}
        </div>

        {operation.totalSizeBytes > 0 && (
          <div className="mt-2 border-t border-slate-100 pt-3 text-sm text-slate-600">
            <div className="flex items-center justify-between">
              <span>Data generated</span>
              <span>{formatBytes(operation.totalSizeBytes)}</span>
            </div>
          </div>
        )}

        <div className="mt-3 border-t border-slate-100 pt-3">
          <button onClick={() => setShowResources((v) => !v)} className="text-sm font-medium text-[#1b2fc4] hover:underline">
            {showResources ? "Hide" : "View"} Generated Resources
          </button>
          {showResources && (
            <div className="mt-3">
              <GeneratedResourcesView operationId={operation.id} />
            </div>
          )}
        </div>
      </div>

      {operation.errorMessage && <p className="mt-4 rounded-lg bg-rose-50 px-3 py-2 text-sm text-rose-700">{operation.errorMessage}</p>}

      <div className="mt-6 flex flex-wrap items-center gap-3">
        {operation.status === "running" && (
          <button onClick={handlePause} disabled={busy} className="rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50 disabled:opacity-50">
            Pause
          </button>
        )}
        {operation.status === "paused" && (
          <button onClick={handleResume} disabled={busy} className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:opacity-50">
            Resume
          </button>
        )}
        {!isTerminal && (
          <button onClick={handleCancel} disabled={busy} className="rounded-md border border-rose-200 px-5 py-2 text-sm font-semibold text-rose-600 hover:bg-rose-50 disabled:opacity-50">
            Cancel
          </button>
        )}
        <a href={dataDumpReportUrl(operation.id)} className="rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
          Download Report
        </a>
        <button onClick={onClose} className="ml-auto text-sm font-medium text-[#1b2fc4] hover:underline">
          ← Back
        </button>
      </div>
      <p className="mt-3 text-xs text-slate-400">Created {formatDate(operation.createdAt)}</p>
    </div>
  );
}
