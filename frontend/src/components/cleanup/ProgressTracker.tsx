import type { JobStatus } from "../../types";

const STATUS_COPY: Record<JobStatus, string> = {
  export_in_progress: "Exporting backup manifest before any deletes run…",
  queued: "Queued for execution…",
  running: "Deleting items…",
  completed: "Completed successfully",
  completed_with_errors: "Completed with some failures",
  failed: "Job failed",
  cancelled: "Cancelled — items already deleted before cancellation remain deleted",
};

export function ProgressTracker({
  status,
  progress,
  onCancel,
}: {
  status: JobStatus;
  progress: { total: number; completed: number; failed: number };
  onCancel: () => void;
}) {
  const pct = progress.total === 0 ? 0 : Math.round(((progress.completed + progress.failed) / progress.total) * 100);
  const canCancel = status === "queued" || status === "running";

  return (
    <div className="mt-8 rounded-xl border border-slate-200 bg-white p-6">
      <div className="flex items-center justify-between">
        <h2 className="text-lg font-semibold text-slate-800">4. Progress</h2>
        {canCancel && (
          <button
            onClick={onCancel}
            className="rounded-md border border-slate-200 px-3 py-1 text-xs font-medium text-slate-600 hover:border-rose-300 hover:text-rose-600"
          >
            Cancel job
          </button>
        )}
      </div>
      <p className="mt-1 text-sm text-slate-500">{STATUS_COPY[status]}</p>

      <div className="mt-4 h-2.5 w-full overflow-hidden rounded-full bg-slate-100">
        <div
          className={`h-full rounded-full transition-all ${status === "failed" ? "bg-rose-500" : "bg-brand-500"}`}
          style={{ width: `${pct}%` }}
        />
      </div>
      <div className="mt-2 flex justify-between text-xs text-slate-500">
        <span>
          {progress.completed + progress.failed} of {progress.total} items processed
        </span>
        <span>{pct}%</span>
      </div>
      {progress.failed > 0 && (
        <div className="mt-2 text-xs text-rose-600">{progress.failed} item(s) failed — see audit log for details</div>
      )}
    </div>
  );
}
