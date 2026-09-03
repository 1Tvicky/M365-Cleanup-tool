import { formatBytes, formatDate } from "../../utils/format";
import type { CleanupJob } from "../../types";

const STATUS_STYLES: Record<CleanupJob["status"], string> = {
  export_in_progress: "bg-slate-100 text-slate-600",
  queued: "bg-slate-100 text-slate-600",
  running: "bg-brand-50 text-brand-700",
  completed: "bg-emerald-50 text-emerald-700",
  completed_with_errors: "bg-amber-50 text-amber-700",
  failed: "bg-rose-50 text-rose-700",
  cancelled: "bg-slate-100 text-slate-500",
};

export function ReportsTable({ jobs, onExportAudit }: { jobs: CleanupJob[]; onExportAudit: (jobId: string) => void }) {
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
          <tr>
            <th className="px-4 py-3 font-medium">Job</th>
            <th className="px-4 py-3 font-medium">Tenant</th>
            <th className="px-4 py-3 font-medium">Status</th>
            <th className="px-4 py-3 font-medium text-right">Items</th>
            <th className="px-4 py-3 font-medium text-right">Reclaimed</th>
            <th className="px-4 py-3 font-medium">Run by</th>
            <th className="px-4 py-3 font-medium">Finished</th>
            <th className="px-4 py-3 font-medium"></th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {jobs.map((j) => (
            <tr key={j.jobId} className="hover:bg-slate-50/60">
              <td className="px-4 py-3 font-mono text-xs text-slate-500">{j.jobId}</td>
              <td className="px-4 py-3 font-medium text-slate-800">{j.tenantName}</td>
              <td className="px-4 py-3">
                <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${STATUS_STYLES[j.status]}`}>
                  {j.status.replace(/_/g, " ")}
                </span>
              </td>
              <td className="px-4 py-3 text-right text-slate-600">
                {j.progress.completed}/{j.progress.total}
                {j.progress.failed > 0 && <span className="text-rose-600"> ({j.progress.failed} failed)</span>}
              </td>
              <td className="px-4 py-3 text-right text-slate-600">{formatBytes(j.bytesReclaimed)}</td>
              <td className="px-4 py-3 text-slate-600">{j.confirmedByEmail}</td>
              <td className="px-4 py-3 text-slate-600">{formatDate(j.finishedAt)}</td>
              <td className="px-4 py-3 text-right">
                <button
                  onClick={() => onExportAudit(j.jobId)}
                  className="rounded-md border border-slate-200 px-3 py-1 text-xs font-medium text-slate-600 hover:border-brand-300 hover:text-brand-600"
                >
                  Export audit CSV
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
