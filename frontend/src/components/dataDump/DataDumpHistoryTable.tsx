import { dataDumpReportUrl, type DataDumpOperationRow } from "../../api/dataDump";
import { PageFooter } from "../cleaning/DiscoveryTable";
import { formatBytes, formatDate } from "../../utils/format";

/**
 * Data Dump's own history table — deliberately not a reskin of ReportsTable.tsx (Cleanup's list):
 * creation vocabulary throughout (Created/Requested/Generated, never Deleted), and no Retry action
 * (spec §37: "Do not display Retry Cleanup on a Data Dump report" — there is no Data Dump equivalent
 * either, a failed item here is just reported, not retried, since generation isn't a destructive
 * action needing that safety net). Reuses PageFooter (generic pagination UI, not Cleanup-specific)
 * unchanged.
 */

const STATUS_STYLES: Record<string, string> = {
  queued: "bg-slate-100 text-slate-600",
  running: "bg-blue-50 text-blue-700",
  paused: "bg-amber-50 text-amber-700",
  completed: "bg-emerald-50 text-emerald-700",
  completed_with_errors: "bg-amber-50 text-amber-700",
  failed: "bg-rose-50 text-rose-700",
  cancelled: "bg-slate-100 text-slate-500",
};

const STATUS_LABEL: Record<string, string> = {
  queued: "Pending",
  running: "Generating",
  paused: "Paused",
  completed: "Completed",
  completed_with_errors: "Partially Completed",
  failed: "Failed",
  cancelled: "Cancelled",
};

export function DataDumpHistoryTable({
  operations,
  onViewDetails,
  page,
  totalPages,
  total,
  onGoToPage,
  loading,
  search,
  onSearchChange,
}: {
  operations: DataDumpOperationRow[];
  onViewDetails: (operationId: string) => void;
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
          placeholder="Search by operation name…"
          className="w-full max-w-xs rounded-md border border-slate-200 px-3 py-1.5 text-sm text-slate-700 focus:border-[#1b2fc4] focus:outline-none"
        />
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-4 py-3 font-medium">Operation</th>
              <th className="px-4 py-3 font-medium">Profile</th>
              <th className="px-4 py-3 font-medium">Workloads</th>
              <th className="px-4 py-3 font-medium">Status</th>
              <th className="px-4 py-3 font-medium text-right">Created / Requested</th>
              <th className="px-4 py-3 font-medium text-right">Data Size</th>
              <th className="px-4 py-3 font-medium">Started</th>
              <th className="px-4 py-3 font-medium"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {operations.length === 0 && !loading && (
              <tr>
                <td colSpan={8} className="px-4 py-8 text-center text-sm text-slate-400">
                  {search ? `No Data Dump operations match "${search}".` : "No Data Dump operations yet."}
                </td>
              </tr>
            )}
            {operations.map((op) => (
              <tr key={op.id} className="hover:bg-slate-50/60">
                <td className="px-4 py-3 font-medium text-slate-800">
                  <button onClick={() => onViewDetails(op.id)} className="text-left hover:text-[#1b2fc4] hover:underline">
                    {op.label}
                  </button>
                </td>
                <td className="px-4 py-3 capitalize text-slate-600">{op.profile.replace(/_/g, " ")}</td>
                <td className="px-4 py-3 text-slate-600">{op.workloads.map((w) => w[0]!.toUpperCase() + w.slice(1)).join(" + ")}</td>
                <td className="px-4 py-3">
                  <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${STATUS_STYLES[op.status] ?? "bg-slate-100 text-slate-600"}`}>
                    {STATUS_LABEL[op.status] ?? op.status}
                  </span>
                </td>
                <td className="px-4 py-3 text-right text-slate-600">
                  {op.createdItems.toLocaleString()} / {op.requestedItems.toLocaleString()}
                  {op.failedItems > 0 && <span className="text-rose-600"> ({op.failedItems.toLocaleString()} failed)</span>}
                </td>
                <td className="px-4 py-3 text-right text-slate-600">{formatBytes(op.totalSizeBytes)}</td>
                <td className="px-4 py-3 text-slate-600">{formatDate(op.createdAt)}</td>
                <td className="px-4 py-3 text-right">
                  <a
                    href={dataDumpReportUrl(op.id)}
                    className="rounded-md border border-slate-200 px-3 py-1 text-xs font-medium text-slate-600 hover:border-[#1b2fc4]/40 hover:text-[#1b2fc4]"
                  >
                    Download
                  </a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <PageFooter page={page} totalPages={totalPages} total={total} disabled={loading} onGoToPage={onGoToPage} />
    </div>
  );
}
