import { ReportsTable } from "../components/reports/ReportsTable";
import { MOCK_JOBS } from "../api/mockData";
import { formatBytes } from "../utils/format";

export function ReportsPage() {
  const totalReclaimed = MOCK_JOBS.reduce((sum, j) => sum + j.bytesReclaimed, 0);
  const totalItems = MOCK_JOBS.reduce((sum, j) => sum + j.progress.completed, 0);
  const totalFailed = MOCK_JOBS.reduce((sum, j) => sum + j.progress.failed, 0);

  function handleExportAudit(jobId: string) {
    // In production: GET /api/v1/jobs/:jobId/audit/export?format=csv
    console.log("export audit for", jobId);
  }

  return (
    <div className="mx-auto max-w-5xl px-8 py-8">
      <h1 className="mb-1 text-xl font-semibold text-slate-900">Reports</h1>
      <p className="mb-6 text-sm text-slate-500">Cleanup job history, storage reclaimed, and audit trail.</p>

      <div className="mb-6 grid grid-cols-3 gap-4">
        <StatCard label="Storage reclaimed" value={formatBytes(totalReclaimed)} />
        <StatCard label="Items deleted" value={totalItems.toLocaleString()} />
        <StatCard label="Items failed" value={totalFailed.toLocaleString()} accent={totalFailed > 0} />
      </div>

      <ReportsTable jobs={MOCK_JOBS} onExportAudit={handleExportAudit} />
    </div>
  );
}

function StatCard({ label, value, accent }: { label: string; value: string; accent?: boolean }) {
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <div className="text-xs font-medium uppercase tracking-wide text-slate-400">{label}</div>
      <div className={`mt-1 text-2xl font-semibold ${accent ? "text-rose-600" : "text-slate-800"}`}>{value}</div>
    </div>
  );
}
