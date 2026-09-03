import { useState } from "react";
import { formatBytes, formatDate } from "../../utils/format";
import { Microsoft365Icon } from "./CloudIcons";
import type { CleanupJob, Tenant } from "../../types";

const STATUS_STYLES: Record<Tenant["status"], string> = {
  connected: "bg-emerald-50 text-emerald-700",
  consent_pending: "bg-amber-50 text-amber-700",
  token_error: "bg-rose-50 text-rose-700",
  disconnected: "bg-slate-100 text-slate-500",
};

const STATUS_LABEL: Record<Tenant["status"], string> = {
  connected: "Connected",
  consent_pending: "Awaiting admin consent",
  token_error: "Token error",
  disconnected: "Disconnected",
};

/** A tenant's most recent running/queued job, if any — drives the sync progress bar on its row. */
function activeJobFor(tenant: Tenant, jobs: CleanupJob[]): CleanupJob | null {
  return jobs.find((j) => j.tenantName === tenant.displayName && (j.status === "running" || j.status === "queued" || j.status === "export_in_progress")) ?? null;
}

export function ManageClouds({
  tenants,
  jobs,
  onDisconnect,
}: {
  tenants: Tenant[];
  jobs: CleanupJob[];
  onDisconnect: (tenantId: string) => void;
}) {
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  function toggle(id: string) {
    setExpanded((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  return (
    <div>
      <p className="mb-4 text-sm text-slate-500">Connected tenant status, cleanup progress, and disconnect.</p>

      <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
        {tenants.map((t, i) => {
          const job = activeJobFor(t, jobs);
          const isExpanded = expanded.has(t.id);
          const pct = job ? Math.round(((job.progress.completed + job.progress.failed) / job.progress.total) * 100) : 0;

          return (
            <div key={t.id} className={i > 0 ? "border-t border-slate-100" : ""}>
              <div className="flex flex-wrap items-center gap-x-7 gap-y-3 px-6 py-5">
                <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-full bg-slate-100">
                  <Microsoft365Icon className="h-6 w-6" />
                </div>

                <div className="min-w-[150px]">
                  <div className="text-base font-semibold text-slate-800">{t.displayName}</div>
                  <span className={`mt-1 inline-block rounded-full px-2 py-0.5 text-[11px] font-medium ${STATUS_STYLES[t.status]}`}>
                    {STATUS_LABEL[t.status]}
                  </span>
                </div>

                {job && (
                  <div className="flex min-w-[200px] flex-1 items-center gap-3">
                    <div className="flex-1">
                      <div className="mb-1.5 text-sm text-slate-500">
                        {(job.progress.completed + job.progress.failed).toLocaleString()} of {job.progress.total.toLocaleString()} items{" "}
                        <span className="font-semibold text-slate-700">{pct}%</span>
                      </div>
                      <div className="flex h-2 w-full max-w-[260px] overflow-hidden rounded-full bg-slate-100">
                        <div className="h-full bg-emerald-500" style={{ width: `${Math.round((job.progress.completed / job.progress.total) * 100)}%` }} />
                        <div className="h-full bg-rose-500" style={{ width: `${Math.round((job.progress.failed / job.progress.total) * 100)}%` }} />
                      </div>
                    </div>
                    <span className="animate-spin text-slate-400" aria-hidden title="Cleanup running">
                      ⟳
                    </span>
                  </div>
                )}

                <div className="ml-auto flex items-center gap-5">
                  <span className="hidden text-sm text-slate-500 sm:inline">{t.connectedByAdminUpn ?? "—"}</span>
                  {t.workloads.length > 0 && (
                    <span className="text-sm font-medium text-blue-700">
                      {t.workloads.length} workload{t.workloads.length > 1 ? "s" : ""}
                    </span>
                  )}
                  {t.status !== "disconnected" && t.status !== "consent_pending" && (
                    <button
                      onClick={() => onDisconnect(t.id)}
                      aria-label="Disconnect"
                      className="text-slate-400 hover:text-rose-600"
                    >
                      <TrashIcon className="h-[18px] w-[18px]" />
                    </button>
                  )}
                  <button
                    onClick={() => toggle(t.id)}
                    aria-label="Toggle details"
                    className="flex h-8 w-8 items-center justify-center rounded-full border border-slate-200 text-slate-400 transition-transform hover:text-slate-600"
                  >
                    <ChevronIcon className={`h-4 w-4 transition-transform ${isExpanded ? "rotate-180" : ""}`} />
                  </button>
                </div>
              </div>

              {isExpanded && (
                <div className="grid grid-cols-2 gap-x-6 gap-y-2 border-t border-slate-100 bg-slate-50 px-5 py-4 text-xs text-slate-600 sm:grid-cols-4">
                  <Detail label="M365 tenant ID" value={t.m365TenantId} mono />
                  <Detail label="Connected" value={formatDate(t.connectedAt)} />
                  <Detail
                    label="Token health"
                    value={t.status === "token_error" ? "Refresh failed — reconnect required" : formatDate(t.lastTokenRefreshAt)}
                  />
                  <Detail label="Workloads" value={t.workloads.length > 0 ? t.workloads.join(", ") : "—"} />
                  {job && <Detail label="Bytes reclaimed so far" value={formatBytes(job.bytesReclaimed)} />}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function Detail({ label, value, mono }: { label: string; value: string; mono?: boolean }) {
  return (
    <div>
      <div className="text-[10px] font-medium uppercase tracking-wide text-slate-400">{label}</div>
      <div className={mono ? "font-mono text-[11px] text-slate-700" : "text-slate-700"}>{value}</div>
    </div>
  );
}

function TrashIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 20 20" className={className} fill="none" stroke="currentColor" strokeWidth="1.5">
      <path d="M4 6h12M8 6V4.5A1.5 1.5 0 0 1 9.5 3h1A1.5 1.5 0 0 1 12 4.5V6m2 0-.7 9.1a1.5 1.5 0 0 1-1.5 1.4H8.2a1.5 1.5 0 0 1-1.5-1.4L6 6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function ChevronIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 20 20" className={className} fill="none" stroke="currentColor" strokeWidth="1.5">
      <path d="M5.5 7.5 10 12l4.5-4.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
