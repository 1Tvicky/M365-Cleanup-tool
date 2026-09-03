import { useState } from "react";
import { listConnectionUsers, type ConnectionUserRow, type ManageCloudsRow } from "../../api/clouds";
import { OneDriveIcon, SharePointIcon, TeamsIcon } from "./CloudIcons";
import type { Workload } from "../../types";

const ICONS: Record<Workload, (props: { className?: string }) => JSX.Element> = {
  onedrive: OneDriveIcon,
  sharepoint: SharePointIcon,
  teams: TeamsIcon,
};

const CLOUD_LABELS: Record<Workload, string> = {
  onedrive: "OneDrive for Business",
  sharepoint: "SharePoint Online",
  teams: "Microsoft Teams",
};

const STATUS_BADGE: Partial<Record<ManageCloudsRow["status"], { label: string; style: string }>> = {
  connecting: { label: "Connecting…", style: "bg-blue-50 text-blue-700" },
  error: { label: "Error", style: "bg-rose-50 text-rose-700" },
  needs_reauth: { label: "Needs reauthorization", style: "bg-amber-50 text-amber-700" },
};

/**
 * Row layout and the expand-panel summary match the reference product exactly (verified against a
 * screen recording, not the original prompt's "per-user list" description — the chevron reveals a
 * summary + a "Failed Users Details" drill-in, not a raw table by default).
 */
export function ManageClouds({
  connections,
  onResync,
  onDisconnect,
  resyncingId,
}: {
  connections: ManageCloudsRow[];
  onResync: (id: string) => void;
  onDisconnect: (id: string) => void;
  resyncingId: string | null;
}) {
  const [expandedId, setExpandedId] = useState<string | null>(null);

  if (connections.length === 0) {
    return (
      <p className="rounded-lg border border-dashed border-slate-300 px-4 py-8 text-center text-sm text-slate-500">
        No clouds connected yet — add one from the Add Clouds tab.
      </p>
    );
  }

  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      {connections.map((c, i) => (
        <div key={c.id} className={i > 0 ? "border-t border-slate-100" : ""}>
          <ManageCloudsRowView
            row={c}
            expanded={expandedId === c.id}
            onToggleExpand={() => setExpandedId((prev) => (prev === c.id ? null : c.id))}
            onResync={() => onResync(c.id)}
            onDisconnect={() => onDisconnect(c.id)}
            resyncing={resyncingId === c.id}
          />
        </div>
      ))}
    </div>
  );
}

function ManageCloudsRowView({
  row,
  expanded,
  onToggleExpand,
  onResync,
  onDisconnect,
  resyncing,
}: {
  row: ManageCloudsRow;
  expanded: boolean;
  onToggleExpand: () => void;
  onResync: () => void;
  onDisconnect: () => void;
  resyncing: boolean;
}) {
  const Icon = ICONS[row.cloudType];
  const badge = STATUS_BADGE[row.status];
  const hasJob = row.totalUsers > 0;

  return (
    <div>
      <div className="flex flex-wrap items-center gap-x-6 gap-y-2 px-6 py-4">
        <Icon className="h-9 w-9 shrink-0" />

        <div className="min-w-[170px]">
          <div className="flex items-center gap-2">
            <span className="text-[15px] font-semibold text-slate-800">{CLOUD_LABELS[row.cloudType]}</span>
            {badge && <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${badge.style}`}>{badge.label}</span>}
          </div>
        </div>

        {hasJob && (
          <div className="min-w-[200px] flex-1">
            <div className="mb-1 text-sm text-slate-600">
              {row.addedUsers.toLocaleString()} out of {row.totalUsers.toLocaleString()} Users{" "}
              <span className="ml-1 font-semibold text-slate-800">{row.percent}%</span>
            </div>
            <div className="flex h-1.5 w-full max-w-[240px] overflow-hidden rounded-full bg-slate-100">
              <div className="h-full bg-emerald-500" style={{ width: `${(row.addedUsers / row.totalUsers) * 100}%` }} />
              <div className="h-full bg-rose-500" style={{ width: `${(row.notAddedUsers / row.totalUsers) * 100}%` }} />
            </div>
          </div>
        )}

        <button
          onClick={onResync}
          disabled={resyncing || row.status === "disconnected"}
          aria-label="Resync"
          className="text-slate-400 hover:text-slate-600 disabled:cursor-not-allowed disabled:opacity-40"
        >
          <span className={resyncing ? "inline-block animate-spin" : "inline-block"} aria-hidden>
            ⟳
          </span>
        </button>

        <span className="hidden text-sm text-slate-500 sm:inline">{row.adminEmail}</span>

        <div className="ml-auto flex items-center gap-4">
          <span className="text-sm font-medium text-slate-600">Multiuser</span>
          <button onClick={onDisconnect} aria-label="Disconnect" className="text-slate-400 hover:text-rose-600">
            <TrashIcon className="h-[18px] w-[18px]" />
          </button>
          <button
            onClick={onToggleExpand}
            aria-label="Toggle details"
            className={`flex h-8 w-8 items-center justify-center rounded-full border transition-colors ${
              expanded ? "border-[#1b2fc4] bg-[#1b2fc4] text-white" : "border-slate-200 text-slate-400 hover:text-slate-600"
            }`}
          >
            <ChevronIcon className={`h-4 w-4 transition-transform ${expanded ? "rotate-180" : ""}`} />
          </button>
        </div>
      </div>

      {expanded && <ExpandedSummary row={row} />}
    </div>
  );
}

function ExpandedSummary({ row }: { row: ManageCloudsRow }) {
  const [showFailed, setShowFailed] = useState(false);
  const [failedUsers, setFailedUsers] = useState<ConnectionUserRow[] | null>(null);
  const [loadingFailed, setLoadingFailed] = useState(false);

  async function handleShowFailed() {
    setShowFailed(true);
    if (failedUsers) return;
    setLoadingFailed(true);
    try {
      const { users } = await listConnectionUsers(row.id, { status: "failed" });
      setFailedUsers(users);
    } finally {
      setLoadingFailed(false);
    }
  }

  return (
    <div className="border-t border-slate-100 bg-slate-50 px-6 py-4">
      <div className="grid grid-cols-2 gap-x-8 gap-y-3 sm:grid-cols-4">
        <Field label="Name" value={row.adminDisplayName ?? row.adminEmail} />
        <Field label="Domain Name" value={row.tenantDomain} />
        <Field label="Total Users" value={row.totalUsers.toLocaleString()} />
        <Field label="Added Users" value={row.addedUsers.toLocaleString()} />
      </div>

      {row.notAddedUsers > 0 && (
        <div className="mt-4 flex items-center gap-4">
          <span className="rounded-md bg-rose-50 px-3 py-1.5 text-sm font-semibold text-rose-600">
            Users Not Added <span className="ml-1">{row.notAddedUsers.toLocaleString()}</span>
          </span>
          <button onClick={handleShowFailed} className="text-sm font-semibold text-[#1b2fc4] underline underline-offset-2">
            Failed Users Details
          </button>
        </div>
      )}

      {showFailed && (
        <div className="mt-4 overflow-hidden rounded-lg border border-slate-200 bg-white">
          {loadingFailed ? (
            <div className="px-4 py-3 text-sm text-slate-500">Loading…</div>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                <tr>
                  <th className="px-4 py-2 font-medium">User</th>
                  <th className="px-4 py-2 font-medium">Reason</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {(failedUsers ?? []).map((u) => (
                  <tr key={u.id}>
                    <td className="px-4 py-2 text-slate-700">{u.displayName ?? u.upn}</td>
                    <td className="px-4 py-2 text-slate-500">{u.errorMessage ?? "Unknown error"}</td>
                  </tr>
                ))}
                {failedUsers?.length === 0 && (
                  <tr>
                    <td colSpan={2} className="px-4 py-3 text-slate-400">
                      No failed users found.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          )}
        </div>
      )}

      {row.status === "needs_reauth" && (
        <p className="mt-3 text-sm text-amber-700">
          Admin consent may have been revoked or expired for this connection. Reconnect from the Add Clouds tab.
        </p>
      )}
      {row.lastSyncedAt && <p className="mt-3 text-xs text-slate-400">Last synced {formatDate(row.lastSyncedAt)}</p>}
    </div>
  );
}

function Field({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <div className="text-[11px] font-medium uppercase tracking-wide text-slate-400">{label}</div>
      <div className="text-sm text-slate-800">{value}</div>
    </div>
  );
}

function formatDate(iso: string): string {
  return new Date(iso).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
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
