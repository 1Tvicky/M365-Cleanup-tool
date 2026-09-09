import { useEffect, useState } from "react";
import {
  exportConnectionUsersUrl,
  getSyncJobResources,
  listAvailableResources,
  listConnectionUsers,
  resyncCloudConnection,
  type AvailableResourceRow,
  type ConnectionUserRow,
  type ManageCloudsRow,
  type SyncJobResourceRow,
} from "../../api/clouds";
import { ApiClientError } from "../../api/client";
import { DiscoveryTable, useDebouncedValue, type DiscoveryColumn } from "../cleaning/DiscoveryTable";
import { Spinner } from "../cleaning/ItemFilesDrilldown";
import { GoogleIcon, OneDriveIcon, OutlookIcon, SharePointIcon, TeamsIcon } from "./CloudIcons";
import type { Workload } from "../../types";
import { formatBytes } from "../../utils/format";

const USERS_PAGE_SIZE = 50;

const ICONS: Record<Workload, (props: { className?: string }) => JSX.Element> = {
  onedrive: OneDriveIcon,
  sharepoint: SharePointIcon,
  teams: TeamsIcon,
  outlook: OutlookIcon,
  google_my_drive: GoogleIcon,
};

const CLOUD_LABELS: Record<Workload, string> = {
  onedrive: "OneDrive for Business",
  sharepoint: "SharePoint Online",
  teams: "Microsoft Teams",
  outlook: "Outlook",
  google_my_drive: "Google My Drive",
};

// SharePoint enumerates sites, not people — a tenant with a hundred users can easily have a
// thousand+ sites (one per team/group, communication sites, hub sites, etc.), so labeling that
// count "Users" the same way OneDrive does is misleading, not just cosmetically wrong. Teams synced
// to mean actual Teams (not per-user joined-team counts — see migrations/012_sync_job_resources.sql)
// for the same reason: "Users" would now be wrong there too.
const UNIT_LABELS: Record<Workload, { singular: string; plural: string }> = {
  onedrive: { singular: "User", plural: "Users" },
  teams: { singular: "Team", plural: "Teams" },
  sharepoint: { singular: "Site", plural: "Sites" },
  outlook: { singular: "Mailbox", plural: "Mailboxes" },
  google_my_drive: { singular: "User", plural: "Users" },
};

// The secondary identifier shown alongside a resource's name in the Sync Resources picker — null
// where the workload has nothing meaningful to show there (a Team has no email/URL equivalent).
const SECONDARY_COLUMN_LABEL: Record<Workload, string | null> = {
  onedrive: "Email",
  outlook: "Email",
  sharepoint: "URL",
  teams: null,
  google_my_drive: "Email",
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
  const unit = UNIT_LABELS[row.cloudType];
  const [confirmingDisconnect, setConfirmingDisconnect] = useState(false);

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
              {row.addedUsers.toLocaleString()} out of {row.totalUsers.toLocaleString()} {unit.plural}{" "}
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
          <button onClick={() => setConfirmingDisconnect(true)} aria-label="Disconnect" className="text-slate-400 hover:text-rose-600">
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

      {expanded && <ExpandedPanel row={row} />}

      {confirmingDisconnect && (
        <DisconnectConfirmModal
          row={row}
          onCancel={() => setConfirmingDisconnect(false)}
          onConfirm={() => {
            setConfirmingDisconnect(false);
            onDisconnect();
          }}
        />
      )}
    </div>
  );
}

/**
 * Disconnecting was previously one click on the trash icon with no way back — easy to trigger by
 * accident while trying to verify something else in the row (e.g. the resync icon sits right next
 * to it). Requires an explicit confirm now; Cancel (or clicking the backdrop) leaves the connection
 * untouched.
 */
function DisconnectConfirmModal({ row, onCancel, onConfirm }: { row: ManageCloudsRow; onCancel: () => void; onConfirm: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4" onClick={onCancel}>
      <div className="w-full max-w-sm rounded-xl bg-white p-6 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <h3 className="text-base font-semibold text-slate-900">Disconnect {CLOUD_LABELS[row.cloudType]}?</h3>
        <p className="mt-2 text-sm text-slate-500">
          This removes the <span className="font-medium text-slate-700">{row.tenantDomain}</span> {CLOUD_LABELS[row.cloudType]}{" "}
          connection from CloudFuze — discovery and cleanup for it won't be available here until it's reconnected. This doesn't
          revoke Microsoft's own consent grant; do that from the tenant's Enterprise Applications page if you want to fully
          remove access.
        </p>
        <div className="mt-5 flex justify-end gap-3">
          <button onClick={onCancel} className="rounded-md border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button onClick={onConfirm} className="rounded-md bg-rose-600 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-700">
            Disconnect
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * Two tabs behind the expand chevron: "Synced Resources" (the existing read-only, already-synced
 * view — unchanged) and "Sync Resources" (new — browse every resource the workload currently has,
 * select a subset, sync only those). Reuses the same expand/collapse mechanism already in place
 * rather than introducing a new page or route for the new flow.
 */
function ExpandedPanel({ row }: { row: ManageCloudsRow }) {
  const [tab, setTab] = useState<"synced" | "sync">("synced");
  return (
    <div className="border-t border-slate-100 bg-slate-50">
      <div className="flex gap-1 px-6 pt-3">
        <ExpandedTabButton label="Synced Resources" active={tab === "synced"} onClick={() => setTab("synced")} />
        <ExpandedTabButton label="Sync Resources" active={tab === "sync"} onClick={() => setTab("sync")} />
      </div>
      {tab === "synced" ? <ExpandedSummary row={row} /> : <SyncResourcePicker connectionId={row.id} cloudType={row.cloudType} />}
    </div>
  );
}

function ExpandedTabButton({ label, active, onClick }: { label: string; active: boolean; onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className={`rounded-t-md px-3 py-1.5 text-sm font-medium transition-colors ${
        active ? "bg-white text-[#1b2fc4]" : "text-slate-500 hover:text-slate-700"
      }`}
    >
      {label}
    </button>
  );
}

/**
 * Total/Active/In-Active tiles + the full per-user table (not gated behind a "show failed" click —
 * every row, every status, always visible once expanded) + a CSV export of the whole list. "Active"
 * here means this app's own last sync actually found and read real data for that user
 * (sync_status='synced') — not Azure AD's accountEnabled, which would still count guests/unlicensed
 * accounts as "active" and reproduce the same over-counting this was built to fix (see the "395
 * mailboxes" investigation this session — most were guest/unlicensed accounts with no real mailbox).
 */
function ExpandedSummary({ row }: { row: ManageCloudsRow }) {
  const unit = UNIT_LABELS[row.cloudType];
  const [cursorStack, setCursorStack] = useState<(string | null)[]>([null]);
  const [pageIndex, setPageIndex] = useState(0);
  const [users, setUsers] = useState<ConnectionUserRow[] | null>(null);
  const [nextCursor, setNextCursor] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    listConnectionUsers(row.id, { cursor: cursorStack[pageIndex] ?? undefined, limit: USERS_PAGE_SIZE })
      .then(({ users: rows, nextCursor: next }) => {
        if (cancelled) return;
        setUsers(rows);
        setNextCursor(next);
      })
      .catch(() => {
        if (!cancelled) setError("Couldn't load the user list. Try again.");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [row.id, cursorStack, pageIndex]);

  return (
    <div className="px-6 py-4">
      <div className="grid grid-cols-3 gap-4">
        <StatTile label={`Total ${unit.plural}`} value={row.totalUsers} />
        <StatTile label={`Active ${unit.plural}`} value={row.addedUsers} />
        <StatTile label={`In-Active ${unit.plural}`} value={row.notAddedUsers} />
      </div>

      <div className="mt-4 overflow-hidden rounded-lg border border-slate-200 bg-white">
        <div className="flex items-center justify-end border-b border-slate-100 px-4 py-2">
          <a
            href={exportConnectionUsersUrl(row.id)}
            className="flex items-center gap-1.5 text-sm font-semibold text-[#1b2fc4] hover:opacity-80"
          >
            <DownloadIcon className="h-4 w-4" /> CSV
          </a>
        </div>

        {error ? (
          <div className="px-4 py-3 text-sm text-rose-600">{error}</div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
              <tr>
                <th className="px-4 py-2 font-medium">S.No</th>
                <th className="px-4 py-2 font-medium">Name</th>
                <th className="px-4 py-2 font-medium">Email</th>
                <th className="px-4 py-2 font-medium">{unit.singular} Size</th>
                <th className="px-4 py-2 font-medium">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan={5} className="px-4 py-3 text-slate-400">
                    Loading…
                  </td>
                </tr>
              ) : users && users.length > 0 ? (
                users.map((u, i) => (
                  <tr key={u.id}>
                    <td className="px-4 py-2 text-slate-500">{pageIndex * USERS_PAGE_SIZE + i + 1}</td>
                    <td className="px-4 py-2 text-slate-700">{u.displayName ?? u.upn}</td>
                    <td className="px-4 py-2 text-slate-500">{u.upn}</td>
                    <td className="px-4 py-2 text-slate-500">{u.storageUsedBytes > 0 ? formatBytes(u.storageUsedBytes) : "-"}</td>
                    <td className="px-4 py-2">
                      <UserStatusBadge syncStatus={u.syncStatus} />
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} className="px-4 py-3 text-slate-400">
                    No {unit.plural.toLowerCase()} found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}

        {(pageIndex > 0 || nextCursor) && (
          <div className="flex items-center justify-between border-t border-slate-100 px-4 py-2 text-sm">
            <button
              onClick={() => setPageIndex((p) => p - 1)}
              disabled={pageIndex === 0}
              className="font-medium text-slate-500 hover:text-slate-700 disabled:cursor-not-allowed disabled:opacity-40"
            >
              ← Previous
            </button>
            <button
              onClick={() => {
                setCursorStack((prev) => [...prev.slice(0, pageIndex + 1), nextCursor]);
                setPageIndex((p) => p + 1);
              }}
              disabled={!nextCursor}
              className="font-medium text-slate-500 hover:text-slate-700 disabled:cursor-not-allowed disabled:opacity-40"
            >
              Next →
            </button>
          </div>
        )}
      </div>

      {row.status === "needs_reauth" && (
        <p className="mt-3 text-sm text-amber-700">
          Admin consent may have been revoked or expired for this connection. Reconnect from the Add Clouds tab.
        </p>
      )}
      {row.lastSyncedAt && <p className="mt-3 text-xs text-slate-400">Last synced {formatDate(row.lastSyncedAt)}</p>}
    </div>
  );
}

const RESOURCE_PAGE_SIZE = 20;

/**
 * The new resource-level sync flow: browse every resource this workload currently has (live, via
 * GET /:id/available-resources — not connection_users, which can be empty/stale pre-sync), select a
 * subset, "Sync Selected". Reuses DiscoveryTable exactly as the Cleaning module's own discovery
 * tables do (search, select-all, per-row checkbox, page-number pagination) — "Select All" here only
 * ever toggles the currently-fetched page's rows, the same as every other DiscoveryTable caller,
 * never a separate "select every resource in the tenant" action.
 */
function SyncResourcePicker({ connectionId, cloudType }: { connectionId: string; cloudType: Workload }) {
  const [page, setPage] = useState(1);
  const [searchInput, setSearchInput] = useState("");
  const search = useDebouncedValue(searchInput, 400);
  const [rows, setRows] = useState<AvailableResourceRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [starting, setStarting] = useState(false);
  const [startError, setStartError] = useState<string | null>(null);
  const [activeJobId, setActiveJobId] = useState<string | null>(null);

  // A search change always starts back at page 1 — a stale page number from a differently-filtered
  // list wouldn't make sense against the new one (same pattern as ReportsPage/CategoryItemsList).
  useEffect(() => {
    setPage(1);
  }, [search]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    listAvailableResources(connectionId, { search, page, pageSize: RESOURCE_PAGE_SIZE })
      .then((res) => {
        if (cancelled) return;
        setRows(res.resources);
        setTotal(res.total);
        setError(null);
      })
      .catch((err) => {
        if (!cancelled) setError(err instanceof ApiClientError ? err.message : "Couldn't load resources.");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [connectionId, search, page]);

  function toggle(id: string) {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleAll() {
    setSelected((prev) => {
      const next = new Set(prev);
      const allSelected = rows.length > 0 && rows.every((r) => next.has(r.id));
      for (const r of rows) {
        if (allSelected) next.delete(r.id);
        else next.add(r.id);
      }
      return next;
    });
  }

  async function handleSyncSelected() {
    setStarting(true);
    setStartError(null);
    try {
      const result = await resyncCloudConnection(connectionId, [...selected]);
      setActiveJobId(result.jobId);
      setSelected(new Set());
      if (result.skipped.length > 0) {
        setStartError(`${result.skipped.length} selected ${result.skipped.length === 1 ? "resource" : "resources"} couldn't be synced (already syncing, or no longer found).`);
      }
    } catch (err) {
      setStartError(err instanceof ApiClientError ? err.message : "Couldn't start sync.");
    } finally {
      setStarting(false);
    }
  }

  const unit = UNIT_LABELS[cloudType];
  const secondaryLabel = SECONDARY_COLUMN_LABEL[cloudType];
  const totalPages = Math.max(1, Math.ceil(total / RESOURCE_PAGE_SIZE));

  const columns: DiscoveryColumn<AvailableResourceRow>[] = [
    { label: `${unit.singular} Name`, render: (r) => r.displayName },
    ...(secondaryLabel ? [{ label: secondaryLabel, render: (r: AvailableResourceRow) => r.secondary ?? "—" }] : []),
  ];

  return (
    <div className="px-6 py-4">
      {activeJobId && <SyncJobProgress connectionId={connectionId} jobId={activeJobId} onDone={() => setActiveJobId(null)} />}

      <DiscoveryTable<AvailableResourceRow>
        title={`Select ${unit.plural.toLowerCase()} to sync`}
        columns={columns}
        rows={rows}
        loading={loading}
        error={error}
        page={page}
        totalPages={totalPages}
        total={total}
        onGoToPage={setPage}
        search={searchInput}
        onSearchChange={setSearchInput}
        searchPlaceholder={`Search ${unit.plural.toLowerCase()}…`}
        selected={selected}
        onToggle={toggle}
        onToggleAll={toggleAll}
        emptyMessage={`No ${unit.plural.toLowerCase()} found.`}
      />

      <div className="mt-3 flex items-center justify-between">
        <span className="text-sm font-medium text-slate-600">Selected: {selected.size}</span>
        <button
          onClick={handleSyncSelected}
          disabled={selected.size === 0 || starting}
          className="rounded-md bg-[#1b2fc4] px-4 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {starting ? "Starting…" : "Sync Selected"}
        </button>
      </div>
      {startError && <p className="mt-2 text-sm text-rose-600">{startError}</p>}
    </div>
  );
}

const RESOURCE_STATUS_STYLE: Record<SyncJobResourceRow["status"], { label: string; className: string }> = {
  pending: { label: "Pending", className: "text-slate-400" },
  processing: { label: "Processing", className: "text-blue-600" },
  completed: { label: "Completed", className: "text-emerald-600" },
  failed: { label: "Failed", className: "text-rose-600" },
  cancelled: { label: "Cancelled", className: "text-slate-500" },
};

const TERMINAL_RESOURCE_STATUSES = new Set<SyncJobResourceRow["status"]>(["completed", "failed", "cancelled"]);

/**
 * Live per-resource progress for the run just started — not just one aggregate bar, since the whole
 * point of resource-level sync is that "65% done" means nothing once the operator picked exactly
 * which resources they care about. Polls every 3s (same cadence established elsewhere this session
 * for live progress) until every resource has settled one way or another.
 */
function SyncJobProgress({ connectionId, jobId, onDone }: { connectionId: string; jobId: string; onDone: () => void }) {
  const [resources, setResources] = useState<SyncJobResourceRow[] | null>(null);

  useEffect(() => {
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const data = await getSyncJobResources(connectionId, jobId).catch(() => null);
      if (cancelled || !data) return;
      setResources(data.resources);
      const allDone = data.resources.length > 0 && data.resources.every((r) => TERMINAL_RESOURCE_STATUSES.has(r.status));
      if (!allDone) timer = setTimeout(poll, 3000);
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [connectionId, jobId]);

  if (!resources) return <p className="mb-4 text-sm text-slate-500">Starting sync…</p>;

  const completedCount = resources.filter((r) => TERMINAL_RESOURCE_STATUSES.has(r.status)).length;
  const allDone = resources.length > 0 && completedCount === resources.length;

  return (
    <div className="mb-4 overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-center justify-between border-b border-slate-100 px-4 py-2.5 text-sm">
        <span className="flex items-center gap-1.5 font-medium text-slate-700">
          {!allDone && <Spinner className="h-3.5 w-3.5" />}
          {completedCount} of {resources.length} completed
        </span>
        {allDone && (
          <button onClick={onDone} className="font-medium text-[#1b2fc4] hover:underline">
            Dismiss
          </button>
        )}
      </div>
      <ul className="max-h-56 divide-y divide-slate-100 overflow-y-auto text-sm">
        {resources.map((r) => (
          <li key={r.id} className="flex items-center justify-between px-4 py-2">
            <span className="text-slate-700">{r.displayName}</span>
            <span className={`flex items-center gap-1.5 text-xs font-medium ${RESOURCE_STATUS_STYLE[r.status].className}`}>
              {(r.status === "pending" || r.status === "processing") && <Spinner className="h-3 w-3" />}
              {RESOURCE_STATUS_STYLE[r.status].label}
            </span>
          </li>
        ))}
      </ul>
    </div>
  );
}

function StatTile({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white px-4 py-3">
      <div className="text-sm text-slate-500">{label}</div>
      <div className="mt-1 text-2xl font-bold text-slate-900">{value.toLocaleString()}</div>
    </div>
  );
}

function UserStatusBadge({ syncStatus }: { syncStatus: ConnectionUserRow["syncStatus"] }) {
  if (syncStatus === "synced") return <span className="text-emerald-600">Active</span>;
  if (syncStatus === "failed") return <span className="text-rose-500">Inactive</span>;
  return <span className="italic text-slate-400">Pending</span>;
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

function DownloadIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 20 20" className={className} fill="none" stroke="currentColor" strokeWidth="1.5">
      <path d="M10 3v9m0 0-3.5-3.5M10 12l3.5-3.5M4 14.5v1a1.5 1.5 0 0 0 1.5 1.5h9a1.5 1.5 0 0 0 1.5-1.5v-1" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
