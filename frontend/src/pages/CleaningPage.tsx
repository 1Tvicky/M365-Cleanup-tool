import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  calculateTeamsMessageCounts,
  getTeamsSummary,
  listCleaningConnections,
  listOneDriveAccounts,
  listSharePointSites,
  listTeamsChannels,
  listTeamsDMs,
  type CleaningChannelRow,
  type CleaningChatRow,
  type CleaningConnectionRow,
  type CleaningResourceRow,
  type CleaningTeamsSummary,
  type PageResult,
} from "../api/cleaning";
import { ApiClientError } from "../api/client";
import { DiscoveryTable, useDebouncedValue, type DiscoveryColumn } from "../components/cleaning/DiscoveryTable";
import { TeamsChannels } from "../components/cleaning/TeamsChannels";
import { SelectionSummary, hasSelection, messagesFragment, type SelectionTotals } from "../components/cleaning/SelectionSummary";
import { formatBytes, formatDate } from "../utils/format";

interface TenantGroup {
  domain: string;
  adminEmail: string;
  adminDisplayName: string | null;
  status: CleaningConnectionRow["status"];
  lastSyncedAt: string | null;
  onedrive?: CleaningConnectionRow;
  sharepoint?: CleaningConnectionRow;
  teams?: CleaningConnectionRow;
}

type View = "landing" | "dashboard" | "onedrive" | "sharepoint" | "teams" | "review";

/**
 * Distinguishes "still working on it" from "gave up" from "genuinely counted zero" — conflating
 * these (as an earlier version did) meant a fully-failed connection (e.g. ChannelMessage.Read.All
 * not yet granted) displayed a literal fake "0 messages" instead of an honest status.
 */
function teamsMessagesLabel(summary: CleaningTeamsSummary): { text: string; className: string } {
  if (summary.itemsAwaitingCount > 0) {
    const discovering = summary.structureScan?.status === "running" || summary.structureScan?.status === "queued";
    return { text: discovering ? "Discovering teams…" : "Waiting to be calculated", className: "italic text-slate-400" };
  }
  if (summary.itemsFailedCount > 0 && summary.messagesCountedSoFar === 0) {
    return { text: "Unable to calculate", className: "text-rose-500" };
  }
  if (summary.itemsFailedCount > 0) {
    return { text: `${summary.messagesCountedSoFar.toLocaleString()} messages (some unavailable)`, className: "text-amber-600" };
  }
  return { text: `${summary.messagesCountedSoFar.toLocaleString()} messages`, className: "" };
}

const STATUS_LABEL: Record<CleaningConnectionRow["status"], { label: string; style: string }> = {
  active: { label: "Connected", style: "bg-emerald-50 text-emerald-700" },
  connecting: { label: "Connecting…", style: "bg-blue-50 text-blue-700" },
  error: { label: "Needs attention", style: "bg-rose-50 text-rose-700" },
  needs_reauth: { label: "Needs reconnection", style: "bg-amber-50 text-amber-700" },
  disconnected: { label: "Disconnected", style: "bg-slate-100 text-slate-500" },
};

export function CleaningPage() {
  const [view, setView] = useState<View>("landing");
  const [groups, setGroups] = useState<TenantGroup[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);
  const [activeGroup, setActiveGroup] = useState<TenantGroup | null>(null);
  const [teamsTab, setTeamsTab] = useState<"channels" | "dms">("channels");

  const [selectedOneDrive, setSelectedOneDrive] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedSharePoint, setSelectedSharePoint] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedChannels, setSelectedChannels] = useState<Map<string, CleaningChannelRow>>(new Map());
  const [selectedChats, setSelectedChats] = useState<Map<string, CleaningChatRow>>(new Map());

  useEffect(() => {
    listCleaningConnections()
      .then(({ connections }) => {
        const byDomain = new Map<string, TenantGroup>();
        for (const c of connections) {
          if (!byDomain.has(c.displayName)) {
            byDomain.set(c.displayName, {
              domain: c.displayName,
              adminEmail: c.adminEmail,
              adminDisplayName: c.adminDisplayName,
              status: c.status,
              lastSyncedAt: c.lastSyncedAt,
            });
          }
          const group = byDomain.get(c.displayName)!;
          group[c.cloudType] = c;
          if (c.lastSyncedAt && (!group.lastSyncedAt || c.lastSyncedAt > group.lastSyncedAt)) group.lastSyncedAt = c.lastSyncedAt;
        }
        setGroups([...byDomain.values()]);
      })
      .finally(() => setLoadingGroups(false));
  }, []);

  const totals: SelectionTotals = useMemo(
    () => ({
      oneDriveAccounts: selectedOneDrive.size,
      oneDriveBytes: [...selectedOneDrive.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
      sharePointSites: selectedSharePoint.size,
      sharePointBytes: [...selectedSharePoint.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
      teamsChannels: selectedChannels.size,
      teamsMessages: [...selectedChannels.values()].reduce((s, r) => s + (r.countStatus === "completed" ? r.messageCount ?? 0 : 0), 0),
      teamsChannelsWithKnownCount: [...selectedChannels.values()].filter((r) => r.countStatus === "completed").length,
      dms: selectedChats.size,
      dmMessages: [...selectedChats.values()].reduce((s, r) => s + (r.countStatus === "completed" ? r.messageCount ?? 0 : 0), 0),
      dmsWithKnownCount: [...selectedChats.values()].filter((r) => r.countStatus === "completed").length,
    }),
    [selectedOneDrive, selectedSharePoint, selectedChannels, selectedChats]
  );

  function openTenant(group: TenantGroup) {
    // Selections are scoped to one Microsoft 365 cloud at a time — switching tenants without
    // clearing them would let a review mix data from two different customers' tenants together.
    setSelectedOneDrive(new Map());
    setSelectedSharePoint(new Map());
    setSelectedChannels(new Map());
    setSelectedChats(new Map());
    setActiveGroup(group);
    setView("dashboard");
  }

  if (view === "review") {
    return <ReviewPage totals={totals} onBack={() => setView("dashboard")} />;
  }

  const showSelectionBar = view !== "landing" && hasSelection(totals);

  return (
    <div className={`px-8 py-6 ${showSelectionBar ? "pb-24" : ""}`}>
      <div className="pb-6">
        {view !== "landing" && (
          <button
            onClick={() => (view === "dashboard" ? (setView("landing"), setActiveGroup(null)) : setView("dashboard"))}
            className="mb-4 text-sm font-medium text-slate-500 hover:text-slate-700"
          >
            ← {view === "dashboard" ? "All clouds" : "Dashboard"}
          </button>
        )}

        {view === "landing" && (
          <Landing groups={groups} loading={loadingGroups} onOpen={openTenant} />
        )}

        {view === "dashboard" && activeGroup && (
          <Dashboard
            group={activeGroup}
            onOpenOneDrive={() => setView("onedrive")}
            onOpenSharePoint={() => setView("sharepoint")}
            onOpenTeams={() => setView("teams")}
          />
        )}

        {view === "onedrive" && activeGroup?.onedrive && (
          <OneDriveView connectionId={activeGroup.onedrive.id} selected={selectedOneDrive} setSelected={setSelectedOneDrive} />
        )}

        {view === "sharepoint" && activeGroup?.sharepoint && (
          <SharePointView connectionId={activeGroup.sharepoint.id} selected={selectedSharePoint} setSelected={setSelectedSharePoint} />
        )}

        {view === "teams" && activeGroup?.teams && (
          <TeamsView
            connectionId={activeGroup.teams.id}
            tab={teamsTab}
            onTabChange={setTeamsTab}
            selectedChannels={selectedChannels}
            setSelectedChannels={setSelectedChannels}
            selectedChats={selectedChats}
            setSelectedChats={setSelectedChats}
          />
        )}
      </div>

      {showSelectionBar && <SelectionSummary totals={totals} onReview={() => setView("review")} />}
    </div>
  );
}

function Landing({ groups, loading, onOpen }: { groups: TenantGroup[]; loading: boolean; onOpen: (g: TenantGroup) => void }) {
  if (loading) return <p className="text-sm text-slate-500">Loading…</p>;
  if (groups.length === 0) {
    return (
      <p className="rounded-lg border border-dashed border-slate-300 px-4 py-8 text-center text-sm text-slate-500">
        No Microsoft 365 clouds connected yet — connect one from the Clouds tab first.
      </p>
    );
  }
  return (
    <div>
      <h2 className="mb-5 text-base font-semibold text-slate-800">Cleaning</h2>
      <div className="flex flex-wrap gap-4">
        {groups.map((g) => {
          const badge = STATUS_LABEL[g.status];
          return (
            <div key={g.domain} className="w-72 rounded-xl border border-slate-200 bg-white p-5">
              <div className="mb-3 flex items-center justify-between">
                <span className="text-sm font-semibold text-slate-800">Microsoft 365</span>
                <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${badge.style}`}>{badge.label}</span>
              </div>
              <div className="space-y-1 text-sm text-slate-600">
                <div>👤 {g.adminDisplayName ?? g.adminEmail}</div>
                <div>🌐 {g.domain}</div>
              </div>
              {g.lastSyncedAt && <div className="mt-2 text-xs text-slate-400">Last updated {formatDate(g.lastSyncedAt)}</div>}
              <button
                onClick={() => onOpen(g)}
                className="mt-4 w-full rounded-md bg-[#1b2fc4] py-2 text-sm font-semibold text-white hover:opacity-90"
              >
                Open Cleaning
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ServiceCard({ icon, name, stats, action, onClick, disabled }: { icon: string; name: string; stats: React.ReactNode; action: string; onClick: () => void; disabled?: boolean }) {
  return (
    <div className="w-64 rounded-xl border border-slate-200 bg-white p-5">
      <div className="mb-3 flex items-center gap-2">
        <span className="text-xl" aria-hidden>{icon}</span>
        <span className="text-sm font-semibold text-slate-800">{name}</span>
      </div>
      <div className="space-y-1 text-sm text-slate-600">{stats}</div>
      <button
        onClick={onClick}
        disabled={disabled}
        className="mt-4 w-full rounded-md border border-[#1b2fc4] py-1.5 text-sm font-semibold text-[#1b2fc4] hover:bg-[#1b2fc4]/5 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400"
      >
        {action}
      </button>
    </div>
  );
}

function Dashboard({ group, onOpenOneDrive, onOpenSharePoint, onOpenTeams }: { group: TenantGroup; onOpenOneDrive: () => void; onOpenSharePoint: () => void; onOpenTeams: () => void }) {
  const [oneDriveTotals, setOneDriveTotals] = useState<{ count: number; bytes: number } | null>(null);
  const [sharePointTotals, setSharePointTotals] = useState<{ count: number; bytes: number } | null>(null);
  const [teamsSummary, setTeamsSummary] = useState<CleaningTeamsSummary | null>(null);

  useEffect(() => {
    if (group.onedrive) {
      listOneDriveAccounts(group.onedrive.id, { sort: "storage", pageSize: 200 }).then(({ accounts, total }) => {
        // `total` is the accurate count; the byte sum is over the (large) page fetched here, which
        // covers the dashboard headline for realistically-sized tenants — the full table (with real
        // pagination) is what "View Accounts" opens.
        setOneDriveTotals({ count: total, bytes: accounts.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
    if (group.sharepoint) {
      listSharePointSites(group.sharepoint.id, { sort: "storage", pageSize: 200 }).then(({ sites, total }) => {
        setSharePointTotals({ count: total, bytes: sites.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
  }, [group]);

  useEffect(() => {
    if (!group.teams) return;
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const summary = await getTeamsSummary(group.teams!.id).catch(() => null);
      if (cancelled || !summary) return;
      setTeamsSummary(summary);
      const live = summary.structureScan?.status === "running" || summary.structureScan?.status === "queued" ||
        summary.countScan?.status === "running" || summary.countScan?.status === "queued";
      if (live) timer = setTimeout(poll, 3000);
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [group]);

  return (
    <div>
      <h2 className="mb-5 text-base font-semibold text-slate-800">{group.domain}</h2>
      <div className="flex flex-wrap gap-4">
        <ServiceCard
          icon="☁️"
          name="OneDrive"
          stats={
            group.onedrive ? (
              oneDriveTotals ? (
                <>
                  <div>{oneDriveTotals.count.toLocaleString()}{oneDriveTotals.count > 0 ? "+" : ""} accounts</div>
                  <div>{formatBytes(oneDriveTotals.bytes)} used</div>
                </>
              ) : (
                <div className="italic text-slate-400">Loading…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Accounts"
          onClick={onOpenOneDrive}
          disabled={!group.onedrive}
        />
        <ServiceCard
          icon="📁"
          name="SharePoint"
          stats={
            group.sharepoint ? (
              sharePointTotals ? (
                <>
                  <div>{sharePointTotals.count.toLocaleString()}{sharePointTotals.count > 0 ? "+" : ""} sites</div>
                  <div>{formatBytes(sharePointTotals.bytes)} used</div>
                </>
              ) : (
                <div className="italic text-slate-400">Loading…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Sites"
          onClick={onOpenSharePoint}
          disabled={!group.sharepoint}
        />
        <ServiceCard
          icon="👥"
          name="Microsoft Teams"
          stats={
            group.teams ? (
              teamsSummary ? (
                <>
                  <div>{teamsSummary.teamCount.toLocaleString()} teams</div>
                  <div>{teamsSummary.channelCount.toLocaleString()} channels</div>
                  <div>{teamsSummary.chatCount.toLocaleString()} direct messages</div>
                  <div className={teamsMessagesLabel(teamsSummary).className}>{teamsMessagesLabel(teamsSummary).text}</div>
                </>
              ) : (
                <div className="italic text-slate-400">Discovering teams…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Teams"
          onClick={onOpenTeams}
          disabled={!group.teams}
        />
      </div>
    </div>
  );
}

/**
 * Real page-by-page navigation (Prev/Back included) over the backend's keyset (cursor) pagination,
 * which is itself forward-only — going "back" means re-using a cursor we've already seen rather
 * than asking the server for one. Each page's rows and the cursor that produced the NEXT page are
 * cached by page index, so Previous is instant (no re-fetch) and only Next past the last-seen page
 * hits the network. Search/sort changes reset everything back to page 1.
 */
const PAGE_SIZE = 20;

function usePagedList<T extends { id: string }>(
  fetcher: (opts: { search?: string; sort?: "storage" | "name"; page?: number; pageSize?: number }) => Promise<{ rows: T[] } & PageResult<T>>,
  search: string,
  sort?: "storage" | "name"
) {
  const [page, setPage] = useState(1);
  const [rows, setRows] = useState<T[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const debouncedSearch = useDebouncedValue(search, 300);
  const requestId = useRef(0);

  const load = useCallback(
    (targetPage: number) => {
      const id = ++requestId.current;
      setLoading(true);
      setError(null);
      fetcher({ search: debouncedSearch, sort, page: targetPage, pageSize: PAGE_SIZE })
        .then((res) => {
          if (id !== requestId.current) return;
          setRows(res.rows);
          setTotal(res.total);
        })
        .catch((err) => {
          if (id !== requestId.current) return;
          setError(err instanceof ApiClientError ? err.message : "Couldn't load this data. Try again.");
        })
        .finally(() => {
          if (id === requestId.current) setLoading(false);
        });
    },
    [debouncedSearch, sort, fetcher]
  );

  // Search/sort changed — start over from page 1.
  useEffect(() => {
    setPage(1);
    load(1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [debouncedSearch, sort]);

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));

  const goToPage = useCallback(
    (target: number) => {
      const clamped = Math.min(Math.max(1, target), totalPages);
      setPage(clamped);
      load(clamped);
    },
    [totalPages, load]
  );

  return { rows, loading, error, page, totalPages, total, goToPage };
}

function toggleInMap<T>(map: Map<string, T>, setMap: (m: Map<string, T>) => void, id: string, row: T) {
  const next = new Map(map);
  next.has(id) ? next.delete(id) : next.set(id, row);
  setMap(next);
}

function OneDriveView({ connectionId, selected, setSelected }: { connectionId: string; selected: Map<string, CleaningResourceRow>; setSelected: (m: Map<string, CleaningResourceRow>) => void }) {
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState<"storage" | "name">("storage");
  const fetcher = useCallback(
    (opts: { search?: string; sort?: "storage" | "name"; page?: number; pageSize?: number }) =>
      listOneDriveAccounts(connectionId, opts).then((r) => ({ rows: r.accounts, total: r.total, page: r.page, pageSize: r.pageSize })),
    [connectionId]
  );
  const { rows, loading, error, page, totalPages, total, goToPage } = usePagedList(fetcher, search, sort);

  const columns: DiscoveryColumn<CleaningResourceRow>[] = [
    { label: "User Name", render: (r) => r.name },
    { label: "User Email", render: (r) => r.detail },
    { label: "Storage Used", align: "right", render: (r) => formatBytes(r.storageUsedBytes) },
    { label: "Status", render: (r) => (r.status === "failed" ? <span className="text-rose-500">Unavailable</span> : "Ready") },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">OneDrive</h2>
      <p className="mb-5 text-sm text-slate-500">All user accounts and their storage usage</p>
      <DiscoveryTable
        title="OneDrive Accounts"
        columns={columns}
        rows={rows}
        loading={loading}
        error={error}
        page={page}
        totalPages={totalPages}
        total={total}
        onGoToPage={goToPage}
        search={search}
        onSearchChange={setSearch}
        searchPlaceholder="Search users…"
        sortOptions={[{ value: "storage", label: "Storage" }, { value: "name", label: "Name" }]}
        sort={sort}
        onSortChange={(v) => setSort(v as "storage" | "name")}
        selected={new Set(selected.keys())}
        onToggle={(id) => {
          const row = rows.find((r) => r.id === id);
          if (row) toggleInMap(selected, setSelected, id, row);
        }}
        onToggleAll={() => {
          const allSelected = rows.every((r) => selected.has(r.id));
          const next = new Map(selected);
          for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
          setSelected(next);
        }}
        emptyMessage="No OneDrive accounts found."
      />
    </div>
  );
}

function SharePointView({ connectionId, selected, setSelected }: { connectionId: string; selected: Map<string, CleaningResourceRow>; setSelected: (m: Map<string, CleaningResourceRow>) => void }) {
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState<"storage" | "name">("storage");
  const fetcher = useCallback(
    (opts: { search?: string; sort?: "storage" | "name"; page?: number; pageSize?: number }) =>
      listSharePointSites(connectionId, opts).then((r) => ({ rows: r.sites, total: r.total, page: r.page, pageSize: r.pageSize })),
    [connectionId]
  );
  const { rows, loading, error, page, totalPages, total, goToPage } = usePagedList(fetcher, search, sort);

  const columns: DiscoveryColumn<CleaningResourceRow>[] = [
    { label: "Site Name", render: (r) => r.name },
    { label: "Site URL", render: (r) => <span className="text-xs text-slate-500">{r.detail}</span> },
    { label: "Storage Used", align: "right", render: (r) => formatBytes(r.storageUsedBytes) },
    { label: "Status", render: (r) => (r.status === "failed" ? <span className="text-rose-500">Unavailable</span> : "Ready") },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">SharePoint</h2>
      <p className="mb-5 text-sm text-slate-500">All sites and their storage usage</p>
      <DiscoveryTable
        title="SharePoint Sites"
        columns={columns}
        rows={rows}
        loading={loading}
        error={error}
        page={page}
        totalPages={totalPages}
        total={total}
        onGoToPage={goToPage}
        search={search}
        onSearchChange={setSearch}
        searchPlaceholder="Search sites…"
        sortOptions={[{ value: "storage", label: "Storage" }, { value: "name", label: "Name" }]}
        sort={sort}
        onSortChange={(v) => setSort(v as "storage" | "name")}
        selected={new Set(selected.keys())}
        onToggle={(id) => {
          const row = rows.find((r) => r.id === id);
          if (row) toggleInMap(selected, setSelected, id, row);
        }}
        onToggleAll={() => {
          const allSelected = rows.every((r) => selected.has(r.id));
          const next = new Map(selected);
          for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
          setSelected(next);
        }}
        emptyMessage="No SharePoint sites found."
      />
    </div>
  );
}

function TeamsView({
  connectionId,
  tab,
  onTabChange,
  selectedChannels,
  setSelectedChannels,
  selectedChats,
  setSelectedChats,
}: {
  connectionId: string;
  tab: "channels" | "dms";
  onTabChange: (t: "channels" | "dms") => void;
  selectedChannels: Map<string, CleaningChannelRow>;
  setSelectedChannels: (m: Map<string, CleaningChannelRow>) => void;
  selectedChats: Map<string, CleaningChatRow>;
  setSelectedChats: (m: Map<string, CleaningChatRow>) => void;
}) {
  const [search, setSearch] = useState("");
  const [channels, setChannels] = useState<CleaningChannelRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [summary, setSummary] = useState<CleaningTeamsSummary | null>(null);
  const [requestingCounts, setRequestingCounts] = useState(false);
  const debouncedSearch = useDebouncedValue(search, 300);

  const refresh = useCallback(async () => {
    try {
      const [{ channels: ch }, s] = await Promise.all([listTeamsChannels(connectionId, { search: debouncedSearch }), getTeamsSummary(connectionId)]);
      setChannels(ch);
      setSummary(s);
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't load Teams data. Try again.");
    } finally {
      setLoading(false);
    }
  }, [connectionId, debouncedSearch]);

  useEffect(() => {
    refresh();
  }, [refresh]);

  const dmsFetcher = useCallback(
    (opts: { search?: string; page?: number; pageSize?: number }) =>
      listTeamsDMs(connectionId, opts).then((r) => ({ rows: r.chats, total: r.total, page: r.page, pageSize: r.pageSize })),
    [connectionId]
  );
  const dms = usePagedList(dmsFetcher, search);

  // While discovery or message-count calculation is running, keep polling so channels/counts fill
  // in live instead of requiring a manual refresh.
  useEffect(() => {
    const live = summary?.structureScan?.status === "running" || summary?.structureScan?.status === "queued" ||
      summary?.countScan?.status === "running" || summary?.countScan?.status === "queued";
    if (!live) return;
    const timer = setTimeout(refresh, 3000);
    return () => clearTimeout(timer);
  }, [summary, refresh]);

  async function handleCalculateCounts() {
    setRequestingCounts(true);
    try {
      await calculateTeamsMessageCounts(connectionId);
      await refresh();
    } catch {
      // A 409 here just means a count job is already running — refresh will pick up its progress.
    } finally {
      setRequestingCounts(false);
    }
  }

  const columns: DiscoveryColumn<CleaningChatRow>[] = [
    {
      label: "Participants",
      render: (r) => r.participants.map((p) => p.displayName ?? p.upn ?? "Unknown").join(" ↔ ") || "—",
    },
    {
      label: "Messages",
      align: "right",
      render: (r) =>
        r.countStatus === "completed" && r.messageCount !== null
          ? r.messageCount.toLocaleString()
          : r.countStatus === "calculating"
            ? <span className="italic text-slate-400">Calculating…</span>
            : r.countStatus === "failed"
              ? <span className="text-rose-500">Unable to calculate</span>
              : <span className="italic text-slate-400">Waiting to be calculated</span>,
    },
    { label: "Last Activity", render: (r) => formatDate(r.lastMessageAt) },
  ];

  // Failed items are retryable via this same button (the backend re-picks up 'pending' AND
  // 'failed' rows), so they count toward what clicking it will actually process.
  const awaitingCount = summary ? summary.itemsAwaitingCount + summary.itemsFailedCount : 0;
  const isCalculating = summary?.countScan?.status === "running" || summary?.countScan?.status === "queued";

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Microsoft Teams</h2>
      <p className="mb-5 text-sm text-slate-500">Teams, channels, and direct messages</p>

      <div className="mb-4 flex items-center justify-between">
        <div className="flex gap-1 rounded-lg bg-slate-100 p-1">
          <button
            onClick={() => onTabChange("channels")}
            className={`rounded-md px-4 py-1.5 text-sm font-medium ${tab === "channels" ? "bg-white text-slate-800 shadow-sm" : "text-slate-500"}`}
          >
            Channels
          </button>
          <button
            onClick={() => onTabChange("dms")}
            className={`rounded-md px-4 py-1.5 text-sm font-medium ${tab === "dms" ? "bg-white text-slate-800 shadow-sm" : "text-slate-500"}`}
          >
            Direct Messages
          </button>
        </div>

        {awaitingCount > 0 && (
          <button
            onClick={handleCalculateCounts}
            disabled={requestingCounts || isCalculating}
            className="rounded-md bg-[#1b2fc4] px-4 py-1.5 text-sm font-semibold text-white hover:opacity-90 disabled:opacity-50"
          >
            {isCalculating ? "Calculating message counts…" : `Calculate message counts (${awaitingCount.toLocaleString()})`}
          </button>
        )}
      </div>

      {tab === "channels" ? (
        <TeamsChannels
          channels={channels}
          loading={loading}
          error={error}
          search={search}
          onSearchChange={setSearch}
          selected={new Set(selectedChannels.keys())}
          onToggle={(id) => {
            const row = channels.find((c) => c.id === id);
            if (row) toggleInMap(selectedChannels, setSelectedChannels, id, row);
          }}
          onToggleTeam={(ids) => {
            const allSelected = ids.every((id) => selectedChannels.has(id));
            const next = new Map(selectedChannels);
            for (const id of ids) {
              const row = channels.find((c) => c.id === id);
              if (!row) continue;
              allSelected ? next.delete(id) : next.set(id, row);
            }
            setSelectedChannels(next);
          }}
        />
      ) : (
        <DiscoveryTable
          title="Direct Messages"
          columns={columns}
          rows={dms.rows}
          loading={dms.loading}
          error={dms.error}
          page={dms.page}
          totalPages={dms.totalPages}
          total={dms.total}
          onGoToPage={dms.goToPage}
          search={search}
          onSearchChange={setSearch}
          searchPlaceholder="Search conversations…"
          selected={new Set(selectedChats.keys())}
          onToggle={(id) => {
            const row = dms.rows.find((c) => c.id === id);
            if (row) toggleInMap(selectedChats, setSelectedChats, id, row);
          }}
          onToggleAll={() => {
            const allSelected = dms.rows.every((r) => selectedChats.has(r.id));
            const next = new Map(selectedChats);
            for (const r of dms.rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
            setSelectedChats(next);
          }}
          emptyMessage="No direct message conversations found."
        />
      )}
    </div>
  );
}

function ReviewPage({ totals, onBack }: { totals: SelectionTotals; onBack: () => void }) {
  const totalBytes = totals.oneDriveBytes + totals.sharePointBytes;
  const totalMessages = totals.teamsMessages + totals.dmMessages;

  return (
    <div className="mx-auto max-w-2xl px-8 py-10">
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Review your selection</h2>
      <p className="mb-6 text-sm text-slate-500">Nothing has been deleted. This is just a summary of what you've selected.</p>

      <div className="space-y-4 rounded-xl border border-slate-200 bg-white p-6">
        {totals.oneDriveAccounts > 0 && (
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <span className="text-sm font-medium text-slate-700">OneDrive</span>
            <span className="text-sm text-slate-600">{totals.oneDriveAccounts.toLocaleString()} accounts · {formatBytes(totals.oneDriveBytes)}</span>
          </div>
        )}
        {totals.sharePointSites > 0 && (
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <span className="text-sm font-medium text-slate-700">SharePoint</span>
            <span className="text-sm text-slate-600">{totals.sharePointSites.toLocaleString()} sites · {formatBytes(totals.sharePointBytes)}</span>
          </div>
        )}
        {totals.teamsChannels > 0 && (
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <span className="text-sm font-medium text-slate-700">Teams Channels</span>
            <span className="text-sm text-slate-600">
              {totals.teamsChannels.toLocaleString()} channels · {messagesFragment(totals.teamsChannels, totals.teamsChannelsWithKnownCount, totals.teamsMessages)}
            </span>
          </div>
        )}
        {totals.dms > 0 && (
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <span className="text-sm font-medium text-slate-700">Direct Messages</span>
            <span className="text-sm text-slate-600">
              {totals.dms.toLocaleString()} conversations · {messagesFragment(totals.dms, totals.dmsWithKnownCount, totals.dmMessages)}
            </span>
          </div>
        )}

        <div className="flex items-center justify-between pt-1">
          <span className="text-sm font-semibold text-slate-800">Total selected</span>
          <span className="text-sm font-semibold text-slate-800">
            {formatBytes(totalBytes)}{totalMessages > 0 ? ` + ${totalMessages.toLocaleString()} messages` : ""}
          </span>
        </div>
      </div>

      <p className="mt-6 rounded-lg border border-dashed border-slate-300 px-4 py-4 text-center text-sm text-slate-500">
        Cleanup actions will be available here after selection review.
      </p>

      <button onClick={onBack} className="mt-6 rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
        ← Back
      </button>
    </div>
  );
}
