import { useEffect, useState } from "react";
import {
  getDataDumpSummary,
  listDataDumpHistory,
  listDataDumpTenants,
  previewDataDump,
  startDataDump,
  type ChannelConfig,
  type DataDumpOperationRow,
  type DataDumpPreviewResult,
  type DataDumpSummary,
  type DataDumpTenant,
  type DataDumpWorkload,
  type NewSiteConfig,
  type NewTeamConfig,
  type OneDriveGenConfig,
  type OutlookGenConfig,
  type SharePointGenConfig,
} from "../api/dataDump";
import type { AvailableResourceRow } from "../api/clouds";
import { ApiClientError } from "../api/client";
import { DataDumpOperationView } from "../components/dataDump/DataDumpOperationView";
import { DataDumpHistoryTable } from "../components/dataDump/DataDumpHistoryTable";
import { ResourceSelectionStep, connectionResourceFetcher } from "../components/dataDump/ResourceSelectionStep";
import { SharePointSiteStep } from "../components/dataDump/SharePointSiteStep";
import { TeamsStep } from "../components/dataDump/TeamsStep";
import { OneDriveConfigureStep, OutlookConfigureStep, SharePointConfigureStep } from "../components/dataDump/WorkloadConfigureStep";
import { DataDumpSelectionBar, type DataDumpSelectionTotals } from "../components/dataDump/DataDumpSelectionBar";
import { useDebouncedValue } from "../components/cleaning/DiscoveryTable";
import { formatBytes, formatDate } from "../utils/format";

/**
 * Data Dump — a separate top-level module from Cleaning (spec §2/§3/§36), but deliberately mirroring
 * Cleanup's own flow shape end to end (spec §38, the "most important requirement"):
 *   Select Resources → Configure Generation → Review → Create → Progress → Results → Report
 * Reuses Cleanup's generic UI building blocks unchanged (DiscoveryTable, its pagination/search/
 * select-all) via components/dataDump/ResourceSelectionStep.tsx, but owns 100% of its own selection
 * state, config, and API calls — never touches CleaningPage's Maps/CleanupManifest.
 */

type Tab = "new" | "history";
type WorkloadView = "landing" | "dashboard" | DataDumpWorkload | "review" | "progress";

function readTab(): Tab {
  return new URLSearchParams(window.location.search).get("tab") === "history" ? "history" : "new";
}

const WORKLOAD_LABELS: Record<DataDumpWorkload, string> = {
  onedrive: "OneDrive",
  sharepoint: "SharePoint",
  teams: "Microsoft Teams",
  outlook: "Outlook",
};

export function DataDumpPage() {
  const [tab, setTab] = useState<Tab>(readTab);
  const [summary, setSummary] = useState<DataDumpSummary | null>(null);

  useEffect(() => {
    const url = tab === "history" ? "/data-dump?tab=history" : "/data-dump";
    if (`${window.location.pathname}${window.location.search}` !== url) window.history.replaceState({}, "", url);
  }, [tab]);

  useEffect(() => {
    getDataDumpSummary().then(setSummary).catch(() => {});
  }, [tab]);

  return (
    <div className="mx-auto max-w-5xl px-8 py-8 pb-28">
      <h1 className="mb-1 text-xl font-semibold text-slate-900">Data Dump</h1>
      <p className="mb-6 text-sm text-slate-500">Generate realistic Microsoft 365 data for demos, migration testing, and performance testing.</p>

      {summary && (
        <div className="mb-6 grid grid-cols-3 gap-4 rounded-xl border border-slate-200 bg-white p-5">
          <SummaryStat label="Total Operations" value={summary.total} />
          <SummaryStat label="In Progress" value={summary.running} accent="text-[#1b2fc4]" />
          <SummaryStat label="Completed" value={summary.completed} accent="text-emerald-600" />
        </div>
      )}

      <div className="mb-6 flex gap-1 border-b border-slate-200">
        <TabButton active={tab === "new"} onClick={() => setTab("new")}>
          New Generation
        </TabButton>
        <TabButton active={tab === "history"} onClick={() => setTab("history")}>
          History
        </TabButton>
      </div>

      {tab === "new" ? <NewGenerationFlow /> : <HistoryTab />}
    </div>
  );
}

function SummaryStat({ label, value, accent }: { label: string; value: number; accent?: string }) {
  return (
    <div>
      <div className="text-xs uppercase tracking-wide text-slate-400">{label}</div>
      <div className={`mt-1 text-lg font-semibold ${accent ?? "text-slate-900"}`}>{value.toLocaleString()}</div>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: React.ReactNode }) {
  return (
    <button
      onClick={onClick}
      className={`-mb-px border-b-2 px-4 py-2 text-sm font-medium ${active ? "border-[#1b2fc4] text-[#1b2fc4]" : "border-transparent text-slate-500 hover:text-slate-700"}`}
    >
      {children}
    </button>
  );
}

// --- Per-workload selection state --------------------------------------------------------------

interface OneDriveSelection {
  users: Map<string, AvailableResourceRow>;
  config: Partial<OneDriveGenConfig>;
}
interface SharePointSelection {
  sites: Map<string, AvailableResourceRow>;
  newSite: NewSiteConfig | null;
  config: Partial<SharePointGenConfig>;
}
interface TeamsSelection {
  teams: Map<string, AvailableResourceRow>;
  newTeam: NewTeamConfig | null;
  channels: ChannelConfig[];
}
interface OutlookSelection {
  users: Map<string, AvailableResourceRow>;
  config: Partial<OutlookGenConfig>;
}

function emptyOneDrive(): OneDriveSelection {
  return { users: new Map(), config: {} };
}
function emptySharePoint(): SharePointSelection {
  return { sites: new Map(), newSite: null, config: {} };
}
function emptyTeams(): TeamsSelection {
  return { teams: new Map(), newTeam: null, channels: [] };
}
function emptyOutlook(): OutlookSelection {
  return { users: new Map(), config: {} };
}

function toggleInMap<T>(map: Map<string, T>, id: string, row: T): Map<string, T> {
  const next = new Map(map);
  next.has(id) ? next.delete(id) : next.set(id, row);
  return next;
}

function NewGenerationFlow() {
  const [view, setView] = useState<WorkloadView>("landing");
  const [tenants, setTenants] = useState<DataDumpTenant[]>([]);
  const [tenantsLoading, setTenantsLoading] = useState(true);
  const [tenantId, setTenantId] = useState("");
  const [namingPrefix, setNamingPrefix] = useState("CF-Demo");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [operationId, setOperationId] = useState<string | null>(null);
  const [preview, setPreview] = useState<DataDumpPreviewResult | null>(null);

  const [onedrive, setOnedrive] = useState<OneDriveSelection>(emptyOneDrive);
  const [sharepoint, setSharepoint] = useState<SharePointSelection>(emptySharePoint);
  const [teams, setTeams] = useState<TeamsSelection>(emptyTeams);
  const [outlook, setOutlook] = useState<OutlookSelection>(emptyOutlook);

  useEffect(() => {
    listDataDumpTenants()
      .then(({ tenants: rows }) => setTenants(rows))
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load connected tenants."))
      .finally(() => setTenantsLoading(false));
  }, []);

  function openTenant(t: DataDumpTenant) {
    setTenantId(t.tenantId);
    resetAllSelections();
    setView("dashboard");
  }

  const tenant = tenants.find((t) => t.tenantId === tenantId);

  function resetAllSelections() {
    setOnedrive(emptyOneDrive());
    setSharepoint(emptySharePoint());
    setTeams(emptyTeams());
    setOutlook(emptyOutlook());
  }

  const totals: DataDumpSelectionTotals = {
    onedriveUsers: onedrive.users.size,
    sharepointSites: sharepoint.sites.size,
    sharepointNewSite: !!sharepoint.newSite?.displayName,
    teamsTeams: teams.teams.size,
    teamsNewTeam: !!teams.newTeam?.displayName,
    outlookUsers: outlook.users.size,
  };

  function buildWorkloads(): DataDumpWorkload[] {
    const list: DataDumpWorkload[] = [];
    if (onedrive.users.size > 0) list.push("onedrive");
    if (sharepoint.sites.size > 0 || sharepoint.newSite?.displayName) list.push("sharepoint");
    if (teams.teams.size > 0 || teams.newTeam?.displayName) list.push("teams");
    if (outlook.users.size > 0) list.push("outlook");
    return list;
  }

  function buildRequestBody(profile: "custom" = "custom") {
    return {
      tenantId,
      profile,
      workloads: buildWorkloads(),
      namingPrefix,
      onedrive: onedrive.users.size > 0 ? { ...onedrive.config, selectedUserIds: [...onedrive.users.keys()] } : undefined,
      sharepoint:
        sharepoint.sites.size > 0 || sharepoint.newSite?.displayName
          ? { ...sharepoint.config, selectedSiteIds: [...sharepoint.sites.keys()], newSite: sharepoint.newSite?.displayName ? sharepoint.newSite : undefined }
          : undefined,
      teams:
        teams.teams.size > 0 || teams.newTeam?.displayName
          ? { selectedTeamIds: [...teams.teams.keys()], newTeam: teams.newTeam?.displayName ? teams.newTeam : undefined, channels: teams.channels }
          : undefined,
      outlook: outlook.users.size > 0 ? { ...outlook.config, selectedUserIds: [...outlook.users.keys()] } : undefined,
    };
  }

  async function handleReview() {
    setError(null);
    setLoading(true);
    try {
      const body = buildRequestBody();
      const result = await previewDataDump(body);
      setPreview(result.preview);
      setView("review");
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't compute a preview.");
    } finally {
      setLoading(false);
    }
  }

  async function handleGenerate() {
    setError(null);
    setLoading(true);
    try {
      const { operationId: newId } = await startDataDump(buildRequestBody());
      setOperationId(newId);
      setView("progress");
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't start data generation.");
    } finally {
      setLoading(false);
    }
  }

  function backToDashboard() {
    setView("dashboard");
  }

  if (view === "landing") {
    return (
      <div>
        <TenantLanding tenants={tenants} loading={tenantsLoading} onOpen={openTenant} />
        {error && <p className="mt-4 text-sm text-rose-600">{error}</p>}
      </div>
    );
  }

  if (view === "progress" && operationId) {
    return (
      <DataDumpOperationView
        operationId={operationId}
        onClose={() => {
          setView("dashboard");
          setOperationId(null);
          resetAllSelections();
        }}
      />
    );
  }

  if (view === "review" && preview) {
    return (
      <ReviewStep
        preview={preview}
        loading={loading}
        error={error}
        onBack={() => setView("dashboard")}
        onGenerate={handleGenerate}
      />
    );
  }

  let content: React.ReactNode;

  if (view === "onedrive" && tenant?.connectionsByWorkload.onedrive) {
    content = (
      <div className="space-y-6">
        <BackLink onClick={backToDashboard} />
        <ResourceSelectionStep
          fetcher={connectionResourceFetcher(tenant.connectionsByWorkload.onedrive)}
          title="OneDrive"
          subtitle="Select users to generate folders/files for."
          searchPlaceholder="Search users…"
          secondaryLabel="Storage Used"
          selected={onedrive.users}
          onToggle={(id, row) => setOnedrive((prev) => ({ ...prev, users: toggleInMap(prev.users, id, row) }))}
          onToggleAll={(rows) =>
            setOnedrive((prev) => {
              const allSelected = rows.every((r) => prev.users.has(r.id));
              const next = new Map(prev.users);
              for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
              return { ...prev, users: next };
            })
          }
          emptyMessage="No OneDrive users found."
        />
        {onedrive.users.size > 0 && (
          <div className="rounded-xl border border-slate-200 bg-white p-5">
            <h3 className="mb-3 text-sm font-semibold text-slate-800">Configure Generation ({onedrive.users.size} user{onedrive.users.size === 1 ? "" : "s"})</h3>
            <OneDriveConfigureStep onChange={(updater) => setOnedrive((prev) => ({ ...prev, config: updater(prev.config as OneDriveGenConfig) }))} />
          </div>
        )}
      </div>
    );
  } else if (view === "sharepoint") {
    const connectionId = tenant?.connectionsByWorkload.sharepoint;
    content = (
      <div className="space-y-6">
        <BackLink onClick={backToDashboard} />
        {connectionId ? (
          <SharePointSiteStep
            connectionId={connectionId}
            selectedSites={sharepoint.sites}
            onToggleSite={(id, row) => setSharepoint((prev) => ({ ...prev, sites: toggleInMap(prev.sites, id, row) }))}
            onToggleAllSites={(rows) =>
              setSharepoint((prev) => {
                const allSelected = rows.every((r) => prev.sites.has(r.id));
                const next = new Map(prev.sites);
                for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
                return { ...prev, sites: next };
              })
            }
            newSite={sharepoint.newSite}
            onNewSiteChange={(site) => setSharepoint((prev) => ({ ...prev, newSite: site }))}
          />
        ) : (
          <p className="rounded-xl border border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-500">Connect SharePoint via Add Clouds to select existing sites, or create a new one.</p>
        )}
        {(sharepoint.sites.size > 0 || sharepoint.newSite?.displayName) && (
          <div className="rounded-xl border border-slate-200 bg-white p-5">
            <h3 className="mb-3 text-sm font-semibold text-slate-800">Configure Generation</h3>
            <SharePointConfigureStep onChange={(updater) => setSharepoint((prev) => ({ ...prev, config: updater(prev.config as SharePointGenConfig) }))} />
          </div>
        )}
      </div>
    );
  } else if (view === "teams") {
    content = (
      <div className="space-y-6">
        <BackLink onClick={backToDashboard} />
        <TeamsStep
          tenantId={tenantId}
          teamsConnectionId={tenant?.connectionsByWorkload.teams ?? null}
          selectedTeams={teams.teams}
          onToggleTeam={(id, row) => setTeams((prev) => ({ ...prev, teams: toggleInMap(prev.teams, id, row) }))}
          onToggleAllTeams={(rows) =>
            setTeams((prev) => {
              const allSelected = rows.every((r) => prev.teams.has(r.id));
              const next = new Map(prev.teams);
              for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
              return { ...prev, teams: next };
            })
          }
          newTeam={teams.newTeam}
          onNewTeamChange={(t) => setTeams((prev) => ({ ...prev, newTeam: t }))}
          channels={teams.channels}
          onChannelsChange={(channels) => setTeams((prev) => ({ ...prev, channels }))}
        />
      </div>
    );
  } else if (view === "outlook" && tenant?.connectionsByWorkload.outlook) {
    content = (
      <div className="space-y-6">
        <BackLink onClick={backToDashboard} />
        <ResourceSelectionStep
          fetcher={connectionResourceFetcher(tenant.connectionsByWorkload.outlook)}
          title="Outlook"
          subtitle="Select users to generate emails, calendar events, and contacts for."
          searchPlaceholder="Search users…"
          secondaryLabel="Mailbox"
          selected={outlook.users}
          onToggle={(id, row) => setOutlook((prev) => ({ ...prev, users: toggleInMap(prev.users, id, row) }))}
          onToggleAll={(rows) =>
            setOutlook((prev) => {
              const allSelected = rows.every((r) => prev.users.has(r.id));
              const next = new Map(prev.users);
              for (const r of rows) allSelected ? next.delete(r.id) : next.set(r.id, r);
              return { ...prev, users: next };
            })
          }
          emptyMessage="No Outlook users found."
        />
        {outlook.users.size > 0 && (
          <div className="rounded-xl border border-slate-200 bg-white p-5">
            <h3 className="mb-3 text-sm font-semibold text-slate-800">Configure Generation ({outlook.users.size} user{outlook.users.size === 1 ? "" : "s"})</h3>
            <OutlookConfigureStep onChange={(updater) => setOutlook((prev) => ({ ...prev, config: updater(prev.config as OutlookGenConfig) }))} />
          </div>
        )}
      </div>
    );
  } else {
    // Dashboard: naming prefix + workload tiles, scoped to the tenant opened from the landing grid.
    content = (
      <div>
        <BackLink label="← Back to Tenants" onClick={() => setView("landing")} />
        <div className="mb-6 mt-4 rounded-xl border border-slate-200 bg-white p-6">
          <div className="mb-4 text-sm text-slate-500">
            Generating for <span className="font-semibold text-slate-800">{tenant?.displayName ?? "…"}</span>
          </div>

          <label className="mb-2 block text-sm font-medium text-slate-700">Naming Prefix</label>
          <input
            value={namingPrefix}
            onChange={(e) => setNamingPrefix(e.target.value)}
            className="w-full max-w-xs rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-[#1b2fc4] focus:outline-none"
            placeholder="CF-Demo"
          />
          <p className="mt-1 text-xs text-slate-400">Makes generated data easy to identify, e.g. CF-Demo-Finance, CF-Demo-QuarterlyReport.xlsx.</p>
        </div>

      <h2 className="mb-3 text-sm font-semibold text-slate-800">Select a workload to choose resources</h2>
      <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
        <WorkloadTile label="OneDrive" selectedCount={onedrive.users.size} onClick={() => setView("onedrive")} />
        <WorkloadTile label="SharePoint" selectedCount={sharepoint.sites.size + (sharepoint.newSite?.displayName ? 1 : 0)} onClick={() => setView("sharepoint")} />
        <WorkloadTile label="Microsoft Teams" selectedCount={teams.teams.size + (teams.newTeam?.displayName ? 1 : 0)} onClick={() => setView("teams")} />
        <WorkloadTile label="Outlook" selectedCount={outlook.users.size} onClick={() => setView("outlook")} />
      </div>

        {error && <p className="mt-4 text-sm text-rose-600">{error}</p>}
      </div>
    );
  }

  return (
    <>
      {content}
      <DataDumpSelectionBar totals={totals} onReview={handleReview} reviewLabel={loading ? "Loading…" : "Review Data Dump"} />
    </>
  );
}

function BackLink({ onClick, label = "← Back to Data Dump" }: { onClick: () => void; label?: string }) {
  return (
    <button onClick={onClick} className="text-sm font-medium text-[#1b2fc4] hover:underline">
      {label}
    </button>
  );
}

const DATA_DUMP_STATUS_BADGE = "bg-emerald-50 text-emerald-700";

function TenantLanding({ tenants, loading, onOpen }: { tenants: DataDumpTenant[]; loading: boolean; onOpen: (t: DataDumpTenant) => void }) {
  if (loading) return <p className="text-sm text-slate-500">Loading…</p>;
  if (tenants.length === 0) {
    return (
      <p className="rounded-lg border border-dashed border-slate-300 px-4 py-8 text-center text-sm text-slate-500">
        No connected Microsoft 365 tenants yet — connect one from the Clouds tab first.
      </p>
    );
  }
  return (
    <div>
      <h2 className="mb-5 text-base font-semibold text-slate-800">Data Dump</h2>
      <div className="flex flex-wrap gap-4">
        {tenants.map((t) => (
          <div key={t.tenantId} className="w-72 rounded-xl border border-slate-200 bg-white p-5">
            <div className="mb-3 flex items-center justify-between">
              <span className="text-sm font-semibold text-slate-800">Microsoft 365</span>
              <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${DATA_DUMP_STATUS_BADGE}`}>Connected</span>
            </div>
            <div className="space-y-1 text-sm text-slate-600">
              <div>👤 {t.adminDisplayName ?? t.adminUpn}</div>
              <div>🌐 {t.displayName}</div>
            </div>
            {t.lastUpdatedAt && <div className="mt-2 text-xs text-slate-400">Last updated {formatDate(t.lastUpdatedAt)}</div>}
            <button
              onClick={() => onOpen(t)}
              className="mt-4 w-full rounded-md bg-[#1b2fc4] py-2 text-sm font-semibold text-white hover:opacity-90"
            >
              Open Data Dump
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function WorkloadTile({ label, selectedCount, onClick }: { label: string; selectedCount: number; onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className={`rounded-xl border p-5 text-left transition-colors ${selectedCount > 0 ? "border-[#1b2fc4] bg-[#1b2fc4]/5" : "border-slate-200 bg-white hover:border-slate-300"}`}
    >
      <div className="text-sm font-semibold text-slate-800">{label}</div>
      <div className="mt-1 text-xs text-slate-500">{selectedCount > 0 ? `${selectedCount} selected` : "Select resources"}</div>
    </button>
  );
}

function ReviewStep({
  preview,
  loading,
  error,
  onBack,
  onGenerate,
}: {
  preview: DataDumpPreviewResult;
  loading: boolean;
  error: string | null;
  onBack: () => void;
  onGenerate: () => void;
}) {
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-6">
      <h2 className="mb-1 text-lg font-semibold text-slate-900">Review Data Dump</h2>
      <p className="mb-4 text-sm text-slate-500">Nothing has been created yet. This is an estimate based on your configuration.</p>
      <div className="mb-6 space-y-4">
        {preview.workloads.map((w) => (
          <div key={w.workload} className="rounded-lg border border-slate-100 bg-slate-50 p-4">
            <div className="mb-2 flex items-center justify-between">
              <span className="font-medium text-slate-800">{WORKLOAD_LABELS[w.workload]}</span>
              {w.estimatedSizeBytes > 0 && <span className="text-sm text-slate-500">{formatBytes(w.estimatedSizeBytes)}</span>}
            </div>
            <div className="grid grid-cols-2 gap-2 text-sm text-slate-600 sm:grid-cols-4">
              {Object.entries(w.breakdown).map(([key, value]) => (
                <div key={key}>
                  <span className="capitalize text-slate-400">{key.replace(/([A-Z])/g, " $1")}: </span>
                  <span className="font-medium">{value.toLocaleString()}</span>
                </div>
              ))}
            </div>
            {w.warnings.map((warning, i) => (
              <p key={i} className="mt-2 text-xs text-amber-700">
                ⚠ {warning}
              </p>
            ))}
          </div>
        ))}
      </div>
      <div className="flex items-center justify-between border-t border-slate-100 pt-4">
        <div>
          <div className="text-xs uppercase tracking-wide text-slate-400">Total</div>
          <div className="text-lg font-semibold text-slate-900">
            {preview.totalRequestedObjects.toLocaleString()} objects{preview.totalEstimatedSizeBytes > 0 ? ` · ${formatBytes(preview.totalEstimatedSizeBytes)}` : ""}
          </div>
        </div>
        <div className="flex gap-3">
          <button onClick={onBack} className="rounded-lg border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">
            ← Back
          </button>
          <button onClick={onGenerate} disabled={loading} className="rounded-lg bg-[#1b2fc4] px-5 py-2 text-sm font-medium text-white hover:bg-[#1725a0] disabled:opacity-50">
            {loading ? "Starting…" : "Start Data Dump"}
          </button>
        </div>
      </div>
      {error && <p className="mt-3 text-sm text-rose-600">{error}</p>}
    </div>
  );
}

function HistoryTab() {
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
    listDataDumpHistory({ page, pageSize: 20, search: search || undefined })
      .then((res) => {
        setOperations(res.operations);
        setTotal(res.total);
        setError(null);
      })
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load Data Dump history."))
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
        <DataDumpOperationView operationId={detailId} onClose={() => { setDetailId(null); load(); }} />
      </div>
    );
  }

  const totalPages = Math.max(1, Math.ceil(total / 20));

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
