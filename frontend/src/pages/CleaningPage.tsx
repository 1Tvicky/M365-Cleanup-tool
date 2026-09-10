import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  calculateTeamsMessageCounts,
  getLatestSyncOperation,
  getSyncOperation,
  getTeamsSummary,
  listCleaningConnections,
  listGoogleMyDriveAccounts,
  listOneDriveAccounts,
  listSharedDrives,
  listOutlookCalendars,
  listOutlookContacts,
  listOutlookMailboxes,
  listOutlookOverview,
  listSharePointSites,
  listTeamsChannels,
  listTeamsDMs,
  startSync,
  validateCleanup,
  type CleaningChannelRow,
  type CleaningChatRow,
  type CleaningConnectionRow,
  type CleaningResourceRow,
  type CleaningSyncOperation,
  type CleaningSyncResourceStatus,
  type CleaningTeamsSummary,
  type CleanupManifest,
  type OutlookMailboxOverviewRow,
  type OutlookOverviewSubResource,
  type PageResult,
} from "../api/cleaning";
import { ApiClientError } from "../api/client";
import { DiscoveryTable, useDebouncedValue, type DiscoveryColumn } from "../components/cleaning/DiscoveryTable";
import { TeamsChannels } from "../components/cleaning/TeamsChannels";
import { SelectionSummary, hasSelection, messagesFragment, type SelectionTotals } from "../components/cleaning/SelectionSummary";
import { CleanupConfirmation } from "../components/cleaning/CleanupConfirmation";
import { CleanupProgressView } from "../components/cleaning/CleanupProgress";
import { CleanupResultsView } from "../components/cleaning/CleanupResults";
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
  outlook?: CleaningConnectionRow;
  google_my_drive?: CleaningConnectionRow;
  shared_drive?: CleaningConnectionRow;
  google_chat?: CleaningConnectionRow;
  gmail?: CleaningConnectionRow;
}

type View =
  | "landing"
  | "dashboard"
  | "onedrive"
  | "sharepoint"
  | "teams"
  | "outlook"
  | "google_my_drive"
  | "shared_drive"
  | "review"
  | "cleanupConfirm"
  | "cleanupProgress"
  | "cleanupResults";

/**
 * Only these "browsing" views are reflected in the URL (as `?group=<domain>&view=<view>` on the same
 * /cleaning path — no new top-level route, so App.tsx's own page routing is untouched). The
 * selection/cleanup flow (review/cleanupConfirm/cleanupProgress/cleanupResults) depends on in-memory
 * selection Maps and a manifest that aren't meaningfully reconstructable from a URL, so it
 * deliberately doesn't touch the address bar at all — reloading mid-flow already falls back to
 * Landing today, and that isn't something this change is meant to fix.
 */
const URL_SYNCED_VIEWS = new Set<View>(["landing", "dashboard", "onedrive", "sharepoint", "teams", "outlook", "google_my_drive", "shared_drive"]);

function cleaningUrlFor(view: View, activeGroup: TenantGroup | null): string {
  if (view === "landing" || !activeGroup) return "/cleaning";
  const params = new URLSearchParams({ group: activeGroup.domain, view });
  return `/cleaning?${params.toString()}`;
}

/** Reads `?group=&view=` back out, resolving `group` against the just-loaded list — used both on initial mount (deep link / reload) and browser Back/Forward. Falls back to Landing whenever the URL doesn't name a real, currently-visible group. */
function cleaningStateFromUrl(groups: TenantGroup[]): { view: View; group: TenantGroup | null } {
  const params = new URLSearchParams(window.location.search);
  const domain = params.get("group");
  const urlView = params.get("view") as View | null;
  const group = domain ? groups.find((g) => g.domain === domain) ?? null : null;
  if (!group || !urlView || !URL_SYNCED_VIEWS.has(urlView) || urlView === "landing") {
    return { view: "landing", group: null };
  }
  return { view: urlView, group };
}

function groupConnectionsByDomain(connections: CleaningConnectionRow[]): TenantGroup[] {
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
  return [...byDomain.values()];
}

/** Keeps only the entries whose key is in `foundIds` — used to drop selected rows that a sync found no longer exist in Microsoft 365, identified by their stable internal id, never by display name. */
function pruneToFound<T>(map: Map<string, T>, foundIds: string[]): Map<string, T> {
  const foundSet = new Set(foundIds);
  const next = new Map<string, T>();
  for (const [id, row] of map) if (foundSet.has(id)) next.set(id, row);
  return next;
}

/** Only 'oneDrive'/'sharePoint'/'outlook'/'outlookCalendar'/'outlookContacts' slots ever lead to a real Graph delete — 'channels'/'chats' always resolve to 'unsupported' server-side (see CleanupConfirmation), but are still included so the confirmation screen can show them transparently rather than silently dropping them. */
function buildCleanupManifest(
  group: TenantGroup,
  selectedOneDrive: Map<string, CleaningResourceRow>,
  selectedSharePoint: Map<string, CleaningResourceRow>,
  selectedOutlook: Map<string, CleaningResourceRow>,
  selectedOutlookCalendar: Map<string, CleaningResourceRow>,
  selectedOutlookContacts: Map<string, CleaningResourceRow>,
  selectedChannels: Map<string, CleaningChannelRow>,
  selectedChats: Map<string, CleaningChatRow>,
  selectedGoogleMyDrive: Map<string, CleaningResourceRow>,
  selectedSharedDrives: Map<string, CleaningResourceRow>
): CleanupManifest {
  const manifest: CleanupManifest = {};
  if (selectedOneDrive.size > 0 && group.onedrive) manifest.oneDrive = { connectionId: group.onedrive.id, ids: [...selectedOneDrive.keys()] };
  if (selectedSharePoint.size > 0 && group.sharepoint) manifest.sharePoint = { connectionId: group.sharepoint.id, ids: [...selectedSharePoint.keys()] };
  if (selectedOutlook.size > 0 && group.outlook) manifest.outlook = { connectionId: group.outlook.id, ids: [...selectedOutlook.keys()] };
  if (selectedOutlookCalendar.size > 0 && group.outlook) manifest.outlookCalendar = { connectionId: group.outlook.id, ids: [...selectedOutlookCalendar.keys()] };
  if (selectedOutlookContacts.size > 0 && group.outlook) manifest.outlookContacts = { connectionId: group.outlook.id, ids: [...selectedOutlookContacts.keys()] };
  if (selectedChannels.size > 0 && group.teams) manifest.channels = { connectionId: group.teams.id, ids: [...selectedChannels.keys()] };
  if (selectedChats.size > 0 && group.teams) manifest.chats = { connectionId: group.teams.id, ids: [...selectedChats.keys()] };
  if (selectedGoogleMyDrive.size > 0 && group.google_my_drive) manifest.googleMyDrive = { connectionId: group.google_my_drive.id, ids: [...selectedGoogleMyDrive.keys()] };
  if (selectedSharedDrives.size > 0 && group.shared_drive) manifest.sharedDrives = { connectionId: group.shared_drive.id, ids: [...selectedSharedDrives.keys()] };
  return manifest;
}

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

export function CleaningPage({ onCleanupStarted }: { onCleanupStarted?: (operationId: string) => void } = {}) {
  const [view, setView] = useState<View>("landing");
  const [groups, setGroups] = useState<TenantGroup[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);
  const [activeGroup, setActiveGroup] = useState<TenantGroup | null>(null);
  const [teamsTab, setTeamsTab] = useState<"channels" | "dms">("channels");

  const [selectedOneDrive, setSelectedOneDrive] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedSharePoint, setSelectedSharePoint] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedOutlook, setSelectedOutlook] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedOutlookCalendar, setSelectedOutlookCalendar] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedOutlookContacts, setSelectedOutlookContacts] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedChannels, setSelectedChannels] = useState<Map<string, CleaningChannelRow>>(new Map());
  const [selectedChats, setSelectedChats] = useState<Map<string, CleaningChatRow>>(new Map());
  const [selectedGoogleMyDrive, setSelectedGoogleMyDrive] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [selectedSharedDrives, setSelectedSharedDrives] = useState<Map<string, CleaningResourceRow>>(new Map());
  const [cleanupOperationId, setCleanupOperationId] = useState<string | null>(null);
  const [reconciliationBanner, setReconciliationBanner] = useState<string | null>(null);

  // Guards the URL-sync effect below from firing with the pre-restore default state (view="landing",
  // no group) before the initial deep-link/reload restore below has had a chance to run — otherwise
  // it would clobber a real "?group=&view=" URL back to plain "/cleaning" for one render, and the
  // resulting extra history entry would make Back one click short once the real state comes in.
  const urlRestoredRef = useRef(false);

  useEffect(() => {
    listCleaningConnections()
      .then(({ connections }) => {
        const nextGroups = groupConnectionsByDomain(connections);
        setGroups(nextGroups);
        const restored = cleaningStateFromUrl(nextGroups);
        if (restored.group) {
          setActiveGroup(restored.group);
          setView(restored.view);
        }
      })
      .finally(() => {
        setLoadingGroups(false);
        urlRestoredRef.current = true;
      });
  }, []);

  // Keeps the address bar naming the tenant + sub-view actually on screen — see URL_SYNCED_VIEWS.
  useEffect(() => {
    if (!urlRestoredRef.current || !URL_SYNCED_VIEWS.has(view)) return;
    const url = cleaningUrlFor(view, activeGroup);
    if (`${window.location.pathname}${window.location.search}` !== url) {
      window.history.pushState({}, "", url);
    }
  }, [view, activeGroup]);

  // Browser Back/Forward while inside Cleaning — App.tsx's own popstate listener only tracks which
  // top-level page (Clouds/Cleaning/Reports) is showing; the tenant + sub-view within Cleaning is
  // this component's own concern, read straight back out of the URL.
  useEffect(() => {
    function onPopState() {
      const restored = cleaningStateFromUrl(groups);
      setActiveGroup(restored.group);
      setView(restored.view);
    }
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, [groups]);

  /**
   * Runs after a sync completes: refreshes lastSyncedAt for the Dashboard/Landing display, then —
   * if there's a current selection — re-validates it and drops any selected id the sync no longer
   * found (identified by its stable id, never display name), surfacing a banner if anything changed.
   * Best-effort: if the reconciliation check itself fails, the selection is left as-is — the
   * existing validate-on-Continue-to-Cleanup check downstream would still catch anything stale.
   */
  async function handleSyncFinished() {
    listCleaningConnections().then(({ connections }) => {
      const nextGroups = groupConnectionsByDomain(connections);
      setGroups(nextGroups);
      setActiveGroup((current) => (current ? (nextGroups.find((g) => g.domain === current.domain) ?? current) : current));
    });

    if (!activeGroup || !hasSelection(totals)) return;
    const manifest = buildCleanupManifest(
      activeGroup,
      selectedOneDrive,
      selectedSharePoint,
      selectedOutlook,
      selectedOutlookCalendar,
      selectedOutlookContacts,
      selectedChannels,
      selectedChats,
      selectedGoogleMyDrive,
      selectedSharedDrives
    );
    try {
      const { foundIds } = await validateCleanup(manifest);
      const removedCount =
        selectedOneDrive.size -
        foundIds.oneDrive.length +
        (selectedSharePoint.size - foundIds.sharePoint.length) +
        (selectedOutlook.size - foundIds.outlook.length) +
        (selectedOutlookCalendar.size - foundIds.outlookCalendar.length) +
        (selectedOutlookContacts.size - foundIds.outlookContacts.length) +
        (selectedChannels.size - foundIds.channels.length) +
        (selectedChats.size - foundIds.chats.length) +
        (selectedGoogleMyDrive.size - foundIds.googleMyDrive.length) +
        (selectedSharedDrives.size - foundIds.sharedDrives.length);
      if (removedCount > 0) {
        setSelectedOneDrive((prev) => pruneToFound(prev, foundIds.oneDrive));
        setSelectedSharePoint((prev) => pruneToFound(prev, foundIds.sharePoint));
        setSelectedOutlook((prev) => pruneToFound(prev, foundIds.outlook));
        setSelectedOutlookCalendar((prev) => pruneToFound(prev, foundIds.outlookCalendar));
        setSelectedOutlookContacts((prev) => pruneToFound(prev, foundIds.outlookContacts));
        setSelectedChannels((prev) => pruneToFound(prev, foundIds.channels));
        setSelectedChats((prev) => pruneToFound(prev, foundIds.chats));
        setSelectedGoogleMyDrive((prev) => pruneToFound(prev, foundIds.googleMyDrive));
        setSelectedSharedDrives((prev) => pruneToFound(prev, foundIds.sharedDrives));
        setReconciliationBanner(
          `${removedCount.toLocaleString()} selected item${removedCount === 1 ? " is" : "s are"} no longer available — your selection has been updated.`
        );
      }
    } catch {
      // Best-effort reconciliation — see comment above.
    }
  }

  const totals: SelectionTotals = useMemo(
    () => ({
      oneDriveAccounts: selectedOneDrive.size,
      oneDriveBytes: [...selectedOneDrive.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
      sharePointSites: selectedSharePoint.size,
      sharePointBytes: [...selectedSharePoint.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
      outlookMailboxes: selectedOutlook.size,
      outlookItems: [...selectedOutlook.values()].reduce((s, r) => s + r.itemCount, 0),
      outlookCalendarUsers: selectedOutlookCalendar.size,
      outlookCalendarEvents: [...selectedOutlookCalendar.values()].reduce((s, r) => s + r.itemCount, 0),
      outlookContactUsers: selectedOutlookContacts.size,
      outlookContactCount: [...selectedOutlookContacts.values()].reduce((s, r) => s + r.itemCount, 0),
      teamsChannels: selectedChannels.size,
      teamsMessages: [...selectedChannels.values()].reduce((s, r) => s + (r.countStatus === "completed" ? r.messageCount ?? 0 : 0), 0),
      teamsChannelsWithKnownCount: [...selectedChannels.values()].filter((r) => r.countStatus === "completed").length,
      dms: selectedChats.size,
      dmMessages: [...selectedChats.values()].reduce((s, r) => s + (r.countStatus === "completed" ? r.messageCount ?? 0 : 0), 0),
      dmsWithKnownCount: [...selectedChats.values()].filter((r) => r.countStatus === "completed").length,
      googleMyDriveAccounts: selectedGoogleMyDrive.size,
      googleMyDriveBytes: [...selectedGoogleMyDrive.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
      sharedDrives: selectedSharedDrives.size,
      sharedDrivesBytes: [...selectedSharedDrives.values()].reduce((s, r) => s + r.storageUsedBytes, 0),
    }),
    [
      selectedOneDrive,
      selectedSharePoint,
      selectedOutlook,
      selectedOutlookCalendar,
      selectedOutlookContacts,
      selectedChannels,
      selectedChats,
      selectedGoogleMyDrive,
      selectedSharedDrives,
    ]
  );

  function openTenant(group: TenantGroup) {
    // Selections are scoped to one tenant at a time — switching tenants without clearing them
    // would let a review mix data from two different customers' tenants together.
    setSelectedOneDrive(new Map());
    setSelectedSharePoint(new Map());
    setSelectedOutlook(new Map());
    setSelectedOutlookCalendar(new Map());
    setSelectedOutlookContacts(new Map());
    setSelectedChannels(new Map());
    setSelectedChats(new Map());
    setSelectedGoogleMyDrive(new Map());
    setSelectedSharedDrives(new Map());
    setActiveGroup(group);
    setView("dashboard");
  }

  if (view === "review") {
    return (
      <ReviewPage
        totals={totals}
        selectedOutlookMail={selectedOutlook}
        setSelectedOutlookMail={setSelectedOutlook}
        selectedOutlookCalendar={selectedOutlookCalendar}
        setSelectedOutlookCalendar={setSelectedOutlookCalendar}
        selectedOutlookContacts={selectedOutlookContacts}
        setSelectedOutlookContacts={setSelectedOutlookContacts}
        onBack={() => setView("dashboard")}
        onContinue={() => setView("cleanupConfirm")}
      />
    );
  }

  if (view === "cleanupConfirm" && activeGroup) {
    return (
      <CleanupConfirmation
        manifest={buildCleanupManifest(
          activeGroup,
          selectedOneDrive,
          selectedSharePoint,
          selectedOutlook,
          selectedOutlookCalendar,
          selectedOutlookContacts,
          selectedChannels,
          selectedChats,
          selectedGoogleMyDrive,
          selectedSharedDrives
        )}
        onBack={() => setView("review")}
        onStarted={(operationId) => {
          if (onCleanupStarted) {
            onCleanupStarted(operationId);
          } else {
            // Fallback for when no redirect handler is provided — normally unreachable once App.tsx
            // always passes onCleanupStarted, kept so this component still works standalone.
            setCleanupOperationId(operationId);
            setView("cleanupProgress");
          }
        }}
      />
    );
  }

  if (view === "cleanupProgress" && cleanupOperationId) {
    return <CleanupProgressView operationId={cleanupOperationId} onFinished={() => setView("cleanupResults")} />;
  }

  if (view === "cleanupResults" && cleanupOperationId) {
    return (
      <CleanupResultsView
        operationId={cleanupOperationId}
        onRetried={(newOperationId) => {
          setCleanupOperationId(newOperationId);
          setView("cleanupProgress");
        }}
        onDone={() => {
          setSelectedOneDrive(new Map());
          setSelectedSharePoint(new Map());
          setSelectedOutlook(new Map());
          setSelectedOutlookCalendar(new Map());
          setSelectedOutlookContacts(new Map());
          setSelectedChannels(new Map());
          setSelectedChats(new Map());
          setSelectedGoogleMyDrive(new Map());
          setSelectedSharedDrives(new Map());
          setCleanupOperationId(null);
          setView("dashboard");
        }}
      />
    );
  }

  // Only shown on the actual selection tables — not on the dashboard or landing page, where it
  // would float over content with no table underneath it for the selection to relate to.
  const showSelectionBar =
    (view === "onedrive" ||
      view === "sharepoint" ||
      view === "outlook" ||
      view === "teams" ||
      view === "google_my_drive" ||
      view === "shared_drive") &&
    hasSelection(totals);

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
          <>
            {reconciliationBanner && (
              <div className="mb-4 flex items-start justify-between gap-3 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
                <span>{reconciliationBanner}</span>
                <button onClick={() => setReconciliationBanner(null)} className="shrink-0 font-medium text-amber-700 hover:text-amber-900">
                  Dismiss
                </button>
              </div>
            )}
            <Dashboard
              group={activeGroup}
              onOpenOneDrive={() => setView("onedrive")}
              onOpenSharePoint={() => setView("sharepoint")}
              onOpenTeams={() => setView("teams")}
              onOpenOutlook={() => setView("outlook")}
              onOpenGoogleMyDrive={() => setView("google_my_drive")}
              onOpenSharedDrives={() => setView("shared_drive")}
              onSyncFinished={handleSyncFinished}
            />
          </>
        )}

        {view === "onedrive" && activeGroup?.onedrive && (
          <OneDriveView connectionId={activeGroup.onedrive.id} selected={selectedOneDrive} setSelected={setSelectedOneDrive} />
        )}

        {view === "sharepoint" && activeGroup?.sharepoint && (
          <SharePointView connectionId={activeGroup.sharepoint.id} selected={selectedSharePoint} setSelected={setSelectedSharePoint} />
        )}

        {view === "outlook" && activeGroup?.outlook && (
          <OutlookView
            connectionId={activeGroup.outlook.id}
            selectedMail={selectedOutlook}
            setSelectedMail={setSelectedOutlook}
            selectedCalendar={selectedOutlookCalendar}
            setSelectedCalendar={setSelectedOutlookCalendar}
            selectedContacts={selectedOutlookContacts}
            setSelectedContacts={setSelectedOutlookContacts}
          />
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

        {view === "google_my_drive" && activeGroup?.google_my_drive && (
          <GoogleMyDriveView connectionId={activeGroup.google_my_drive.id} selected={selectedGoogleMyDrive} setSelected={setSelectedGoogleMyDrive} />
        )}

        {view === "shared_drive" && activeGroup?.shared_drive && (
          <SharedDrivesView connectionId={activeGroup.shared_drive.id} selected={selectedSharedDrives} setSelected={setSelectedSharedDrives} />
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
        No clouds connected yet — connect one from the Clouds tab first.
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

function ServiceCard({
  icon,
  name,
  stats,
  action,
  onClick,
  disabled,
  syncControl,
}: {
  icon: string;
  name: string;
  stats: React.ReactNode;
  action: string;
  onClick: () => void;
  disabled?: boolean;
  syncControl?: React.ReactNode;
}) {
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
      {syncControl}
    </div>
  );
}

/** ✓ / ⏳ / ✗ plus a running "X of Y" count while in flight, e.g. "⏳ 432/785" — so there's at least a concrete sense of how much is left, not just a spinner with no scale. */
function syncResourceLabel(
  r: { status: CleaningSyncResourceStatus; processed: number; total: number; unavailableCount?: number } | undefined
): string {
  if (!r) return "";
  const isLive = r.status === "queued" || r.status === "running";
  if (isLive) return r.total > 0 ? `⏳ ${r.processed.toLocaleString()}/${r.total.toLocaleString()}` : "⏳";
  if (r.status === "completed") return "✓";
  if (r.status === "completed_with_errors") {
    // "completed_with_errors" here almost always just means some accounts have no provisioned
    // drive, or Microsoft blocked access to a specific site — not that the sync itself broke. A
    // plain ✗ would read as "this failed", so use a warning instead and say how many, when known.
    return r.unavailableCount ? `⚠ ${r.unavailableCount.toLocaleString()} unavailable` : "⚠";
  }
  return "✗"; // genuinely 'failed' or 'cancelled' — the sync job itself didn't complete
}

const DISCOVERING_LABEL: Record<"onedrive" | "sharepoint" | "teams" | "outlook", string> = {
  onedrive: "Finding accounts…",
  sharepoint: "Finding sites…",
  teams: "Finding teams…",
  // Sync Now bundles Mail + Calendar + Contacts into this one connection's sync_jobs row (see
  // cloudSyncWorker.ts's syncOutlook) — this label covers all three phases, not just Mail.
  outlook: "Finding mailboxes, calendars, and contacts…",
};

/**
 * Visual progress bar for an in-flight sync — determinate (fills to processed/total) once the
 * backend has finished enumerating and knows a real total; an animated indeterminate sweep before
 * that, so a large tenant's enumeration phase (which can itself take a while) doesn't look stalled.
 */
function SyncProgressBar({ total, processed }: { total: number; processed: number }) {
  if (total <= 0) {
    return (
      <div className="relative h-1.5 w-full overflow-hidden rounded-full bg-slate-100">
        <div className="absolute inset-y-0 w-1/3 animate-indeterminate-bar rounded-full bg-[#1b2fc4]" />
      </div>
    );
  }
  const pct = Math.min(100, Math.round((processed / total) * 100));
  return (
    <div className="h-1.5 w-full overflow-hidden rounded-full bg-slate-100">
      <div className="h-full rounded-full bg-[#1b2fc4] transition-[width] duration-500 ease-out" style={{ width: `${pct}%` }} />
    </div>
  );
}

/** A rough "N minutes left" from the rate observed so far — never shown until there's at least one completed item to base a rate on, so it's never just a guess dressed up as a number. */
function estimateSyncTimeRemaining(startedAt: string, processed: number, total: number): string | null {
  if (processed <= 0 || total <= 0 || processed >= total) return null;
  const elapsedMs = Date.now() - new Date(startedAt).getTime();
  if (elapsedMs <= 0) return null;
  const remainingMs = (elapsedMs / processed) * (total - processed);
  const remainingMin = Math.round(remainingMs / 60_000);
  if (remainingMin < 1) return "less than a minute left";
  return remainingMin === 1 ? "about 1 minute left" : `about ${remainingMin} minutes left`;
}

/**
 * Self-contained per-cloud sync control — each of OneDrive/SharePoint/Teams gets its own instance,
 * its own "Last synced" timestamp, and its own independent Sync button, so syncing one never blocks
 * or bundles in the others; the user decides exactly what to sync. `startSync` is called with only
 * this one connectionId (the backend already treats sync concurrency per-connection, not per-tenant,
 * for this exact reason).
 */
function CloudSyncControl({
  connectionId,
  cloudType,
  lastSyncedAt,
  onSyncFinished,
}: {
  connectionId: string;
  /** Which of the operation's byResource slots belongs to this card — must be explicit, not guessed, since an operation can (e.g. an older bundled sync) have more than one slot populated. */
  cloudType: "onedrive" | "sharepoint" | "teams" | "outlook";
  lastSyncedAt: string | null;
  onSyncFinished: () => void;
}) {
  const [syncOperationId, setSyncOperationId] = useState<string | null>(null);
  const [syncOperation, setSyncOperation] = useState<CleaningSyncOperation | null>(null);
  const [syncError, setSyncError] = useState<string | null>(null);

  const isSyncing = syncOperationId != null && (!syncOperation || syncOperation.status === "queued" || syncOperation.status === "running");
  // A single-connection sync only ever populates exactly one of the three byResource slots.
  const resourceStatus = syncOperation?.byResource[cloudType];

  // Resumes tracking on mount (navigating away and back, or reloading, shouldn't make an
  // in-progress or just-finished sync look like it never happened — see the Sync Now bugfix notes).
  useEffect(() => {
    getLatestSyncOperation([connectionId])
      .then(({ operation }) => {
        if (!operation) return;
        if (operation.status === "queued" || operation.status === "running") {
          setSyncOperationId(operation.id);
        } else {
          setSyncOperation(operation);
          onSyncFinished();
        }
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [connectionId]);

  useEffect(() => {
    if (!syncOperationId) return;
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      const op = await getSyncOperation(syncOperationId!).catch(() => null);
      if (cancelled || !op) return;
      setSyncOperation(op);
      if (op.status === "queued" || op.status === "running") {
        // Faster than the old 3s cadence — this drives a live progress bar now, not just a static
        // "Syncing…" label, so it should feel like it's actually moving.
        timer = setTimeout(poll, 1500);
      } else {
        onSyncFinished();
      }
    }
    poll();
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [syncOperationId]);

  async function handleSync() {
    setSyncError(null);
    setSyncOperation(null);
    try {
      const { operationId } = await startSync([connectionId]);
      setSyncOperationId(operationId);
    } catch (err) {
      setSyncError(err instanceof ApiClientError ? err.message : "Couldn't start sync. Try again.");
    }
  }

  const finishedClassName =
    resourceStatus?.status === "completed" ? "text-emerald-600" : resourceStatus?.status === "completed_with_errors" ? "text-amber-600" : "text-rose-600";

  return (
    <div className="mt-3 border-t border-slate-100 pt-3">
      <div className="flex items-center justify-between gap-2">
        <div className="text-xs text-slate-400">{lastSyncedAt ? `Last synced ${formatDate(lastSyncedAt)}` : "Not synced yet"}</div>
        <button
          onClick={handleSync}
          disabled={isSyncing}
          className="shrink-0 rounded-md border border-slate-200 px-2 py-1 text-xs font-semibold text-[#1b2fc4] hover:bg-[#1b2fc4]/5 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {isSyncing ? "Syncing…" : "Sync"}
        </button>
      </div>
      {isSyncing && (
        <div className="mt-2">
          <SyncProgressBar total={resourceStatus?.total ?? 0} processed={resourceStatus?.processed ?? 0} />
          <div className="mt-1 flex items-center justify-between gap-2 text-xs text-slate-500">
            <span>
              {resourceStatus && resourceStatus.total > 0
                ? `Syncing ${resourceStatus.processed.toLocaleString()} / ${resourceStatus.total.toLocaleString()}`
                : DISCOVERING_LABEL[cloudType]}
            </span>
            {resourceStatus && resourceStatus.total > 0 && syncOperation && (
              <span className="shrink-0">{estimateSyncTimeRemaining(syncOperation.startedAt, resourceStatus.processed, resourceStatus.total)}</span>
            )}
          </div>
        </div>
      )}
      {!isSyncing && resourceStatus && <p className={`mt-1 text-xs ${finishedClassName}`}>{syncResourceLabel(resourceStatus)}</p>}
      {syncError && <p className="mt-1 text-xs text-rose-600">{syncError}</p>}
    </div>
  );
}

function Dashboard({
  group,
  onOpenOneDrive,
  onOpenSharePoint,
  onOpenTeams,
  onOpenOutlook,
  onOpenGoogleMyDrive,
  onOpenSharedDrives,
  onSyncFinished,
}: {
  group: TenantGroup;
  onOpenOneDrive: () => void;
  onOpenSharePoint: () => void;
  onOpenTeams: () => void;
  onOpenOutlook: () => void;
  onOpenGoogleMyDrive: () => void;
  onOpenSharedDrives: () => void;
  onSyncFinished: () => void;
}) {
  const [oneDriveTotals, setOneDriveTotals] = useState<{ count: number; bytes: number } | null>(null);
  const [sharePointTotals, setSharePointTotals] = useState<{ count: number; bytes: number } | null>(null);
  const [outlookTotals, setOutlookTotals] = useState<{ count: number; items: number } | null>(null);
  const [outlookCalendarTotals, setOutlookCalendarTotals] = useState<{ events: number } | null>(null);
  const [outlookContactTotals, setOutlookContactTotals] = useState<{ contacts: number } | null>(null);
  const [teamsSummary, setTeamsSummary] = useState<CleaningTeamsSummary | null>(null);
  const [googleMyDriveTotals, setGoogleMyDriveTotals] = useState<{ count: number; bytes: number } | null>(null);
  const [sharedDrivesTotals, setSharedDrivesTotals] = useState<{ count: number; bytes: number } | null>(null);

  const refreshTotals = useCallback(() => {
    if (group.onedrive) {
      listOneDriveAccounts(group.onedrive.id, { sort: "storage", pageSize: 200 }).then(({ accounts, total }) => {
        // `total` is the accurate count; the byte sum is over the (large) page fetched here, which
        // covers the dashboard headline for realistically-sized tenants — the full table (with real
        // pagination) is what "View Accounts" opens.
        setOneDriveTotals({ count: total, bytes: accounts.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
    if (group.google_my_drive) {
      listGoogleMyDriveAccounts(group.google_my_drive.id, { sort: "storage", pageSize: 200 }).then(({ accounts, total }) => {
        setGoogleMyDriveTotals({ count: total, bytes: accounts.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
    if (group.shared_drive) {
      listSharedDrives(group.shared_drive.id, { sort: "storage", pageSize: 200 }).then(({ drives, total }) => {
        setSharedDrivesTotals({ count: total, bytes: drives.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
    if (group.sharepoint) {
      listSharePointSites(group.sharepoint.id, { sort: "storage", pageSize: 200 }).then(({ sites, total }) => {
        setSharePointTotals({ count: total, bytes: sites.reduce((s, a) => s + a.storageUsedBytes, 0) });
      });
    }
    if (group.outlook) {
      // Three separate calls to three separate dedicated endpoints, each with its own state — same
      // as the rest of this component, Mail/Calendar/Contacts are never funnelled through one shared
      // "which resource" fetch or merged into a single totals object with load-order dependencies.
      listOutlookMailboxes(group.outlook.id, { pageSize: 200 }).then(({ mailboxes, total }) => {
        setOutlookTotals({ count: total, items: mailboxes.reduce((s, m) => s + m.itemCount, 0) });
      });
      listOutlookCalendars(group.outlook.id, { pageSize: 200 }).then(({ calendars }) => {
        setOutlookCalendarTotals({ events: calendars.reduce((s, c) => s + c.itemCount, 0) });
      });
      listOutlookContacts(group.outlook.id, { pageSize: 200 }).then(({ contacts }) => {
        setOutlookContactTotals({ contacts: contacts.reduce((s, c) => s + c.itemCount, 0) });
      });
    }
  }, [group]);

  useEffect(refreshTotals, [refreshTotals]);

  // Called by whichever CloudSyncControl (OneDrive/SharePoint/Teams) just finished — cheap to
  // refresh all of this dashboard's own totals regardless of which one it was, and the parent page
  // always needs to re-check lastSyncedAt/reconcile the selection either way.
  function handleCloudSyncFinished() {
    refreshTotals();
    if (group.teams) getTeamsSummary(group.teams.id).then(setTeamsSummary).catch(() => {});
    onSyncFinished();
  }

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
          syncControl={
            group.onedrive && (
              <CloudSyncControl connectionId={group.onedrive.id} cloudType="onedrive" lastSyncedAt={group.onedrive.lastSyncedAt} onSyncFinished={handleCloudSyncFinished} />
            )
          }
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
          syncControl={
            group.sharepoint && (
              <CloudSyncControl connectionId={group.sharepoint.id} cloudType="sharepoint" lastSyncedAt={group.sharepoint.lastSyncedAt} onSyncFinished={handleCloudSyncFinished} />
            )
          }
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
          syncControl={
            group.teams && (
              <CloudSyncControl connectionId={group.teams.id} cloudType="teams" lastSyncedAt={group.teams.lastSyncedAt} onSyncFinished={handleCloudSyncFinished} />
            )
          }
        />
        <ServiceCard
          icon="📧"
          name="Outlook"
          stats={
            group.outlook ? (
              outlookTotals ? (
                <>
                  <div>{outlookTotals.count.toLocaleString()}{outlookTotals.count > 0 ? "+" : ""} mailboxes</div>
                  <div>{outlookTotals.items.toLocaleString()} mail items</div>
                  <div>{outlookCalendarTotals ? outlookCalendarTotals.events.toLocaleString() : "…"} calendar events</div>
                  <div>{outlookContactTotals ? outlookContactTotals.contacts.toLocaleString() : "…"} contacts</div>
                </>
              ) : (
                <div className="italic text-slate-400">Loading…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Outlook"
          onClick={onOpenOutlook}
          disabled={!group.outlook}
          syncControl={
            group.outlook && (
              <CloudSyncControl connectionId={group.outlook.id} cloudType="outlook" lastSyncedAt={group.outlook.lastSyncedAt} onSyncFinished={handleCloudSyncFinished} />
            )
          }
        />
        <ServiceCard
          icon="🔷"
          name="Google My Drive"
          stats={
            group.google_my_drive ? (
              googleMyDriveTotals ? (
                <>
                  <div>{googleMyDriveTotals.count.toLocaleString()}{googleMyDriveTotals.count > 0 ? "+" : ""} accounts</div>
                  <div>{formatBytes(googleMyDriveTotals.bytes)} used</div>
                </>
              ) : (
                <div className="italic text-slate-400">Loading…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Accounts"
          onClick={onOpenGoogleMyDrive}
          disabled={!group.google_my_drive}
          // This dashboard's own inline "Sync Now" control (CloudSyncControl) wraps a separate
          // tenant-level multi-resource sync_operations table (onedrive/sharepoint/outlook/teams
          // fixed columns) not yet extended for Google in this pass — deferred, see
          // docs/google-workspace-integration.md. My Drive syncing already works end-to-end via
          // Manage Clouds' per-connection Resync (routes/cloudConnections.ts), same as every other
          // workload; only this convenience shortcut is missing.
          syncControl={
            group.google_my_drive && (
              <span className="text-xs text-slate-400">Use Manage Clouds → Resync to sync</span>
            )
          }
        />
        <ServiceCard
          icon="🗄️"
          name="Shared Drives"
          stats={
            group.shared_drive ? (
              sharedDrivesTotals ? (
                <>
                  <div>{sharedDrivesTotals.count.toLocaleString()}{sharedDrivesTotals.count > 0 ? "+" : ""} drives</div>
                  <div>{formatBytes(sharedDrivesTotals.bytes)} used</div>
                </>
              ) : (
                <div className="italic text-slate-400">Loading…</div>
              )
            ) : (
              <div className="text-slate-400">Not connected</div>
            )
          }
          action="View Drives"
          onClick={onOpenSharedDrives}
          disabled={!group.shared_drive}
          // Same deferred inline "Sync Now" shortcut as Google My Drive above — see that card's comment.
          syncControl={
            group.shared_drive && <span className="text-xs text-slate-400">Use Manage Clouds → Resync to sync</span>
          }
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

/**
 * The storage figure alone can't tell a viewer whether it's already caught up with a recent
 * permanent deletion — Microsoft's own storage-quota recalculation runs asynchronously on their
 * backend and can lag a real deletion by minutes or longer, so even a sync that ran *after* the
 * delete can still return the old number. There's no "just sync again" fix for that, so
 * row.deletionRecalcHint exists purely to set the right expectation here instead: a direct note for
 * the common first-24h case, tapering to a softer "verify independently" suggestion for the rare
 * case where Microsoft still hasn't caught up a week later (see routes/cleaning.ts).
 */
function StorageUsedCell({ row }: { row: CleaningResourceRow }) {
  const lastSyncedNote = row.lastSyncedAt ? `Last synced ${formatDate(row.lastSyncedAt)}. ` : "";
  return (
    <div>
      <div>{formatBytes(row.storageUsedBytes)}</div>
      {row.deletionRecalcHint === "recent" && (
        <div
          className="text-xs text-amber-600"
          title={`${lastSyncedNote}Microsoft can take a while to recalculate storage after a deletion — this may not be caught up yet.`}
        >
          May not reflect a recent deletion yet
        </div>
      )}
      {row.deletionRecalcHint === "verify" && (
        <div
          className="text-xs text-slate-400"
          title={`${lastSyncedNote}It's been a while since the last permanent deletion here — Microsoft usually catches up within a day, but doesn't guarantee a timeline. If this figure still looks wrong, check the Microsoft 365 admin center directly.`}
        >
          Still looks off? Check the admin center
        </div>
      )}
    </div>
  );
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
    { label: "Storage Used", align: "right", render: (r) => <StorageUsedCell row={r} /> },
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

/** Identical shape to OneDriveView above — same CleaningResourceRow, same StorageUsedCell (including its deletionRecalcHint lag note, which listCleaningResources already computes generically for google_my_drive_account). */
function GoogleMyDriveView({ connectionId, selected, setSelected }: { connectionId: string; selected: Map<string, CleaningResourceRow>; setSelected: (m: Map<string, CleaningResourceRow>) => void }) {
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState<"storage" | "name">("storage");
  const fetcher = useCallback(
    (opts: { search?: string; sort?: "storage" | "name"; page?: number; pageSize?: number }) =>
      listGoogleMyDriveAccounts(connectionId, opts).then((r) => ({ rows: r.accounts, total: r.total, page: r.page, pageSize: r.pageSize })),
    [connectionId]
  );
  const { rows, loading, error, page, totalPages, total, goToPage } = usePagedList(fetcher, search, sort);

  const columns: DiscoveryColumn<CleaningResourceRow>[] = [
    { label: "User Name", render: (r) => r.name },
    { label: "User Email", render: (r) => r.detail },
    { label: "Storage Used", align: "right", render: (r) => <StorageUsedCell row={r} /> },
    { label: "Status", render: (r) => (r.status === "failed" ? <span className="text-rose-500">Unavailable</span> : "Ready") },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Google My Drive</h2>
      <p className="mb-5 text-sm text-slate-500">All Workspace user accounts and their storage usage</p>
      <DiscoveryTable
        title="Google My Drive Accounts"
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
        emptyMessage="No Google My Drive accounts found."
      />
    </div>
  );
}

/** Same shape as GoogleMyDriveView above — Shared Drives reuses the identical CleaningResourceRow/listCleaningResources plumbing, just a different discovery endpoint and column labels (a drive's "name" is its only identifier, no secondary email/URL). */
function SharedDrivesView({ connectionId, selected, setSelected }: { connectionId: string; selected: Map<string, CleaningResourceRow>; setSelected: (m: Map<string, CleaningResourceRow>) => void }) {
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState<"storage" | "name">("storage");
  const fetcher = useCallback(
    (opts: { search?: string; sort?: "storage" | "name"; page?: number; pageSize?: number }) =>
      listSharedDrives(connectionId, opts).then((r) => ({ rows: r.drives, total: r.total, page: r.page, pageSize: r.pageSize })),
    [connectionId]
  );
  const { rows, loading, error, page, totalPages, total, goToPage } = usePagedList(fetcher, search, sort);

  const columns: DiscoveryColumn<CleaningResourceRow>[] = [
    { label: "Shared Drive Name", render: (r) => r.name },
    { label: "Storage Used", align: "right", render: (r) => <StorageUsedCell row={r} /> },
    { label: "Status", render: (r) => (r.status === "failed" ? <span className="text-rose-500">Unavailable</span> : "Ready") },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Shared Drives</h2>
      <p className="mb-5 text-sm text-slate-500">All shared drives and their storage usage — cleanup removes content only, never the drive itself</p>
      <DiscoveryTable
        title="Shared Drives"
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
        searchPlaceholder="Search drives…"
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
        emptyMessage="No Shared Drives found."
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
    { label: "Storage Used", align: "right", render: (r) => <StorageUsedCell row={r} /> },
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

interface OutlookMailboxRow extends OutlookMailboxOverviewRow {
  id: string; // = upn, stable/unique per mailbox within a connection
}

/** A sub-resource's Graph-facing row, in the shape every selection Map already expects (see buildCleanupManifest). */
function toSelectionRow(mailbox: OutlookMailboxRow, sub: OutlookOverviewSubResource): CleaningResourceRow {
  return {
    id: sub.id,
    name: mailbox.name,
    detail: mailbox.upn,
    storageUsedBytes: 0,
    itemCount: sub.itemCount,
    status: sub.status,
    lastSyncedAt: null,
    deletionRecalcHint: null,
  };
}

/** A row counts as "fully selected" (drives the master checkbox) only when every sub-resource it actually has is selected — a mailbox with no synced Calendar yet doesn't block being "fully" selected on Mail+Contacts. */
function isRowFullySelected(
  row: OutlookMailboxRow,
  selectedMail: Map<string, CleaningResourceRow>,
  selectedCalendar: Map<string, CleaningResourceRow>,
  selectedContacts: Map<string, CleaningResourceRow>
): boolean {
  if (!selectedMail.has(row.mail.id)) return false;
  if (row.calendar && !selectedCalendar.has(row.calendar.id)) return false;
  if (row.contacts && !selectedContacts.has(row.contacts.id)) return false;
  return true;
}

/**
 * One row per mailbox, purely informational Mail/Calendar/Contacts counts, and a single row
 * checkbox — same shape as every other DiscoveryTable in this app (OneDrive/SharePoint/Teams-DMs).
 * Selecting a mailbox here includes all three categories for it by default. Per-category control
 * (excluding just Calendar for one mailbox, say) happens on the Review Selection page instead,
 * where each selected mailbox gets its own Mail/Calendar/Contacts checkboxes — see ReviewPage's
 * OutlookReviewSection. The three selection Maps stay exactly what buildCleanupManifest already
 * expects; only where the fine-tuning UI lives changed. Mail/Calendar/Contacts remain separate
 * resource types, separate manifest slots, separate execution paths throughout the rest of the app.
 */
function OutlookView({
  connectionId,
  selectedMail,
  setSelectedMail,
  selectedCalendar,
  setSelectedCalendar,
  selectedContacts,
  setSelectedContacts,
}: {
  connectionId: string;
  selectedMail: Map<string, CleaningResourceRow>;
  setSelectedMail: (m: Map<string, CleaningResourceRow>) => void;
  selectedCalendar: Map<string, CleaningResourceRow>;
  setSelectedCalendar: (m: Map<string, CleaningResourceRow>) => void;
  selectedContacts: Map<string, CleaningResourceRow>;
  setSelectedContacts: (m: Map<string, CleaningResourceRow>) => void;
}) {
  const [search, setSearch] = useState("");

  const fetcher = useCallback(
    (opts: { search?: string; page?: number; pageSize?: number }) =>
      listOutlookOverview(connectionId, opts).then((r) => ({
        rows: r.mailboxes.map((m) => ({ ...m, id: m.upn })),
        total: r.total,
        page: r.page,
        pageSize: r.pageSize,
      })),
    [connectionId]
  );
  const { rows, loading, error, page, totalPages, total, goToPage } = usePagedList(fetcher, search);

  function setRowResource(
    row: OutlookMailboxRow,
    sub: OutlookOverviewSubResource | null,
    selected: Map<string, CleaningResourceRow>,
    setSelected: (m: Map<string, CleaningResourceRow>) => void,
    include: boolean
  ) {
    if (!sub) return; // nothing synced for this category yet — nothing to (de)select
    const next = new Map(selected);
    include ? next.set(sub.id, toSelectionRow(row, sub)) : next.delete(sub.id);
    setSelected(next);
  }

  /** "Select All" — every category, every mailbox on the current page. Also backs DiscoveryTable's own header checkbox. */
  function handleToggleAll() {
    const allSelected = rows.every((r) => isRowFullySelected(r, selectedMail, selectedCalendar, selectedContacts));
    const nextMail = new Map(selectedMail);
    const nextCalendar = new Map(selectedCalendar);
    const nextContacts = new Map(selectedContacts);
    for (const row of rows) {
      const include = !allSelected;
      include ? nextMail.set(row.mail.id, toSelectionRow(row, row.mail)) : nextMail.delete(row.mail.id);
      if (row.calendar) (include ? nextCalendar.set(row.calendar.id, toSelectionRow(row, row.calendar)) : nextCalendar.delete(row.calendar.id));
      if (row.contacts) (include ? nextContacts.set(row.contacts.id, toSelectionRow(row, row.contacts)) : nextContacts.delete(row.contacts.id));
    }
    setSelectedMail(nextMail);
    setSelectedCalendar(nextCalendar);
    setSelectedContacts(nextContacts);
  }

  function countCell(sub: OutlookOverviewSubResource | null): React.ReactNode {
    if (!sub) return <span className="text-slate-300">—</span>;
    return <span className={sub.status === "failed" ? "text-rose-500" : "text-slate-700"}>{sub.itemCount.toLocaleString()}</span>;
  }

  const columns: DiscoveryColumn<OutlookMailboxRow>[] = [
    { label: "User Name", render: (r) => r.name },
    { label: "User Email", render: (r) => r.upn },
    { label: "Mail Items", align: "right", render: (r) => countCell(r.mail) },
    { label: "Calendar Events", align: "right", render: (r) => countCell(r.calendar) },
    { label: "Contacts", align: "right", render: (r) => countCell(r.contacts) },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Outlook</h2>
      <p className="mb-5 text-sm text-slate-500">
        Mail, calendar events, and contacts for every mailbox. Select a mailbox here, then choose which of Mail/Calendar/Contacts
        to actually clean on the Review Selection page. Never touches mail folders, calendars, or contact folders themselves.
      </p>

      <DiscoveryTable
        title="Outlook Mailboxes"
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
        selected={new Set(rows.filter((r) => isRowFullySelected(r, selectedMail, selectedCalendar, selectedContacts)).map((r) => r.id))}
        onToggle={(id) => {
          const row = rows.find((r) => r.id === id);
          if (!row) return;
          const include = !isRowFullySelected(row, selectedMail, selectedCalendar, selectedContacts);
          setRowResource(row, row.mail, selectedMail, setSelectedMail, include);
          setRowResource(row, row.calendar, selectedCalendar, setSelectedCalendar, include);
          setRowResource(row, row.contacts, selectedContacts, setSelectedContacts, include);
        }}
        onToggleAll={handleToggleAll}
        emptyMessage="No Outlook mailboxes found."
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

/** One row per mailbox appearing in any of the three Outlook selection Maps, grouped by upn (CleaningResourceRow.detail) since each Map only knows its own resource's id — reconstructed here purely for display/editing on this page. */
interface OutlookReviewRow {
  upn: string;
  name: string;
  mail: CleaningResourceRow | null;
  calendar: CleaningResourceRow | null;
  contacts: CleaningResourceRow | null;
}

function buildOutlookReviewRows(
  selectedMail: Map<string, CleaningResourceRow>,
  selectedCalendar: Map<string, CleaningResourceRow>,
  selectedContacts: Map<string, CleaningResourceRow>
): OutlookReviewRow[] {
  const byUpn = new Map<string, OutlookReviewRow>();
  function upsert(row: CleaningResourceRow, slot: "mail" | "calendar" | "contacts") {
    const existing = byUpn.get(row.detail) ?? { upn: row.detail, name: row.name, mail: null, calendar: null, contacts: null };
    existing[slot] = row;
    byUpn.set(row.detail, existing);
  }
  for (const row of selectedMail.values()) upsert(row, "mail");
  for (const row of selectedCalendar.values()) upsert(row, "calendar");
  for (const row of selectedContacts.values()) upsert(row, "contacts");
  return [...byUpn.values()].sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Per-mailbox Mail/Calendar/Contacts checkboxes — this is where per-mailbox control actually lives
 * (not the Outlook discovery table, which only offers one row checkbox that includes all three by
 * default). Unchecking one here removes just that entry from its own selection Map; the other two
 * Maps, and every other cloud's selection, are untouched.
 */
function OutlookReviewSection({
  selectedMail,
  setSelectedMail,
  selectedCalendar,
  setSelectedCalendar,
  selectedContacts,
  setSelectedContacts,
}: {
  selectedMail: Map<string, CleaningResourceRow>;
  setSelectedMail: (m: Map<string, CleaningResourceRow>) => void;
  selectedCalendar: Map<string, CleaningResourceRow>;
  setSelectedCalendar: (m: Map<string, CleaningResourceRow>) => void;
  selectedContacts: Map<string, CleaningResourceRow>;
  setSelectedContacts: (m: Map<string, CleaningResourceRow>) => void;
}) {
  const rows = buildOutlookReviewRows(selectedMail, selectedCalendar, selectedContacts);
  if (rows.length === 0) return null;

  function toggle(row: CleaningResourceRow | null, selected: Map<string, CleaningResourceRow>, setSelected: (m: Map<string, CleaningResourceRow>) => void) {
    if (!row) return;
    const next = new Map(selected);
    next.delete(row.id);
    setSelected(next);
  }

  return (
    <div className="border-b border-slate-100 pb-3">
      <span className="text-sm font-medium text-slate-700">Outlook</span>
      <div className="mt-3 space-y-2">
        {rows.map((row) => (
          <div key={row.upn} className="flex items-center justify-between rounded-md bg-slate-50 px-3 py-2">
            <div className="text-sm text-slate-700">
              {row.name} <span className="text-slate-400">· {row.upn}</span>
            </div>
            <div className="flex items-center gap-4 text-sm text-slate-600">
              <label className={`flex items-center gap-1.5 ${row.mail ? "" : "opacity-30"}`}>
                <input type="checkbox" checked={row.mail != null} disabled={!row.mail} onChange={() => toggle(row.mail, selectedMail, setSelectedMail)} />
                Mail
              </label>
              <label className={`flex items-center gap-1.5 ${row.calendar ? "" : "opacity-30"}`}>
                <input type="checkbox" checked={row.calendar != null} disabled={!row.calendar} onChange={() => toggle(row.calendar, selectedCalendar, setSelectedCalendar)} />
                Calendar
              </label>
              <label className={`flex items-center gap-1.5 ${row.contacts ? "" : "opacity-30"}`}>
                <input type="checkbox" checked={row.contacts != null} disabled={!row.contacts} onChange={() => toggle(row.contacts, selectedContacts, setSelectedContacts)} />
                Contacts
              </label>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function ReviewPage({
  totals,
  selectedOutlookMail,
  setSelectedOutlookMail,
  selectedOutlookCalendar,
  setSelectedOutlookCalendar,
  selectedOutlookContacts,
  setSelectedOutlookContacts,
  onBack,
  onContinue,
}: {
  totals: SelectionTotals;
  selectedOutlookMail: Map<string, CleaningResourceRow>;
  setSelectedOutlookMail: (m: Map<string, CleaningResourceRow>) => void;
  selectedOutlookCalendar: Map<string, CleaningResourceRow>;
  setSelectedOutlookCalendar: (m: Map<string, CleaningResourceRow>) => void;
  selectedOutlookContacts: Map<string, CleaningResourceRow>;
  setSelectedOutlookContacts: (m: Map<string, CleaningResourceRow>) => void;
  onBack: () => void;
  onContinue: () => void;
}) {
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
        <OutlookReviewSection
          selectedMail={selectedOutlookMail}
          setSelectedMail={setSelectedOutlookMail}
          selectedCalendar={selectedOutlookCalendar}
          setSelectedCalendar={setSelectedOutlookCalendar}
          selectedContacts={selectedOutlookContacts}
          setSelectedContacts={setSelectedOutlookContacts}
        />
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

      <div className="mt-6 flex items-center gap-3">
        <button onClick={onBack} className="rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
          ← Back
        </button>
        <button
          onClick={onContinue}
          disabled={!hasSelection(totals)}
          className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
        >
          Continue to Cleanup →
        </button>
      </div>
    </div>
  );
}
