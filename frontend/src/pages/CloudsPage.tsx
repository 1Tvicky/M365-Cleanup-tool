import { useCallback, useEffect, useRef, useState } from "react";
import { CloudTileGrid } from "../components/clouds/CloudTileGrid";
import { ManageClouds } from "../components/clouds/ManageClouds";
import { UserMenu } from "../components/layout/UserMenu";
import {
  disconnectCloudConnection,
  initCloudConnect,
  listManageClouds,
  openConnectPopup,
  resyncCloudConnection,
  type ManageCloudsRow,
} from "../api/clouds";
import { ApiClientError } from "../api/client";
import type { OperatorSummary } from "../api/auth";
import type { Workload } from "../types";

type Tab = "add" | "manage";

// Any connection in one of these states has an active enumeration job worth polling for.
function hasLiveWork(rows: ManageCloudsRow[]): boolean {
  return rows.some((r) => r.status === "connecting" || (r.totalUsers > 0 && r.processedUsers < r.totalUsers));
}

export function CloudsPage({ operator, onLogout }: { operator: OperatorSummary; onLogout: () => void }) {
  const [tab, setTab] = useState<Tab>("add");
  const [connections, setConnections] = useState<ManageCloudsRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [connectingType, setConnectingType] = useState<Workload | null>(null);
  const [resyncingId, setResyncingId] = useState<string | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const refresh = useCallback(async () => {
    try {
      const { connections: rows } = await listManageClouds();
      setConnections(rows);
    } catch {
      // A poll tick failing silently is fine — the next one retries; only the initial load shows an error state.
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
  }, [refresh]);

  // Polls while any connection has an in-progress enumeration job, so the progress bar/percent
  // actually moves without the operator manually refreshing — matches the reference product's
  // "job continues running after the row appears" behavior.
  useEffect(() => {
    if (pollRef.current) clearInterval(pollRef.current);
    if (hasLiveWork(connections)) {
      pollRef.current = setInterval(refresh, 3000);
    }
    return () => {
      if (pollRef.current) clearInterval(pollRef.current);
    };
  }, [connections, refresh]);

  function showToast(message: string) {
    setToast(message);
    setTimeout(() => setToast(null), 3000);
  }

  async function handleConnect(workload: Workload) {
    if (connectingType) return;
    setConnectingType(workload);
    try {
      const { authorizeUrl } = await initCloudConnect(workload);
      const result = await openConnectPopup(authorizeUrl);
      if (result.status === "success") {
        showToast("Connected — enumeration is running in the background.");
        setTab("manage");
        await refresh();
      } else if (result.reason !== "closed") {
        showToast(`Couldn't connect: ${result.reason}`);
      }
    } catch (err) {
      showToast(err instanceof ApiClientError ? err.message : "Couldn't start the connection. Try again.");
    } finally {
      setConnectingType(null);
    }
  }

  async function handleResync(connectionId: string) {
    setResyncingId(connectionId);
    try {
      await resyncCloudConnection(connectionId);
      await refresh();
    } catch (err) {
      showToast(err instanceof ApiClientError ? err.message : "Couldn't start a resync.");
    } finally {
      setResyncingId(null);
    }
  }

  async function handleDisconnect(connectionId: string) {
    try {
      await disconnectCloudConnection(connectionId);
      await refresh();
    } catch (err) {
      showToast(err instanceof ApiClientError ? err.message : "Couldn't disconnect.");
    }
  }

  return (
    <div>
      <div className="flex items-center justify-between border-b border-slate-200 bg-white px-8">
        <div className="flex gap-6">
          <TabButton label="Add clouds" active={tab === "add"} onClick={() => setTab("add")} />
          <TabButton label="Manage clouds" active={tab === "manage"} onClick={() => setTab("manage")} />
        </div>
        <div className="py-2.5">
          <UserMenu name={operator.displayName} onLogout={onLogout} />
        </div>
      </div>

      <div className="px-8 py-6">
        {tab === "add" ? (
          <CloudTileGrid onConnect={handleConnect} connectingType={connectingType} />
        ) : loading ? (
          <p className="text-sm text-slate-500">Loading…</p>
        ) : (
          <ManageClouds connections={connections} onResync={handleResync} onDisconnect={handleDisconnect} resyncingId={resyncingId} />
        )}
      </div>

      {toast && (
        <div className="fixed bottom-6 right-6 rounded-lg bg-slate-900 px-4 py-2.5 text-sm text-white shadow-lg">
          {toast}
        </div>
      )}
    </div>
  );
}

function TabButton({ label, active, onClick }: { label: string; active: boolean; onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className={`relative py-3 text-xs font-semibold uppercase tracking-wide transition-colors ${
        active ? "text-slate-900" : "text-slate-400 hover:text-slate-600"
      }`}
    >
      {label}
      {active && <span className="absolute inset-x-0 -bottom-px h-0.5 rounded-full bg-teal-400" />}
    </button>
  );
}
