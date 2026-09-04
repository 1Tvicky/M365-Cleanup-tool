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
  // A Set, not a single value: onboarding one cloud (e.g. OneDrive) must not block starting
  // another (e.g. SharePoint/Teams) — each tile only cares whether IT is mid-connect.
  const [connectingTypes, setConnectingTypes] = useState<Set<Workload>>(new Set());
  const [resyncingId, setResyncingId] = useState<string | null>(null);
  const [toast, setToast] = useState<{ message: string; ok: boolean } | null>(null);
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

  function showToast(message: string, ok = false) {
    setToast({ message, ok });
    setTimeout(() => setToast(null), 4000);
  }

  async function handleConnect(workload: Workload) {
    if (connectingTypes.has(workload)) return; // re-clicking the SAME tile mid-connect is a no-op; other tiles are unaffected
    setConnectingTypes((prev) => new Set(prev).add(workload));
    try {
      const { authorizeUrl } = await initCloudConnect(workload);
      const result = await openConnectPopup(authorizeUrl, workload);
      // The popup's postMessage handshake back to this window depends on the browser preserving
      // window.opener across Microsoft's entire real login+consent page sequence, which we don't
      // control and can't fully verify — a "closed" result here is ambiguous, not necessarily a
      // cancellation. The connections row is already created server-side by the callback before
      // the popup ever tries to message back, so re-check the real state instead of trusting the
      // message alone; this is what actually makes Manage Clouds reliable regardless of whether
      // the message got through.
      await refresh();
      if (result.status === "success") {
        showToast("Account added successfully.", true);
        setTab("manage");
      } else if (result.reason !== "closed") {
        showToast(`Couldn't connect: ${result.reason}`);
      } else {
        setTab("manage");
      }
    } catch (err) {
      showToast(err instanceof ApiClientError ? err.message : "Couldn't start the connection. Try again.");
    } finally {
      setConnectingTypes((prev) => {
        const next = new Set(prev);
        next.delete(workload);
        return next;
      });
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
          <CloudTileGrid onConnect={handleConnect} connectingTypes={connectingTypes} />
        ) : loading ? (
          <p className="text-sm text-slate-500">Loading…</p>
        ) : (
          <ManageClouds connections={connections} onResync={handleResync} onDisconnect={handleDisconnect} resyncingId={resyncingId} />
        )}
      </div>

      {toast && (
        <div
          className={`fixed right-6 top-20 flex items-center gap-3 rounded-md border-l-4 px-5 py-3 text-sm font-medium shadow-lg ${
            toast.ok ? "border-emerald-500 bg-emerald-50 text-emerald-800" : "border-rose-500 bg-rose-50 text-rose-800"
          }`}
        >
          {toast.ok && (
            <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-emerald-500 text-white">
              ✓
            </span>
          )}
          <span>{toast.message}</span>
          <button onClick={() => setToast(null)} aria-label="Dismiss" className="ml-2 text-current opacity-60 hover:opacity-100">
            ✕
          </button>
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
