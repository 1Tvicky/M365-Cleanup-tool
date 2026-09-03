import { useState } from "react";
import { CloudTileGrid } from "../components/clouds/CloudTileGrid";
import { ManageClouds } from "../components/clouds/ManageClouds";
import { UserMenu } from "../components/layout/UserMenu";
import { MOCK_TENANTS, MOCK_JOBS } from "../api/mockData";
import type { OperatorSummary } from "../api/auth";
import type { Tenant, Workload } from "../types";

type Tab = "add" | "manage";

export function CloudsPage({ operator, onLogout }: { operator: OperatorSummary; onLogout: () => void }) {
  const [tab, setTab] = useState<Tab>("add");
  const [tenants, setTenants] = useState<Tenant[]>(MOCK_TENANTS);
  const [toast, setToast] = useState<string | null>(null);

  function handleConnect(workload: Workload) {
    // In production this calls GET /api/v1/auth/consent-url and redirects the browser there.
    setToast(`Redirecting to Microsoft admin consent for ${workload}…`);
    setTimeout(() => setToast(null), 2500);
  }

  function handleDisconnect(tenantId: string) {
    setTenants((prev) => prev.map((t) => (t.id === tenantId ? { ...t, status: "disconnected" } : t)));
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
          <CloudTileGrid onConnect={handleConnect} />
        ) : (
          <ManageClouds tenants={tenants} jobs={MOCK_JOBS} onDisconnect={handleDisconnect} />
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
