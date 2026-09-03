import type { Tenant } from "../../types";

export function TenantSelector({
  tenants,
  selectedId,
  onSelect,
}: {
  tenants: Tenant[];
  selectedId: string | null;
  onSelect: (tenantId: string) => void;
}) {
  const connected = tenants.filter((t) => t.status === "connected");

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">1. Select tenant</h2>
      <p className="mb-4 text-sm text-slate-500">Choose the connected M365 tenant to run cleanup against.</p>
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
        {connected.map((t) => (
          <button
            key={t.id}
            onClick={() => onSelect(t.id)}
            className={`flex items-center justify-between rounded-lg border px-4 py-3 text-left transition-colors ${
              selectedId === t.id ? "border-brand-500 bg-brand-50" : "border-slate-200 bg-white hover:border-slate-300"
            }`}
          >
            <div>
              <div className="text-sm font-semibold text-slate-800">{t.displayName}</div>
              <div className="text-xs text-slate-500">{t.workloads.join(" · ")}</div>
            </div>
            {selectedId === t.id && <span className="text-brand-600">✓</span>}
          </button>
        ))}
        {connected.length === 0 && (
          <div className="rounded-lg border border-dashed border-slate-300 px-4 py-6 text-center text-sm text-slate-500 sm:col-span-2">
            No connected tenants yet — connect one from the Clouds tab first.
          </div>
        )}
      </div>
    </div>
  );
}
