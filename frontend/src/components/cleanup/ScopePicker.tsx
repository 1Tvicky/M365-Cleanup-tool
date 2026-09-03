import { useState } from "react";

export interface CleanupScopeForm {
  workloads: Set<"teams" | "onedrive" | "sharepoint">;
  cutoffDate: string;
  removeM365Groups: boolean;
  searchQuery: string;
}

const WORKLOAD_OPTIONS: { id: "teams" | "onedrive" | "sharepoint"; label: string; hint: string }[] = [
  { id: "teams", label: "Teams channels", hint: "Delete specific channels; private channels are not restorable" },
  { id: "onedrive", label: "OneDrive files", hint: "Delete files older than a migration cutoff date" },
  { id: "sharepoint", label: "SharePoint libraries", hint: "Delete specific document libraries" },
];

export function ScopePicker({
  form,
  onChange,
  onPreview,
}: {
  form: CleanupScopeForm;
  onChange: (form: CleanupScopeForm) => void;
  onPreview: () => void;
}) {
  const [localSearch, setLocalSearch] = useState(form.searchQuery);

  function toggleWorkload(id: "teams" | "onedrive" | "sharepoint") {
    const next = new Set(form.workloads);
    next.has(id) ? next.delete(id) : next.add(id);
    onChange({ ...form, workloads: next });
  }

  return (
    <div className="mt-8">
      <h2 className="mb-1 text-lg font-semibold text-slate-800">2. Select & delete scope</h2>
      <p className="mb-4 text-sm text-slate-500">
        Search or bulk-select users, teams, or sites, then choose deletion granularity.
      </p>

      <input
        value={localSearch}
        onChange={(e) => {
          setLocalSearch(e.target.value);
          onChange({ ...form, searchQuery: e.target.value });
        }}
        placeholder="Search users, teams, or SharePoint sites…"
        className="mb-4 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none focus:ring-1 focus:ring-brand-500"
      />

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
        {WORKLOAD_OPTIONS.map((opt) => (
          <label
            key={opt.id}
            className={`flex cursor-pointer flex-col gap-1 rounded-lg border px-4 py-3 text-sm transition-colors ${
              form.workloads.has(opt.id) ? "border-brand-500 bg-brand-50" : "border-slate-200 bg-white"
            }`}
          >
            <span className="flex items-center gap-2 font-medium text-slate-800">
              <input
                type="checkbox"
                checked={form.workloads.has(opt.id)}
                onChange={() => toggleWorkload(opt.id)}
                className="accent-brand-500"
              />
              {opt.label}
            </span>
            <span className="text-xs text-slate-500">{opt.hint}</span>
          </label>
        ))}
      </div>

      {form.workloads.has("onedrive") && (
        <div className="mt-4 flex items-center gap-3">
          <label className="text-sm font-medium text-slate-700">Delete files older than</label>
          <input
            type="date"
            value={form.cutoffDate}
            onChange={(e) => onChange({ ...form, cutoffDate: e.target.value })}
            className="rounded-md border border-slate-300 px-3 py-1.5 text-sm focus:border-brand-500 focus:outline-none"
          />
          <span className="text-xs text-slate-500">migration cutoff date</span>
        </div>
      )}

      {form.workloads.has("teams") && (
        <label className="mt-4 flex items-center gap-2 text-sm text-slate-700">
          <input
            type="checkbox"
            checked={form.removeM365Groups}
            onChange={(e) => onChange({ ...form, removeM365Groups: e.target.checked })}
            className="accent-brand-500"
          />
          Also remove the M365 Group behind fully-cleaned Teams
        </label>
      )}

      <button
        onClick={onPreview}
        disabled={form.workloads.size === 0}
        className="mt-6 rounded-lg bg-brand-500 px-4 py-2 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-brand-600 disabled:cursor-not-allowed disabled:bg-slate-300"
      >
        Preview scope (dry-run) →
      </button>
    </div>
  );
}
