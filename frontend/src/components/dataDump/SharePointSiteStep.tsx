import { useState } from "react";
import type { AvailableResourceRow } from "../../api/clouds";
import type { NewSiteConfig } from "../../api/dataDump";
import { ResourceSelectionStep, connectionResourceFetcher } from "./ResourceSelectionStep";

/**
 * SharePoint's Select Resources step supports both target shapes spec §6 asks for — existing sites
 * and a brand-new site — as two independent, simultaneously-usable sections (an operation may
 * target existing sites AND create one new site together), rather than a forced either/or radio.
 */
export function SharePointSiteStep({
  connectionId,
  selectedSites,
  onToggleSite,
  onToggleAllSites,
  newSite,
  onNewSiteChange,
}: {
  connectionId: string;
  selectedSites: Map<string, AvailableResourceRow>;
  onToggleSite: (id: string, row: AvailableResourceRow) => void;
  onToggleAllSites: (rows: AvailableResourceRow[]) => void;
  newSite: NewSiteConfig | null;
  onNewSiteChange: (site: NewSiteConfig | null) => void;
}) {
  const [showNewSiteForm, setShowNewSiteForm] = useState(!!newSite);

  return (
    <div className="space-y-6">
      <ResourceSelectionStep
        fetcher={connectionResourceFetcher(connectionId)}
        title="SharePoint"
        subtitle="Select existing sites to generate data into."
        searchPlaceholder="Search sites…"
        secondaryLabel="URL"
        selected={selectedSites}
        onToggle={onToggleSite}
        onToggleAll={onToggleAllSites}
        emptyMessage="No SharePoint sites found."
      />

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <label className="flex cursor-pointer items-center gap-2 text-sm font-medium text-slate-700">
          <input
            type="checkbox"
            checked={showNewSiteForm}
            onChange={(e) => {
              setShowNewSiteForm(e.target.checked);
              if (!e.target.checked) onNewSiteChange(null);
            }}
          />
          Create New SharePoint Site
        </label>
        {showNewSiteForm && (
          <div className="mt-4 space-y-3">
            <Field label="Site Name">
              <input
                value={newSite?.displayName ?? ""}
                onChange={(e) => onNewSiteChange({ ...emptySite(newSite), displayName: e.target.value })}
                placeholder="Migration Demo Site"
                className="w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-[#1b2fc4] focus:outline-none"
              />
            </Field>
            <Field label="Site URL">
              <div className="flex items-center gap-1 text-sm text-slate-500">
                <span>/sites/</span>
                <input
                  value={newSite?.urlSlug ?? ""}
                  onChange={(e) => onNewSiteChange({ ...emptySite(newSite), urlSlug: e.target.value.replace(/[^a-zA-Z0-9-]/g, "-") })}
                  placeholder="migration-demo-site"
                  className="flex-1 rounded-md border border-slate-200 px-3 py-2 text-sm text-slate-800 focus:border-[#1b2fc4] focus:outline-none"
                />
              </div>
            </Field>
            <Field label="Site Type">
              <div className="space-y-1.5">
                <label className="flex items-center gap-2 text-sm text-slate-700">
                  <input
                    type="radio"
                    checked={(newSite?.template ?? "teamSiteWithoutMicrosoft365Group") === "teamSiteWithoutMicrosoft365Group"}
                    onChange={() => onNewSiteChange({ ...emptySite(newSite), template: "teamSiteWithoutMicrosoft365Group" })}
                  />
                  Team Site
                </label>
                <label className="flex items-center gap-2 text-sm text-slate-700">
                  <input
                    type="radio"
                    checked={newSite?.template === "communicationSite"}
                    onChange={() => onNewSiteChange({ ...emptySite(newSite), template: "communicationSite" })}
                  />
                  Communication Site
                </label>
              </div>
            </Field>
            <Field label="Description">
              <input
                value={newSite?.description ?? ""}
                onChange={(e) => onNewSiteChange({ ...emptySite(newSite), description: e.target.value })}
                placeholder="Synthetic migration test data"
                className="w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-[#1b2fc4] focus:outline-none"
              />
            </Field>
            <p className="text-xs text-amber-700">
              ⚠ Creating a new site requires the Sites.Create.All Graph permission. If this app's Azure AD registration hasn't been granted it, site creation will fail with a clear
              error rather than silently succeeding.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

function emptySite(current: NewSiteConfig | null): NewSiteConfig {
  return current ?? { displayName: "", urlSlug: "", template: "teamSiteWithoutMicrosoft365Group" };
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block text-xs font-medium text-slate-500">
      {label}
      <div className="mt-1">{children}</div>
    </label>
  );
}
