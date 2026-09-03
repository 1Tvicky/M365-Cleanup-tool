import { CONNECTORS } from "../../api/mockData";
import type { Workload } from "../../types";
import { OneDriveIcon, SharePointIcon, TeamsIcon } from "./CloudIcons";

const ICONS: Record<Workload, (props: { className?: string }) => JSX.Element> = {
  onedrive: OneDriveIcon,
  sharepoint: SharePointIcon,
  teams: TeamsIcon,
};

// Per-connector "Add Cloud" button color and label tint, matching each product's brand color.
const ACCENTS: Record<Workload, { button: string; label: string }> = {
  onedrive: { button: "#0078D4", label: "text-slate-800" },
  sharepoint: { button: "#038387", label: "text-slate-800" },
  teams: { button: "#5B5FC7", label: "text-[#4550A8]" },
};

/**
 * Tile grid for connecting an M365 workload — mirrors CloudFuze's "Business Clouds" tile layout,
 * scoped to Phase 1's three M365 workloads only. Clicking "Add Cloud" launches Microsoft admin
 * OAuth (docs/azure-ad-app-registration.md §4); this component only fires onConnect, the redirect
 * to Microsoft's consent screen happens server-side via GET /api/v1/auth/consent-url.
 */
export function CloudTileGrid({ onConnect }: { onConnect: (workload: Workload) => void }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white px-7 py-7">
      <h2 className="mb-5 text-base font-semibold text-slate-800">Business Clouds</h2>
      <div className="flex flex-wrap gap-4">
        {CONNECTORS.map((c) => {
          const Icon = ICONS[c.id];
          const accent = ACCENTS[c.id];
          return (
            <div
              key={c.id}
              className="flex w-[150px] flex-col items-center gap-3 rounded-lg border border-slate-200 bg-white px-3 py-4 transition-colors hover:border-teal-400 hover:shadow-sm"
            >
              <Icon className="h-12 w-12" />
              <span className={`text-center text-sm font-semibold leading-tight ${accent.label}`}>{c.label}</span>
              <button
                onClick={() => onConnect(c.id)}
                className="mt-auto w-full rounded-md py-1.5 text-xs font-semibold text-white transition-opacity hover:opacity-90"
                style={{ backgroundColor: accent.button }}
              >
                Add Cloud
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}
