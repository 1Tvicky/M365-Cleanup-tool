import { CONNECTORS } from "../../api/mockData";
import type { Workload } from "../../types";
import { GmailIcon, GoogleChatIcon, GoogleIcon, OneDriveIcon, OutlookIcon, SharedDrivesIcon, SharePointIcon, TeamsIcon } from "./CloudIcons";

const ICONS: Record<Workload, (props: { className?: string }) => JSX.Element> = {
  onedrive: OneDriveIcon,
  sharepoint: SharePointIcon,
  teams: TeamsIcon,
  outlook: OutlookIcon,
  google_my_drive: GoogleIcon,
  shared_drive: SharedDrivesIcon,
  google_chat: GoogleChatIcon,
  gmail: GmailIcon,
};

// Per-connector "Add Cloud" button color and label tint, matching each product's brand color.
const ACCENTS: Record<Workload, { button: string; label: string }> = {
  onedrive: { button: "#0078D4", label: "text-slate-800" },
  sharepoint: { button: "#038387", label: "text-slate-800" },
  teams: { button: "#5B5FC7", label: "text-[#4550A8]" },
  outlook: { button: "#0078D4", label: "text-slate-800" },
  google_my_drive: { button: "#1a73e8", label: "text-slate-800" },
  shared_drive: { button: "#0f9d58", label: "text-slate-800" },
  google_chat: { button: "#00897b", label: "text-slate-800" },
  gmail: { button: "#ea4335", label: "text-slate-800" },
};

/** Every Google-provider workload — each connects via its own OAuth popup, independently of the others (see CloudsPage.tsx's handleConnectGoogle). */
const GOOGLE_WORKLOADS = new Set<Workload>(["google_my_drive", "shared_drive", "google_chat", "gmail"]);

/**
 * Tile grid for connecting a workload — mirrors CloudFuze's "Business Clouds" tile layout.
 * "Add Cloud" only appears on hover (matches the reference product exactly — verified against a
 * screen recording, not just the earlier cropped screenshots which happened to catch a hover
 * state). Every tile launches its own OAuth popup (Microsoft's or Google's, per GOOGLE_WORKLOADS) —
 * this component only calls onConnect/onConnectGoogle, CloudsPage owns the actual popup lifecycle.
 */
export function CloudTileGrid({
  onConnect,
  onConnectGoogle,
  connectingTypes,
}: {
  onConnect: (workload: Workload) => void;
  onConnectGoogle: (workload: Workload) => void;
  connectingTypes: Set<Workload>;
}) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white px-7 py-7">
      <h2 className="mb-5 text-base font-semibold text-slate-800">Business Clouds</h2>
      <div className="flex flex-wrap gap-4">
        {CONNECTORS.map((c) => {
          const Icon = ICONS[c.id];
          const accent = ACCENTS[c.id];
          const isConnecting = connectingTypes.has(c.id);
          return (
            <div
              key={c.id}
              className="group relative flex w-[150px] flex-col items-center gap-3 rounded-lg border border-slate-200 bg-white px-3 py-4 transition-colors hover:border-teal-400 hover:shadow-sm"
            >
              <Icon className="h-12 w-12" />
              <span className={`text-center text-sm font-semibold leading-tight ${accent.label}`}>{c.label}</span>
              <button
                onClick={() => (GOOGLE_WORKLOADS.has(c.id) ? onConnectGoogle(c.id) : onConnect(c.id))}
                disabled={isConnecting}
                className={`mt-auto w-full rounded-md py-1.5 text-xs font-semibold text-white transition-opacity disabled:cursor-not-allowed ${
                  isConnecting ? "opacity-100" : "opacity-0 group-hover:opacity-100 group-hover:enabled:hover:opacity-90"
                }`}
                style={{ backgroundColor: accent.button }}
              >
                {isConnecting ? "Connecting…" : "Add Cloud"}
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}
