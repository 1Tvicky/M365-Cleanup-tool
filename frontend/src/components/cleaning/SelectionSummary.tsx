import { formatBytes } from "../../utils/format";

export interface SelectionTotals {
  oneDriveAccounts: number;
  oneDriveBytes: number;
  sharePointSites: number;
  sharePointBytes: number;
  teamsChannels: number;
  teamsMessages: number;
  /** How many of the selected channels have an actual computed count — lets the summary tell "0 messages" (genuinely counted) apart from "count not known yet" for the same selection. */
  teamsChannelsWithKnownCount: number;
  dms: number;
  dmMessages: number;
  dmsWithKnownCount: number;
}

export function messagesFragment(selectedCount: number, knownCount: number, total: number): string {
  if (knownCount === 0) return "message count unavailable";
  if (knownCount < selectedCount) return `${total.toLocaleString()} messages known (${selectedCount - knownCount} unavailable)`;
  return `${total.toLocaleString()} messages`;
}

export function hasSelection(t: SelectionTotals): boolean {
  return t.oneDriveAccounts + t.sharePointSites + t.teamsChannels + t.dms > 0;
}

/** "3 accounts selected · 450 GB" style summary, sticky at the bottom while a selection exists. */
export function SelectionSummary({ totals, onReview }: { totals: SelectionTotals; onReview: () => void }) {
  if (!hasSelection(totals)) return null;

  const parts: string[] = [];
  if (totals.oneDriveAccounts > 0) parts.push(`${totals.oneDriveAccounts.toLocaleString()} OneDrive account${totals.oneDriveAccounts === 1 ? "" : "s"} · ${formatBytes(totals.oneDriveBytes)}`);
  if (totals.sharePointSites > 0) parts.push(`${totals.sharePointSites.toLocaleString()} SharePoint site${totals.sharePointSites === 1 ? "" : "s"} · ${formatBytes(totals.sharePointBytes)}`);
  if (totals.teamsChannels > 0)
    parts.push(
      `${totals.teamsChannels.toLocaleString()} Teams channel${totals.teamsChannels === 1 ? "" : "s"} · ${messagesFragment(totals.teamsChannels, totals.teamsChannelsWithKnownCount, totals.teamsMessages)}`
    );
  if (totals.dms > 0)
    parts.push(`${totals.dms.toLocaleString()} conversation${totals.dms === 1 ? "" : "s"} · ${messagesFragment(totals.dms, totals.dmsWithKnownCount, totals.dmMessages)}`);

  return (
    <div className="sticky bottom-0 left-0 right-0 z-10 border-t border-slate-200 bg-white px-8 py-4 shadow-[0_-4px_12px_rgba(0,0,0,0.04)]">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Selected</div>
          <div className="mt-0.5 text-sm text-slate-700">{parts.join(" · ")}</div>
        </div>
        <button
          onClick={onReview}
          className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90"
        >
          Review Selection
        </button>
      </div>
    </div>
  );
}
