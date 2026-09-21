/**
 * Data Dump's own version of components/cleaning/SelectionSummary.tsx — same fixed-bottom-bar visual
 * pattern (spec: "must look and behave like Cleanup"), but with Data Dump's own totals shape, never
 * importing SelectionTotals from Cleaning (keeps the two business functions' selection state fully
 * separate, per the standing "keep Data Dump business logic separate" requirement).
 */
export interface DataDumpSelectionTotals {
  onedriveUsers: number;
  sharepointSites: number;
  sharepointNewSite: boolean;
  teamsTeams: number;
  teamsNewTeam: boolean;
  outlookUsers: number;
}

export function hasDataDumpSelection(t: DataDumpSelectionTotals): boolean {
  return t.onedriveUsers + t.sharepointSites + (t.sharepointNewSite ? 1 : 0) + t.teamsTeams + (t.teamsNewTeam ? 1 : 0) + t.outlookUsers > 0;
}

export function DataDumpSelectionBar({ totals, onReview, reviewLabel = "Configure Data Dump" }: { totals: DataDumpSelectionTotals; onReview: () => void; reviewLabel?: string }) {
  if (!hasDataDumpSelection(totals)) return null;

  const parts: string[] = [];
  if (totals.onedriveUsers > 0) parts.push(`${totals.onedriveUsers.toLocaleString()} OneDrive user${totals.onedriveUsers === 1 ? "" : "s"}`);
  if (totals.sharepointSites > 0) parts.push(`${totals.sharepointSites.toLocaleString()} SharePoint site${totals.sharepointSites === 1 ? "" : "s"}`);
  if (totals.sharepointNewSite) parts.push("1 new SharePoint site");
  if (totals.teamsTeams > 0) parts.push(`${totals.teamsTeams.toLocaleString()} Team${totals.teamsTeams === 1 ? "" : "s"}`);
  if (totals.teamsNewTeam) parts.push("1 new Team");
  if (totals.outlookUsers > 0) parts.push(`${totals.outlookUsers.toLocaleString()} Outlook user${totals.outlookUsers === 1 ? "" : "s"}`);

  return (
    <div className="fixed bottom-0 left-20 right-0 z-20 border-t border-slate-200 bg-white px-8 py-4 shadow-[0_-4px_12px_rgba(0,0,0,0.08)]">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Selected</div>
          <div className="mt-0.5 text-sm text-slate-700">{parts.join(" · ")}</div>
        </div>
        <button onClick={onReview} className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90">
          {reviewLabel}
        </button>
      </div>
    </div>
  );
}
