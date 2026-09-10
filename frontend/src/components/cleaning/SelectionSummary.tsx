import { formatBytes } from "../../utils/format";

export interface SelectionTotals {
  oneDriveAccounts: number;
  oneDriveBytes: number;
  sharePointSites: number;
  sharePointBytes: number;
  outlookMailboxes: number;
  /** Mail item count, not bytes — no cheap per-mailbox byte quota exists under application permissions (see backend/src/graph/cloudEnumeration.ts's getUserMailSummary). */
  outlookItems: number;
  /** How many mailboxes are selected for Calendar cleanup — a separate selection from outlookMailboxes, even though both pick from the same user list (see CleaningPage.tsx's OutlookView tabs). */
  outlookCalendarUsers: number;
  outlookCalendarEvents: number;
  /** Same relationship to outlookMailboxes as outlookCalendarUsers above, for Contacts. */
  outlookContactUsers: number;
  outlookContactCount: number;
  teamsChannels: number;
  teamsMessages: number;
  /** How many of the selected channels have an actual computed count — lets the summary tell "0 messages" (genuinely counted) apart from "count not known yet" for the same selection. */
  teamsChannelsWithKnownCount: number;
  dms: number;
  dmMessages: number;
  dmsWithKnownCount: number;
  googleMyDriveAccounts: number;
  googleMyDriveBytes: number;
  sharedDrives: number;
  sharedDrivesBytes: number;
  gmailMailboxes: number;
  /** Message count, not bytes — Gmail exposes no per-mailbox storage-bytes field at all (see backend/src/graph/gmailEnumeration.ts). */
  gmailItems: number;
  googleChatSpaces: number;
}

export function messagesFragment(selectedCount: number, knownCount: number, total: number): string {
  if (knownCount === 0) return "message count unavailable";
  if (knownCount < selectedCount) return `${total.toLocaleString()} messages known (${selectedCount - knownCount} unavailable)`;
  return `${total.toLocaleString()} messages`;
}

export function hasSelection(t: SelectionTotals): boolean {
  return (
    t.oneDriveAccounts +
      t.sharePointSites +
      t.outlookMailboxes +
      t.outlookCalendarUsers +
      t.outlookContactUsers +
      t.teamsChannels +
      t.dms +
      t.googleMyDriveAccounts +
      t.sharedDrives +
      t.gmailMailboxes +
      t.googleChatSpaces >
    0
  );
}

/** "3 accounts selected · 450 GB" style summary, sticky at the bottom while a selection exists. */
export function SelectionSummary({ totals, onReview }: { totals: SelectionTotals; onReview: () => void }) {
  if (!hasSelection(totals)) return null;

  const parts: string[] = [];
  if (totals.oneDriveAccounts > 0) parts.push(`${totals.oneDriveAccounts.toLocaleString()} OneDrive account${totals.oneDriveAccounts === 1 ? "" : "s"} · ${formatBytes(totals.oneDriveBytes)}`);
  if (totals.sharePointSites > 0) parts.push(`${totals.sharePointSites.toLocaleString()} SharePoint site${totals.sharePointSites === 1 ? "" : "s"} · ${formatBytes(totals.sharePointBytes)}`);
  if (totals.outlookMailboxes > 0) parts.push(`${totals.outlookMailboxes.toLocaleString()} Outlook mailbox${totals.outlookMailboxes === 1 ? "" : "es"} · ${totals.outlookItems.toLocaleString()} mail items`);
  if (totals.outlookCalendarUsers > 0)
    parts.push(`${totals.outlookCalendarUsers.toLocaleString()} Outlook calendar${totals.outlookCalendarUsers === 1 ? "" : "s"} · ${totals.outlookCalendarEvents.toLocaleString()} events`);
  if (totals.outlookContactUsers > 0)
    parts.push(`${totals.outlookContactUsers.toLocaleString()} Outlook contacts mailbox${totals.outlookContactUsers === 1 ? "" : "es"} · ${totals.outlookContactCount.toLocaleString()} contacts`);
  if (totals.teamsChannels > 0)
    parts.push(
      `${totals.teamsChannels.toLocaleString()} Teams channel${totals.teamsChannels === 1 ? "" : "s"} · ${messagesFragment(totals.teamsChannels, totals.teamsChannelsWithKnownCount, totals.teamsMessages)}`
    );
  if (totals.dms > 0)
    parts.push(`${totals.dms.toLocaleString()} conversation${totals.dms === 1 ? "" : "s"} · ${messagesFragment(totals.dms, totals.dmsWithKnownCount, totals.dmMessages)}`);
  if (totals.googleMyDriveAccounts > 0)
    parts.push(`${totals.googleMyDriveAccounts.toLocaleString()} Google My Drive account${totals.googleMyDriveAccounts === 1 ? "" : "s"} · ${formatBytes(totals.googleMyDriveBytes)}`);
  if (totals.sharedDrives > 0)
    parts.push(`${totals.sharedDrives.toLocaleString()} Shared Drive${totals.sharedDrives === 1 ? "" : "s"} · ${formatBytes(totals.sharedDrivesBytes)}`);
  if (totals.gmailMailboxes > 0)
    parts.push(`${totals.gmailMailboxes.toLocaleString()} Gmail mailbox${totals.gmailMailboxes === 1 ? "" : "es"} · ${totals.gmailItems.toLocaleString()} messages`);
  if (totals.googleChatSpaces > 0)
    parts.push(`${totals.googleChatSpaces.toLocaleString()} Chat space${totals.googleChatSpaces === 1 ? "" : "s"}`);

  return (
    // Fixed to the viewport, not sticky within the page's own scroll flow — a long discovery list
    // (found via testing: 50 rows per page) meant `sticky` only caught up once you'd scrolled all
    // the way to the bottom of the content, which defeats the point of an always-visible summary.
    // left-20 clears the sidebar's fixed w-20 width (see components/layout/SideNav.tsx).
    <div className="fixed bottom-0 left-20 right-0 z-20 border-t border-slate-200 bg-white px-8 py-4 shadow-[0_-4px_12px_rgba(0,0,0,0.08)]">
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
