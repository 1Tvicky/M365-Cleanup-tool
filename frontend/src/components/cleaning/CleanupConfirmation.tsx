import { useEffect, useState } from "react";
import {
  startCleanup,
  validateCleanup,
  type CleanupDeletionMode,
  type CleanupManifest,
  type CleanupResourceType,
  type CleanupValidationResult,
} from "../../api/cleaning";
import { ApiClientError } from "../../api/client";

/**
 * Names the specific cloud(s) actually being cleaned instead of a vendor-specific default —
 * e.g. "OneDrive", "OneDrive and SharePoint", or "Google My Drive". Callers always follow this
 * with a trailing space before "data" (e.g. "the selected {name} data"), so the empty-names
 * fallback is "" rather than a filler phrase — that would otherwise read as "the selected the
 * selected data" once the caller's own "the selected" is added.
 *
 * `modeAwareOnly` excludes Microsoft Teams/Google Chat — neither respects the recycle-bin/
 * permanent choice at all (see MODE_AGNOSTIC comment below), so naming them in the text that's
 * specifically about that choice would misstate what actually happens to their data.
 */
function selectedCloudNames(summary: CleanupValidationResult["summary"], opts: { modeAwareOnly?: boolean } = {}): string {
  const names: string[] = [];
  if (summary.oneDriveAccounts > 0) names.push("OneDrive");
  if (summary.sharePointSites > 0) names.push("SharePoint");
  if (summary.outlookMailboxes > 0 || summary.outlookCalendars > 0 || summary.outlookContacts > 0) names.push("Outlook");
  if (!opts.modeAwareOnly && (summary.channels > 0 || summary.teams > 0)) names.push("Microsoft Teams");
  if (summary.googleMyDriveAccounts > 0) names.push("Google My Drive");
  if (summary.sharedDrives > 0) names.push("Google Shared Drives");
  if (summary.gmailMailboxes > 0) names.push("Gmail");
  if (!opts.modeAwareOnly && summary.googleChatSpaces > 0) names.push("Google Chat");
  return names.length > 0 ? `${names.join(" and ")} ` : "";
}

const RESOURCE_LABEL: Record<CleanupResourceType, string> = {
  onedrive_account: "OneDrive account",
  sharepoint_site: "SharePoint site",
  outlook_mailbox: "Outlook mailbox",
  outlook_calendar: "Outlook calendar event",
  outlook_contacts: "Outlook contact",
  channel: "Teams channel",
  chat: "Direct message conversation",
  team: "Microsoft Team",
  google_my_drive_account: "Google My Drive account",
  shared_drive: "Google Shared Drive",
  gmail_mailbox: "Gmail mailbox",
  google_chat_space: "Google Chat space",
};

/**
 * Sits between "Review your selection" and the background cleanup job. Re-validates the selection
 * itself (rather than trusting whatever the Review page computed) since time may have passed and
 * the backend re-checks everything again anyway on Start — this screen should never claim a
 * selection is fine if the backend would immediately reject it.
 */
export function CleanupConfirmation({
  manifest,
  onBack,
  onStarted,
}: {
  manifest: CleanupManifest;
  onBack: () => void;
  onStarted: (operationId: string) => void;
}) {
  const [result, setResult] = useState<CleanupValidationResult | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  // Defaults to the safer, recoverable option — an operator must explicitly opt into permanent
  // deletion, never the reverse (matches the backend's own default-if-missing behavior).
  const [deletionMode, setDeletionMode] = useState<CleanupDeletionMode>("recycle_bin");
  const [confirmed, setConfirmed] = useState(false);
  const [starting, setStarting] = useState(false);
  const [startError, setStartError] = useState<string | null>(null);

  useEffect(() => {
    validateCleanup(manifest)
      .then(setResult)
      .catch((err) => setLoadError(err instanceof ApiClientError ? err.message : "Couldn't check your selection. Try again."));
  }, [manifest]);

  // Switching modes mid-review resets the acknowledgement — a "yes I understand" ticked for one
  // mode must never silently carry over and authorize the other.
  function handleDeletionModeChange(mode: CleanupDeletionMode) {
    setDeletionMode(mode);
    setConfirmed(false);
  }

  async function handleStart() {
    setStarting(true);
    setStartError(null);
    try {
      const { operationId } = await startCleanup(manifest, deletionMode);
      onStarted(operationId);
    } catch (err) {
      setStartError(err instanceof ApiClientError ? err.message : "Couldn't start cleanup. Try again.");
    } finally {
      setStarting(false);
    }
  }

  const executableCount = result
    ? result.summary.oneDriveAccounts +
      result.summary.sharePointSites +
      result.summary.outlookMailboxes +
      result.summary.channels +
      result.summary.teams +
      result.summary.googleMyDriveAccounts +
      result.summary.sharedDrives +
      result.summary.gmailMailboxes +
      result.summary.googleChatSpaces
    : 0;

  // Teams (Team/Channel) and Google Chat Space deletion never consult deletion_mode at all — see
  // cleanupExecutionWorker.ts's executeAnyItem / googleChatCleanupExecution.ts. Neither Microsoft
  // nor Google offers a recoverable alternative for them, so the recycle-bin/permanent radio choice
  // below is meaningless when that's all that's selected — showing it (and a message promising a
  // "moved to recycle bin" outcome) would flatly contradict the unconditional warning above it.
  const modeAgnosticCount = result ? result.summary.teams + result.summary.channels + result.summary.googleChatSpaces : 0;
  const hasModeChoice = executableCount - modeAgnosticCount > 0;

  return (
    <div className="mx-auto max-w-2xl px-8 py-10">
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Ready to clean up</h2>
      <p className="mb-6 text-sm text-slate-500">
        {!result || hasModeChoice
          ? `Choose how the selected ${result ? selectedCloudNames(result.summary, { modeAwareOnly: true }) : ""}data should be removed.`
          : `Review the selected ${selectedCloudNames(result.summary)}data below before starting cleanup.`}
      </p>

      {loadError ? (
        <p className="rounded-lg border border-dashed border-rose-300 px-4 py-4 text-center text-sm text-rose-600">{loadError}</p>
      ) : !result ? (
        <p className="rounded-lg border border-dashed border-slate-300 px-4 py-8 text-center text-sm text-slate-500">Checking your selection…</p>
      ) : !result.valid ? (
        <div className="rounded-lg border border-dashed border-rose-300 px-4 py-4 text-sm text-rose-600">
          <p className="mb-2 font-medium">Your selection needs to be reviewed again.</p>
          <ul className="list-inside list-disc space-y-1">
            {result.errors.map((e, i) => (
              <li key={i}>{e}</li>
            ))}
          </ul>
        </div>
      ) : (
        <>
          <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-6">
            {result.summary.oneDriveAccounts > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">OneDrive Accounts</span>
                <span className="text-sm text-slate-600">{result.summary.oneDriveAccounts.toLocaleString()}</span>
              </div>
            )}
            {result.summary.sharePointSites > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">SharePoint Sites</span>
                <span className="text-sm text-slate-600">{result.summary.sharePointSites.toLocaleString()}</span>
              </div>
            )}
            {result.summary.outlookMailboxes > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Outlook Mailboxes</span>
                <span className="text-sm text-slate-600">{result.summary.outlookMailboxes.toLocaleString()}</span>
              </div>
            )}
            {result.summary.teams > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Microsoft Teams (whole team, all channels)</span>
                <span className="text-sm text-slate-600">{result.summary.teams.toLocaleString()}</span>
              </div>
            )}
            {result.summary.channels > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Teams Channels</span>
                <span className="text-sm text-slate-600">{result.summary.channels.toLocaleString()}</span>
              </div>
            )}
            {result.summary.googleMyDriveAccounts > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Google My Drive Accounts</span>
                <span className="text-sm text-slate-600">{result.summary.googleMyDriveAccounts.toLocaleString()}</span>
              </div>
            )}
            {result.summary.sharedDrives > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Google Shared Drives</span>
                <span className="text-sm text-slate-600">{result.summary.sharedDrives.toLocaleString()}</span>
              </div>
            )}
            {result.summary.gmailMailboxes > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Gmail Mailboxes</span>
                <span className="text-sm text-slate-600">{result.summary.gmailMailboxes.toLocaleString()}</span>
              </div>
            )}
            {result.summary.googleChatSpaces > 0 && (
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-sm font-medium text-slate-700">Google Chat Spaces</span>
                <span className="text-sm text-slate-600">{result.summary.googleChatSpaces.toLocaleString()}</span>
              </div>
            )}
            {executableCount === 0 && (
              <p className="text-sm text-slate-500">Nothing in your selection can be cleaned up automatically yet — see below.</p>
            )}
          </div>

          {result.unsupported.length > 0 && (
            <div className="mt-4 rounded-lg border border-amber-200 bg-amber-50 px-4 py-4 text-sm text-amber-800">
              <p className="mb-1 font-medium">Not supported yet</p>
              <p className="mb-2">
                Microsoft doesn't currently allow this app to remove Teams direct message content automatically. These{" "}
                {result.unsupported.length.toLocaleString()} item{result.unsupported.length === 1 ? "" : "s"} will be skipped:
              </p>
              <ul className="list-inside list-disc space-y-0.5">
                {result.unsupported.slice(0, 8).map((u, i) => (
                  <li key={i}>
                    {RESOURCE_LABEL[u.resourceType]}: {u.displayName}
                  </li>
                ))}
                {result.unsupported.length > 8 && <li>…and {(result.unsupported.length - 8).toLocaleString()} more</li>}
              </ul>
            </div>
          )}

          {(result.summary.teams > 0 || result.summary.channels > 0 || result.summary.googleChatSpaces > 0) && (
            <div className="mt-4 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
              <p className="mb-1 font-medium">Teams and Google Chat deletions are immediate and permanent</p>
              <p>
                {hasModeChoice
                  ? "These do not follow the recycle-bin/permanent choice below — Microsoft and Google don't offer a recoverable option for them:"
                  : "Microsoft and Google don't offer a recoverable option for these — there is no recycle bin to choose:"}
              </p>
              <ul className="mt-1 list-inside list-disc space-y-0.5">
                {result.summary.teams > 0 && (
                  <li>
                    <span className="font-medium">{result.summary.teams.toLocaleString()} whole Team{result.summary.teams === 1 ? "" : "s"}</span> —
                    deletes the Team itself and every channel in it, not just its messages.
                  </li>
                )}
                {result.summary.channels > 0 && (
                  <li>
                    <span className="font-medium">{result.summary.channels.toLocaleString()} Teams channel{result.summary.channels === 1 ? "" : "s"}</span> —
                    deletes only the selected channel(s); the parent Team and its other channels are unaffected.
                  </li>
                )}
                {result.summary.googleChatSpaces > 0 && (
                  <li>
                    <span className="font-medium">{result.summary.googleChatSpaces.toLocaleString()} Google Chat space{result.summary.googleChatSpaces === 1 ? "" : "s"}</span> —
                    deletes the entire Space (all messages and memberships in it), not just its messages.
                  </li>
                )}
              </ul>
            </div>
          )}

          {executableCount > 0 && (
            <>
              {hasModeChoice && (
                <>
                  <fieldset className="mt-6 rounded-xl border border-slate-200 bg-white p-4">
                    <legend className="px-1 text-sm font-semibold text-slate-700">How should this data be removed?</legend>
                    <label className="mt-2 flex items-start gap-2 text-sm text-slate-700">
                      <input
                        type="radio"
                        name="deletionMode"
                        className="mt-0.5"
                        checked={deletionMode === "recycle_bin"}
                        onChange={() => handleDeletionModeChange("recycle_bin")}
                      />
                      <span>
                        <span className="font-medium">Move to recycle bin</span> — recoverable from the recycle bin / Deleted Items for a limited time.
                      </span>
                    </label>
                    <label className="mt-3 flex items-start gap-2 text-sm text-slate-700">
                      <input
                        type="radio"
                        name="deletionMode"
                        className="mt-0.5"
                        checked={deletionMode === "permanent"}
                        onChange={() => handleDeletionModeChange("permanent")}
                      />
                      <span>
                        <span className="font-medium">Permanently delete</span> — removed from the recycle bin / Deleted Items immediately. This cannot
                        be undone.
                      </span>
                    </label>
                  </fieldset>

                  <p className="mt-4 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
                    {deletionMode === "permanent"
                      ? `The selected ${selectedCloudNames(result.summary, { modeAwareOnly: true })}data will be permanently deleted — removed from the recycle bin / Deleted Items immediately, not just moved there. This covers everything currently in the selected resources when the job runs, not only what's listed above. This action cannot be undone.`
                      : `The selected ${selectedCloudNames(result.summary, { modeAwareOnly: true })}data will be moved to the recycle bin / Deleted Items. This covers everything currently in the selected resources when the job runs, not only what's listed above.`}{" "}
                    Make sure you have reviewed your selection before continuing.
                  </p>
                </>
              )}

              <label className="mt-4 flex items-start gap-2 text-sm text-slate-700">
                <input type="checkbox" className="mt-0.5" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} />
                {!hasModeChoice || deletionMode === "permanent"
                  ? "I understand that this cleanup will permanently remove the selected data and cannot be undone."
                  : "I understand that this cleanup will remove the selected data."}
              </label>

              {startError && <p className="mt-3 text-sm text-rose-600">{startError}</p>}
            </>
          )}
        </>
      )}

      <div className="mt-6 flex items-center gap-3">
        <button onClick={onBack} className="rounded-md border border-slate-200 px-5 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
          ← Back
        </button>
        {result?.valid && executableCount > 0 && (
          <button
            onClick={handleStart}
            disabled={!confirmed || starting}
            className="rounded-md bg-[#1b2fc4] px-5 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {starting ? "Starting…" : "Start Cleanup"}
          </button>
        )}
      </div>
    </div>
  );
}
