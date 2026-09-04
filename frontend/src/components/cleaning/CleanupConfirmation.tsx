import { useEffect, useState } from "react";
import { startCleanup, validateCleanup, type CleanupManifest, type CleanupResourceType, type CleanupValidationResult } from "../../api/cleaning";
import { ApiClientError } from "../../api/client";

/** Names the specific cloud(s) actually being cleaned instead of the generic "Microsoft 365 data" — e.g. "OneDrive", or "OneDrive and SharePoint" when both are in the selection. */
function selectedCloudNames(summary: CleanupValidationResult["summary"]): string {
  const names: string[] = [];
  if (summary.oneDriveAccounts > 0) names.push("OneDrive");
  if (summary.sharePointSites > 0) names.push("SharePoint");
  return names.length > 0 ? names.join(" and ") : "Microsoft 365";
}

const RESOURCE_LABEL: Record<CleanupResourceType, string> = {
  onedrive_account: "OneDrive account",
  sharepoint_site: "SharePoint site",
  channel: "Teams channel",
  chat: "Direct message conversation",
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
  const [confirmed, setConfirmed] = useState(false);
  const [starting, setStarting] = useState(false);
  const [startError, setStartError] = useState<string | null>(null);

  useEffect(() => {
    validateCleanup(manifest)
      .then(setResult)
      .catch((err) => setLoadError(err instanceof ApiClientError ? err.message : "Couldn't check your selection. Try again."));
  }, [manifest]);

  async function handleStart() {
    setStarting(true);
    setStartError(null);
    try {
      const { operationId } = await startCleanup(manifest);
      onStarted(operationId);
    } catch (err) {
      setStartError(err instanceof ApiClientError ? err.message : "Couldn't start cleanup. Try again.");
    } finally {
      setStarting(false);
    }
  }

  const executableCount = result ? result.summary.oneDriveAccounts + result.summary.sharePointSites : 0;

  return (
    <div className="mx-auto max-w-2xl px-8 py-10">
      <h2 className="mb-1 text-lg font-semibold text-slate-800">Ready to clean up</h2>
      <p className="mb-6 text-sm text-slate-500">This will permanently remove the selected {result ? selectedCloudNames(result.summary) : "Microsoft 365"} data.</p>

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
            {executableCount === 0 && (
              <p className="text-sm text-slate-500">Nothing in your selection can be cleaned up automatically yet — see below.</p>
            )}
          </div>

          {result.unsupported.length > 0 && (
            <div className="mt-4 rounded-lg border border-amber-200 bg-amber-50 px-4 py-4 text-sm text-amber-800">
              <p className="mb-1 font-medium">Not supported yet</p>
              <p className="mb-2">
                Microsoft doesn't currently allow this app to remove Teams channel or direct message content automatically. These{" "}
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

          {executableCount > 0 && (
            <>
              <p className="mt-6 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
                This action will remove the selected {selectedCloudNames(result.summary)} data. Make sure you have reviewed your selection before continuing.
              </p>

              <label className="mt-4 flex items-start gap-2 text-sm text-slate-700">
                <input type="checkbox" className="mt-0.5" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} />
                I understand that this cleanup will remove the selected data.
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
