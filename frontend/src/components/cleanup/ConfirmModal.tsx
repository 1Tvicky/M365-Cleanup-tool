import { useState } from "react";
import { formatBytes } from "../../utils/format";
import type { Preview } from "../../types";

/**
 * Mandatory safety flow per docs/rollback-safety.md: typed "DELETE" confirmation, plus a
 * pre-delete export that is always generated (metadata manifest always; content zip optional) —
 * this modal cannot be dismissed into an execute call without both.
 */
export function ConfirmModal({
  preview,
  tenantName,
  onClose,
  onExecute,
}: {
  preview: Preview;
  tenantName: string;
  onClose: () => void;
  onExecute: (exportManifestOnly: boolean) => void;
}) {
  const [typed, setTyped] = useState("");
  const [manifestOnly, setManifestOnly] = useState(false);
  const canExecute = typed === "DELETE";

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 px-4">
      <div className="w-full max-w-lg rounded-xl bg-white p-6 shadow-xl">
        <h3 className="text-lg font-semibold text-slate-900">Confirm permanent deletion</h3>
        <p className="mt-1 text-sm text-slate-500">
          You are about to delete data from <span className="font-medium text-slate-700">{tenantName}</span>.
        </p>

        <div className="mt-4 rounded-lg bg-slate-50 px-4 py-3 text-sm">
          <div className="flex justify-between py-0.5">
            <span className="text-slate-500">Items to delete</span>
            <span className="font-medium text-slate-800">{preview.totals.itemCount.toLocaleString()}</span>
          </div>
          <div className="flex justify-between py-0.5">
            <span className="text-slate-500">Total size</span>
            <span className="font-medium text-slate-800">{formatBytes(preview.totals.totalSizeBytes)}</span>
          </div>
        </div>

        <div className="mt-4 rounded-lg border border-slate-200 px-4 py-3">
          <label className="flex items-start gap-2 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={manifestOnly}
              onChange={(e) => setManifestOnly(e.target.checked)}
              className="mt-0.5 accent-brand-500"
            />
            <span>
              Export <span className="font-medium">manifest only</span> (skip zipping file contents to blob
              storage). A manifest CSV is always generated before deletion regardless of this choice.
            </span>
          </label>
        </div>

        <p className="mt-4 text-xs text-slate-500">
          Graph API deletes are largely irreversible outside their native retention window — see the{" "}
          <span className="underline decoration-dotted">rollback &amp; safety doc</span>. Private Teams channels
          cannot be restored at all.
        </p>

        <label className="mt-4 block text-sm font-medium text-slate-700">
          Type <span className="font-mono font-semibold text-rose-600">DELETE</span> to confirm
        </label>
        <input
          value={typed}
          onChange={(e) => setTyped(e.target.value)}
          className="mt-1 w-full rounded-md border border-slate-300 px-3 py-2 text-sm focus:border-rose-500 focus:outline-none focus:ring-1 focus:ring-rose-500"
          placeholder="DELETE"
        />

        <div className="mt-6 flex justify-end gap-2">
          <button
            onClick={onClose}
            className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
          >
            Cancel
          </button>
          <button
            onClick={() => onExecute(manifestOnly)}
            disabled={!canExecute}
            className="rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
          >
            Export backup & delete
          </button>
        </div>
      </div>
    </div>
  );
}
