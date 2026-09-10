import { useEffect, useState } from "react";
import { getCleanupOperationItemFiles, type CleanupDeletionMode, type CleanupItemFileRow, type CleanupResourceType } from "../../api/cleaning";
import { ApiClientError } from "../../api/client";
import { formatBytes, formatDate } from "../../utils/format";
import { PageFooter } from "./DiscoveryTable";

const PAGE_SIZE = 20;

// Only OneDrive/SharePoint/Outlook-mail items perform real permanent deletion
// (graph/cleanupDeletion.ts) — Calendar/Contacts stay on plain soft delete regardless of the
// operation's deletion_mode, so a 'deleted' file row for those must keep reading as "Removed."
const PERMANENT_DELETE_RESOURCE_TYPES = new Set<CleanupResourceType>(["onedrive_account", "sharepoint_site", "outlook_mailbox", "google_my_drive_account", "shared_drive"]);

const BASE_STATUS_STYLE: Record<CleanupItemFileRow["status"], { label: string; className: string }> = {
  pending: { label: "Pending", className: "text-slate-400" },
  deleted: { label: "Removed", className: "text-emerald-600" },
  already_gone: { label: "Already gone", className: "text-slate-500" },
  failed: { label: "Failed", className: "text-rose-500" },
};

function fileStatusLabel(status: CleanupItemFileRow["status"], resourceType: CleanupResourceType, deletionMode: CleanupDeletionMode): string {
  if (status === "deleted" && deletionMode === "permanent" && PERMANENT_DELETE_RESOURCE_TYPES.has(resourceType)) return "Permanently Deleted";
  return BASE_STATUS_STYLE[status].label;
}

/** A small spinning ring — Tailwind's animate-spin on a two-tone border circle, no custom keyframes. */
export function Spinner({ className = "h-3.5 w-3.5" }: { className?: string }) {
  return <span className={`inline-block animate-spin rounded-full border-2 border-slate-200 border-t-[#1b2fc4] ${className}`} aria-label="In progress" />;
}

/**
 * The third drill-down level (operation → item → file) — shared by CleanupProgressView (live,
 * while a cleanup is still running) and CleanupResultsView (after it finishes). Given one item's
 * id, paginated-fetches and renders its file list. Files still 'pending' get the spinner; this
 * component doesn't poll on its own — the parent view re-mounts/refetches it on its own poll cycle
 * while a cleanup is running, same as everything else on that page.
 */
export function ItemFilesDrilldown({
  operationId,
  itemId,
  resourceType,
  deletionMode,
}: {
  operationId: string;
  itemId: string;
  resourceType: CleanupResourceType;
  deletionMode: CleanupDeletionMode;
}) {
  const [page, setPage] = useState(1);
  const [files, setFiles] = useState<CleanupItemFileRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    getCleanupOperationItemFiles(operationId, itemId, { page, pageSize: PAGE_SIZE })
      .then((res) => {
        if (cancelled) return;
        setFiles(res.files);
        setTotal(res.total);
      })
      .catch((err) => {
        if (!cancelled) setError(err instanceof ApiClientError ? err.message : "Couldn't load files for this item.");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [operationId, itemId, page]);

  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));

  return (
    <div className="border-t border-slate-100 bg-slate-50/60">
      {loading ? (
        <p className="px-4 py-3 text-sm text-slate-500">Loading…</p>
      ) : error ? (
        <p className="px-4 py-3 text-sm text-rose-600">{error}</p>
      ) : files.length === 0 ? (
        <p className="px-4 py-3 text-sm text-slate-400">Nothing removed for this item yet.</p>
      ) : (
        <>
          <table className="w-full text-sm">
            <thead className="text-left text-xs uppercase tracking-wide text-slate-500">
              <tr>
                <th className="px-4 py-2 font-medium">File Name</th>
                <th className="px-4 py-2 font-medium">Status</th>
                <th className="px-4 py-2 font-medium text-right">Size</th>
                <th className="px-4 py-2 font-medium">Date</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {files.map((f) => (
                <tr key={f.id}>
                  <td className="px-4 py-2 text-slate-700">
                    {f.fileName}
                    {f.errorMessage && <div className="text-xs text-rose-500">{f.errorMessage}</div>}
                  </td>
                  <td className={`px-4 py-2 font-medium ${BASE_STATUS_STYLE[f.status].className}`}>
                    <span className="inline-flex items-center gap-1.5">
                      {f.status === "pending" && <Spinner />}
                      {fileStatusLabel(f.status, resourceType, deletionMode)}
                    </span>
                  </td>
                  <td className="px-4 py-2 text-right text-slate-500">{f.fileSizeBytes > 0 ? formatBytes(f.fileSizeBytes) : "—"}</td>
                  <td className="px-4 py-2 text-slate-500">{f.completedAt ? formatDate(f.completedAt) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {totalPages > 1 && <PageFooter page={page} totalPages={totalPages} total={total} onGoToPage={setPage} disabled={loading} />}
        </>
      )}
    </div>
  );
}
