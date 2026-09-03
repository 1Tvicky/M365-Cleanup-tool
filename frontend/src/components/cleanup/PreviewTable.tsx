import { formatBytes, formatDate } from "../../utils/format";
import type { Preview } from "../../types";

/** Dry-run preview — no destructive calls have been made to reach this screen. See docs/api-spec.md. */
export function PreviewTable({ preview, onConfirm }: { preview: Preview; onConfirm: () => void }) {
  return (
    <div className="mt-8">
      <div className="mb-1 flex items-center gap-2">
        <h2 className="text-lg font-semibold text-slate-800">3. Preview (dry-run)</h2>
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500">
          No changes made yet
        </span>
      </div>
      <p className="mb-4 text-sm text-slate-500">Generated {formatDate(preview.generatedAt)}. Review before confirming.</p>

      {preview.warnings.length > 0 && (
        <div className="mb-4 space-y-1.5 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3">
          {preview.warnings.map((w, i) => (
            <div key={i} className="flex gap-2 text-sm text-amber-800">
              <span aria-hidden>⚠️</span>
              {w}
            </div>
          ))}
        </div>
      )}

      <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-4 py-3 font-medium">Item</th>
              <th className="px-4 py-3 font-medium">Category</th>
              <th className="px-4 py-3 font-medium text-right">Count</th>
              <th className="px-4 py-3 font-medium text-right">Size</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {preview.rows.map((row, i) => (
              <tr key={i} className="hover:bg-slate-50/60">
                <td className="px-4 py-3 font-medium text-slate-800">{row.label}</td>
                <td className="px-4 py-3 text-slate-600">
                  {row.category}
                  {row.note && <div className="text-xs text-slate-400">{row.note}</div>}
                </td>
                <td className="px-4 py-3 text-right text-slate-600">{row.itemCount.toLocaleString()}</td>
                <td className="px-4 py-3 text-right text-slate-600">
                  {row.sizeBytes > 0 ? formatBytes(row.sizeBytes) : "—"}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="border-t-2 border-slate-200 bg-slate-50 font-semibold text-slate-800">
              <td className="px-4 py-3" colSpan={2}>
                Total
              </td>
              <td className="px-4 py-3 text-right">{preview.totals.itemCount.toLocaleString()}</td>
              <td className="px-4 py-3 text-right">{formatBytes(preview.totals.totalSizeBytes)}</td>
            </tr>
          </tfoot>
        </table>
      </div>

      <button
        onClick={onConfirm}
        className="mt-6 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-rose-700"
      >
        Review & confirm deletion →
      </button>
    </div>
  );
}
