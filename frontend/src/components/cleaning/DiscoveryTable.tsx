import { useEffect, useState } from "react";

export interface DiscoveryColumn<T> {
  label: string;
  align?: "left" | "right";
  render: (row: T) => React.ReactNode;
}

/** "Showing page X of Y" + a direct page-jump dropdown — no Prev/Next buttons needed at this scale. Exported for reuse by the Cleanup Results table, which shares the same footer but isn't a selection table. */
export function PageFooter({ page, totalPages, total, onGoToPage, disabled }: { page: number; totalPages: number; total: number; onGoToPage: (page: number) => void; disabled: boolean }) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 bg-slate-50/60 px-4 py-2.5 text-sm text-slate-500">
      <span>{total.toLocaleString()} total</span>
      <div className="flex items-center gap-2">
        <span>
          Showing page {page} of {totalPages}
        </span>
        {totalPages > 1 && (
          <label className="flex items-center gap-1.5">
            Go to:
            <select
              value={page}
              disabled={disabled}
              onChange={(e) => onGoToPage(Number(e.target.value))}
              className="rounded-md border border-slate-200 bg-white px-2 py-1 text-slate-700 focus:border-[#1b2fc4] focus:outline-none"
            >
              {Array.from({ length: totalPages }, (_, i) => i + 1).map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
          </label>
        )}
      </div>
    </div>
  );
}

/**
 * Generic search + sort + select-all + paged table — shared by OneDrive Accounts, SharePoint
 * Sites, and Teams Direct Messages, since none of this existed anywhere in the app before
 * (verified: every other <table> in this codebase is a static, non-interactive listing).
 *
 * The row area has a fixed max-height with its own internal scrollbar — the header (search/sort)
 * and footer (page count + jump-to-page) stay permanently visible without ever scrolling the page
 * itself, no matter how many rows are on a page. Matches the reference product's table layout.
 */
export function DiscoveryTable<T extends { id: string }>({
  title,
  columns,
  rows,
  loading,
  error,
  page,
  totalPages,
  total,
  onGoToPage,
  search,
  onSearchChange,
  searchPlaceholder,
  sortOptions,
  sort,
  onSortChange,
  selected,
  onToggle,
  onToggleAll,
  emptyMessage,
}: {
  title: string;
  columns: DiscoveryColumn<T>[];
  rows: T[];
  loading: boolean;
  error: string | null;
  page: number;
  totalPages: number;
  total: number;
  onGoToPage: (page: number) => void;
  search: string;
  onSearchChange: (value: string) => void;
  searchPlaceholder: string;
  sortOptions?: { value: string; label: string }[];
  sort?: string;
  onSortChange?: (value: string) => void;
  selected: Set<string>;
  onToggle: (id: string) => void;
  onToggleAll: () => void;
  emptyMessage: string;
}) {
  const allSelected = rows.length > 0 && rows.every((r) => selected.has(r.id));

  return (
    <div className="flex max-h-[65vh] flex-col overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 px-4 py-3">
        <h3 className="text-sm font-semibold text-slate-800">{title}</h3>
        <div className="flex items-center gap-3">
          <input
            type="search"
            value={search}
            onChange={(e) => onSearchChange(e.target.value)}
            placeholder={searchPlaceholder}
            className="w-52 rounded-md border border-slate-200 px-3 py-1.5 text-sm text-slate-700 placeholder:text-slate-400 focus:border-[#1b2fc4] focus:outline-none"
          />
          {sortOptions && (
            <select
              value={sort}
              onChange={(e) => onSortChange?.(e.target.value)}
              className="rounded-md border border-slate-200 px-2 py-1.5 text-sm text-slate-600 focus:border-[#1b2fc4] focus:outline-none"
            >
              {sortOptions.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  Sort: {opt.label}
                </option>
              ))}
            </select>
          )}
        </div>
      </div>

      {loading ? (
        <p className="px-4 py-8 text-center text-sm text-slate-500">Loading…</p>
      ) : error ? (
        <p className="px-4 py-8 text-center text-sm text-rose-600">{error}</p>
      ) : rows.length === 0 ? (
        <p className="px-4 py-8 text-center text-sm text-slate-500">{emptyMessage}</p>
      ) : (
        <>
          <div className="overflow-y-auto overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                <tr>
                  <th className="w-10 px-4 py-3">
                    <input type="checkbox" checked={allSelected} onChange={onToggleAll} aria-label="Select all" />
                  </th>
                  {columns.map((col) => (
                    <th key={col.label} className={`px-4 py-3 font-medium ${col.align === "right" ? "text-right" : ""}`}>
                      {col.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {rows.map((row) => (
                  <tr key={row.id} className="hover:bg-slate-50/60">
                    <td className="px-4 py-3">
                      <input type="checkbox" checked={selected.has(row.id)} onChange={() => onToggle(row.id)} aria-label="Select row" />
                    </td>
                    {columns.map((col) => (
                      <td key={col.label} className={`px-4 py-3 text-slate-700 ${col.align === "right" ? "text-right" : ""}`}>
                        {col.render(row)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <PageFooter page={page} totalPages={totalPages} total={total} onGoToPage={onGoToPage} disabled={loading} />
        </>
      )}
    </div>
  );
}

/** Debounces free-typed search input so every keystroke doesn't re-fetch. */
export function useDebouncedValue<T>(value: T, delayMs: number): T {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    const handle = setTimeout(() => setDebounced(value), delayMs);
    return () => clearTimeout(handle);
  }, [value, delayMs]);
  return debounced;
}
