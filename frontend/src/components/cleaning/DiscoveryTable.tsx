import { useEffect, useState } from "react";

export interface DiscoveryColumn<T> {
  label: string;
  align?: "left" | "right";
  render: (row: T) => React.ReactNode;
}

function PageControls({ pageNumber, hasPrevious, hasNext, onPrevious, onNext, disabled }: { pageNumber: number; hasPrevious: boolean; hasNext: boolean; onPrevious: () => void; onNext: () => void; disabled: boolean }) {
  return (
    <div className="flex items-center gap-2 text-sm text-slate-600">
      <button
        onClick={onPrevious}
        disabled={!hasPrevious || disabled}
        className="rounded-md border border-slate-200 px-2.5 py-1 font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
        aria-label="Previous page"
      >
        ← Prev
      </button>
      <span className="min-w-[4.5rem] text-center text-xs text-slate-500">Page {pageNumber}</span>
      <button
        onClick={onNext}
        disabled={!hasNext || disabled}
        className="rounded-md border border-slate-200 px-2.5 py-1 font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
        aria-label="Next page"
      >
        Next →
      </button>
    </div>
  );
}

/**
 * Generic search + sort + select-all + paged table — shared by OneDrive Accounts, SharePoint
 * Sites, and Teams Direct Messages, since none of this existed anywhere in the app before
 * (verified: every other <table> in this codebase is a static, non-interactive listing).
 *
 * Prev/Next controls are shown both above and below the rows — above so switching pages never
 * requires scrolling down first (the original "Load more" link at the bottom only did).
 */
export function DiscoveryTable<T extends { id: string }>({
  title,
  columns,
  rows,
  loading,
  error,
  pageNumber,
  hasNext,
  hasPrevious,
  onNext,
  onPrevious,
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
  pageNumber: number;
  hasNext: boolean;
  hasPrevious: boolean;
  onNext: () => void;
  onPrevious: () => void;
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
  const showPageControls = hasNext || hasPrevious;

  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 px-4 py-3">
        <h3 className="text-sm font-semibold text-slate-800">{title}</h3>
        <div className="flex flex-wrap items-center gap-3">
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
          {showPageControls && (
            <PageControls pageNumber={pageNumber} hasPrevious={hasPrevious} hasNext={hasNext} onPrevious={onPrevious} onNext={onNext} disabled={loading} />
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
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
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
          {showPageControls && (
            <div className="flex justify-end border-t border-slate-100 px-4 py-3">
              <PageControls pageNumber={pageNumber} hasPrevious={hasPrevious} hasNext={hasNext} onPrevious={onPrevious} onNext={onNext} disabled={loading} />
            </div>
          )}
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
