import { useCallback, useEffect, useRef, useState } from "react";
import { listAvailableResources, type AvailableResourceRow } from "../../api/clouds";
import { ApiClientError } from "../../api/client";
import { DiscoveryTable, useDebouncedValue, type DiscoveryColumn } from "../cleaning/DiscoveryTable";

type Fetcher = (opts: { search?: string; page?: number; pageSize?: number }) => Promise<{ resources: AvailableResourceRow[]; total: number }>;

/** Fetches from a specific workload connection's live resource browser (GET /api/clouds/:id/available-resources). */
export function connectionResourceFetcher(connectionId: string): Fetcher {
  return (opts) => listAvailableResources(connectionId, opts);
}

/**
 * Reuses Add Clouds' existing live resource browser and Cleaning's DiscoveryTable UNCHANGED — this
 * is exactly the "shared infrastructure: resource discovery, pagination, common UI components" the
 * spec calls for, not a Data-Dump-specific reimplementation. Same search/paginate/select-all/
 * selected-count behavior an operator already knows from Cleanup's own resource tables (spec
 * §3/§7/§13/§19). `fetcher` is pluggable so this same component also powers the Teams member picker
 * (api/dataDump.ts's listTenantUsers, a tenant-wide listing with no single connectionId).
 */
export function ResourceSelectionStep({
  fetcher,
  title,
  subtitle,
  searchPlaceholder,
  secondaryLabel,
  selected,
  onToggle,
  onToggleAll,
  emptyMessage,
}: {
  fetcher: Fetcher;
  title: string;
  subtitle: string;
  searchPlaceholder: string;
  secondaryLabel: string;
  selected: Map<string, AvailableResourceRow>;
  onToggle: (id: string, row: AvailableResourceRow) => void;
  onToggleAll: (rows: AvailableResourceRow[]) => void;
  emptyMessage: string;
}) {
  const [search, setSearch] = useState("");
  const debouncedSearch = useDebouncedValue(search, 300);
  const [page, setPage] = useState(1);
  const [rows, setRows] = useState<AvailableResourceRow[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const requestId = useRef(0);

  const load = useCallback(
    (targetPage: number) => {
      const id = ++requestId.current;
      setLoading(true);
      setError(null);
      fetcher({ search: debouncedSearch, page: targetPage, pageSize: 20 })
        .then((res) => {
          if (id !== requestId.current) return;
          setRows(res.resources);
          setTotal(res.total);
        })
        .catch((err) => {
          if (id !== requestId.current) return;
          setError(err instanceof ApiClientError ? err.message : "Couldn't load resources.");
        })
        .finally(() => {
          if (id === requestId.current) setLoading(false);
        });
    },
    [fetcher, debouncedSearch]
  );

  useEffect(() => setPage(1), [debouncedSearch]);
  useEffect(() => load(page), [load, page]);

  const columns: DiscoveryColumn<AvailableResourceRow>[] = [
    { label: "Name", render: (r) => r.displayName },
    { label: secondaryLabel, render: (r) => r.secondary ?? "—" },
  ];

  return (
    <div>
      <h2 className="mb-1 text-lg font-semibold text-slate-800">{title}</h2>
      <p className="mb-5 text-sm text-slate-500">{subtitle}</p>
      <DiscoveryTable
        title={title}
        columns={columns}
        rows={rows}
        loading={loading}
        error={error}
        page={page}
        totalPages={Math.max(1, Math.ceil(total / 20))}
        total={total}
        onGoToPage={setPage}
        search={search}
        onSearchChange={setSearch}
        searchPlaceholder={searchPlaceholder}
        selected={new Set(selected.keys())}
        onToggle={(id) => {
          const row = rows.find((r) => r.id === id);
          if (row) onToggle(id, row);
        }}
        onToggleAll={() => onToggleAll(rows)}
        emptyMessage={emptyMessage}
      />
    </div>
  );
}
