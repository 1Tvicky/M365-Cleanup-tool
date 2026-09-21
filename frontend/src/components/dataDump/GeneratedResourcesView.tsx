import { useEffect, useState } from "react";
import { getDataDumpResources, type DataDumpContainerRow, type DataDumpWorkload } from "../../api/dataDump";
import { ApiClientError } from "../../api/client";
import { formatBytes } from "../../utils/format";

/**
 * "View Generated Resources" (spec §37): the actual per-resource detail an operator needs to verify
 * what was created — which user's OneDrive, which SharePoint site and library, which Team and
 * channel, which Outlook mailbox — not just an aggregate count. Mirrors Cleanup's own per-item
 * drill-down in spirit (CleanupProgress.tsx's CategoryRow → CategoryItemsList), but built around
 * Data Dump's own container hierarchy (structural resources, not one row per file) instead of
 * Cleanup's per-item table.
 */

const KIND_LABEL: Record<string, string> = {
  root_folder: "OneDrive",
  folder: "Folder",
  site: "Site",
  site_library: "Library",
  team: "Team",
  channel: "Channel",
  mailbox_folder: "Mailbox",
};

interface TreeNode extends DataDumpContainerRow {
  children: TreeNode[];
}

function buildForest(rows: DataDumpContainerRow[]): Map<DataDumpWorkload, TreeNode[]> {
  const byId = new Map<string, TreeNode>(rows.map((r) => [r.id, { ...r, children: [] }]));
  const roots = new Map<DataDumpWorkload, TreeNode[]>();
  for (const row of rows) {
    const node = byId.get(row.id)!;
    if (row.parentContainerId && byId.has(row.parentContainerId)) {
      byId.get(row.parentContainerId)!.children.push(node);
    } else {
      if (!roots.has(row.workload)) roots.set(row.workload, []);
      roots.get(row.workload)!.push(node);
    }
  }
  return roots;
}

function ContainerNode({ node, depth }: { node: TreeNode; depth: number }) {
  const [expanded, setExpanded] = useState(depth < 1);
  const hasChildren = node.children.length > 0;
  const hasLeafCounts = node.createdCount > 0 || node.failedCount > 0;

  return (
    <div style={{ marginLeft: depth * 16 }}>
      <div
        role={hasChildren ? "button" : undefined}
        tabIndex={hasChildren ? 0 : undefined}
        onClick={() => hasChildren && setExpanded((e) => !e)}
        className={`flex items-center justify-between gap-3 rounded-md px-2 py-1.5 text-sm ${hasChildren ? "cursor-pointer hover:bg-slate-50" : ""}`}
      >
        <span className="flex items-center gap-2 text-slate-700">
          {hasChildren && <span className={`inline-block text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`}>▸</span>}
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">{KIND_LABEL[node.kind] ?? node.kind}</span>
          <span className="font-medium">{node.displayName}</span>
        </span>
        {hasLeafCounts && (
          <span className="shrink-0 text-xs text-slate-500">
            <span className="text-emerald-600">{node.createdCount.toLocaleString()} created</span>
            {node.failedCount > 0 && <span className="text-rose-600"> · {node.failedCount.toLocaleString()} failed</span>}
            {node.sizeBytes > 0 && <span> · {formatBytes(node.sizeBytes)}</span>}
          </span>
        )}
      </div>
      {expanded && hasChildren && (
        <div>
          {node.children.map((child) => (
            <ContainerNode key={child.id} node={child} depth={depth + 1} />
          ))}
        </div>
      )}
    </div>
  );
}

const WORKLOAD_LABELS: Record<DataDumpWorkload, string> = {
  onedrive: "OneDrive",
  sharepoint: "SharePoint",
  teams: "Microsoft Teams",
  outlook: "Outlook",
};

export function GeneratedResourcesView({ operationId }: { operationId: string }) {
  const [containers, setContainers] = useState<DataDumpContainerRow[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    getDataDumpResources(operationId)
      .then((res) => setContainers(res.containers))
      .catch((err) => setError(err instanceof ApiClientError ? err.message : "Couldn't load generated resources."));
  }, [operationId]);

  if (error) return <p className="text-sm text-rose-600">{error}</p>;
  if (!containers) return <p className="text-sm text-slate-400">Loading…</p>;
  if (containers.length === 0) return <p className="text-sm text-slate-400">Nothing created yet.</p>;

  const forest = buildForest(containers);

  return (
    <div className="space-y-4">
      {[...forest.entries()].map(([workload, roots]) => (
        <div key={workload}>
          <h4 className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-500">{WORKLOAD_LABELS[workload]}</h4>
          <div className="rounded-lg border border-slate-100 bg-white p-2">
            {roots.map((node) => (
              <ContainerNode key={node.id} node={node} depth={0} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
