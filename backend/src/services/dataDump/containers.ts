import { query } from "../../db/pool.js";

export type ContainerKind = "root_folder" | "folder" | "team" | "channel" | "site" | "site_library" | "mailbox_folder";

export interface ContainerRow {
  id: string;
  graphId: string;
  displayName: string;
  parentContainerId: string | null;
}

/**
 * Resumability for structural containers (spec §25): before creating a folder/team/channel/library,
 * check whether this workload task already recorded one with the same display name under the same
 * parent — a resumed run reuses it instead of creating a duplicate or erroring. Bounded query (a
 * task's own containers are always a small set, never per-file), safe to call once per container.
 */
export async function findExistingContainer(workloadTaskId: string, kind: ContainerKind, displayName: string, parentContainerId: string | null): Promise<ContainerRow | null> {
  const result = await query<{ id: string; graph_id: string; display_name: string; parent_container_id: string | null }>(
    `SELECT id, graph_id, display_name, parent_container_id FROM data_dump_containers
     WHERE workload_task_id = $1 AND kind = $2 AND display_name = $3 AND parent_container_id IS NOT DISTINCT FROM $4
     LIMIT 1`,
    [workloadTaskId, kind, displayName, parentContainerId]
  );
  const row = result.rows[0];
  if (!row) return null;
  return { id: row.id, graphId: row.graph_id, displayName: row.display_name, parentContainerId: row.parent_container_id };
}

export async function recordContainer(workloadTaskId: string, kind: ContainerKind, graphId: string, displayName: string, parentContainerId: string | null): Promise<ContainerRow> {
  const result = await query<{ id: string }>(
    `INSERT INTO data_dump_containers (workload_task_id, kind, graph_id, display_name, parent_container_id)
     VALUES ($1, $2, $3, $4, $5) RETURNING id`,
    [workloadTaskId, kind, graphId, displayName, parentContainerId]
  );
  return { id: result.rows[0]!.id, graphId, displayName, parentContainerId };
}

/** Get-or-create: reuses an existing container row (resume case) or calls `create` and records a new one. */
export async function getOrCreateContainer(
  workloadTaskId: string,
  kind: ContainerKind,
  displayName: string,
  parentContainerId: string | null,
  create: () => Promise<{ graphId: string }>
): Promise<ContainerRow> {
  const existing = await findExistingContainer(workloadTaskId, kind, displayName, parentContainerId);
  if (existing) return existing;
  const { graphId } = await create();
  return recordContainer(workloadTaskId, kind, graphId, displayName, parentContainerId);
}
