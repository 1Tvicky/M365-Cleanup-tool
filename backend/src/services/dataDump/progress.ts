import { query } from "../../db/pool.js";
import type { ResourceSubKind } from "../../types/dataDump.js";

export interface ProgressDelta {
  created: number;
  failed: number;
  skipped: number;
  sizeBytes: number;
}

/** Applies an incremental progress update to one workload task AND its parent operation in one round trip each — mirrors cleanupExecutionWorker.ts's `processed_items = processed_items + 1` pattern, just batched to once per BatchAccumulator flush instead of once per item, since Data Dump volumes can be orders of magnitude larger. */
export async function applyTaskProgress(taskId: string, operationId: string, delta: ProgressDelta): Promise<void> {
  await query(
    `UPDATE data_dump_workload_tasks
     SET created_items = created_items + $2, failed_items = failed_items + $3, skipped_items = skipped_items + $4, total_size_bytes = total_size_bytes + $5
     WHERE id = $1`,
    [taskId, delta.created, delta.failed, delta.skipped, delta.sizeBytes]
  );
  await query(
    `UPDATE data_dump_operations
     SET created_items = created_items + $2, failed_items = failed_items + $3, skipped_items = skipped_items + $4, total_size_bytes = total_size_bytes + $5
     WHERE id = $1`,
    [operationId, delta.created, delta.failed, delta.skipped, delta.sizeBytes]
  );
}

/**
 * Bumps one resource kind's requested/created/failed/skipped counters inside a task's `subcounts`
 * JSONB (spec §27 — every kind reconciles independently). A single jsonb_set expression rather than
 * read-then-write: two runners can call this concurrently for the same task (rare, since runners
 * are per-workload-task, but sub-kinds like "channel" and "message" can be bumped from overlapping
 * async work within one runner) without a lost-update race the way a JS-side read/merge/write would have.
 */
export async function bumpSubcount(taskId: string, kind: ResourceSubKind, delta: { requested?: number; created?: number; failed?: number; skipped?: number }): Promise<void> {
  const zero = { requested: 0, created: 0, failed: 0, skipped: 0 };
  const add = { ...zero, ...delta };
  await query(
    `UPDATE data_dump_workload_tasks
     SET subcounts = jsonb_set(
       subcounts,
       $2::text[],
       jsonb_build_object(
         'requested', COALESCE((subcounts #>> $2::text[])::jsonb ->> 'requested', '0')::int + $3,
         'created', COALESCE((subcounts #>> $2::text[])::jsonb ->> 'created', '0')::int + $4,
         'failed', COALESCE((subcounts #>> $2::text[])::jsonb ->> 'failed', '0')::int + $5,
         'skipped', COALESCE((subcounts #>> $2::text[])::jsonb ->> 'skipped', '0')::int + $6
       ),
       true
     )
     WHERE id = $1`,
    [taskId, [kind], add.requested, add.created, add.failed, add.skipped]
  );
}

export async function setTaskCheckpoint(taskId: string, checkpoint: Record<string, unknown>): Promise<void> {
  await query(`UPDATE data_dump_workload_tasks SET checkpoint = $2 WHERE id = $1`, [taskId, JSON.stringify(checkpoint)]);
}

/** Cooperative pause/cancel signal — checked between units of work, same "checked at boundaries, not mid-item" shape as cleanupExecutionWorker.ts's isCancelled(). Pause is Data-Dump-specific (Cleanup has no pause, only cancel); a paused operation's checkpoint is left exactly where it stopped so a later Resume continues from there. */
export async function readControlSignal(operationId: string): Promise<{ cancelled: boolean; paused: boolean }> {
  const result = await query<{ cancel_requested_at: string | null; pause_requested_at: string | null }>(
    `SELECT cancel_requested_at, pause_requested_at FROM data_dump_operations WHERE id = $1`,
    [operationId]
  );
  const row = result.rows[0];
  return { cancelled: row?.cancel_requested_at != null, paused: row?.pause_requested_at != null };
}
