import { query } from "../db/pool.js";

/**
 * A row can be left permanently stuck at status 'queued'/'running' if the process handling it dies
 * (a crash, a deploy, a forced restart) after BullMQ gives up redelivering the job (its default
 * maxStalledCount is 1 retry) — nothing else ever calls back in to move that row to a terminal
 * status, so the UI would show "Syncing…"/"Cleaning…" forever with no way to tell the user it's
 * actually dead. Run once at process startup: anything already 'queued'/'running' at boot cannot
 * belong to *this* process (it hasn't dispatched any work yet), so it must be left over from a
 * previous process's lifetime. If BullMQ does still redeliver that job a little after this runs,
 * the worker's own `status = 'running'` update on pickup simply supersedes this — this is a
 * best-effort safety net, not a lock, and it's harmless to race against a legitimate redelivery.
 */
export async function recoverOrphanedJobs(): Promise<void> {
  const message = "Interrupted by a server restart before it finished. Please try again.";

  const syncJobs = await query(
    `UPDATE sync_jobs SET status = 'failed', finished_at = now(), error_log = error_log || $1::jsonb
     WHERE status IN ('queued', 'running') RETURNING id`,
    [JSON.stringify([{ message, at: new Date().toISOString() }])]
  );
  const scans = await query(
    `UPDATE cleaning_scans SET status = 'failed', finished_at = now(), error_log = error_log || $1::jsonb
     WHERE status IN ('queued', 'running') RETURNING id`,
    [JSON.stringify([{ message, at: new Date().toISOString() }])]
  );
  const cleanups = await query<{ id: string }>(
    `UPDATE cleanup_operations SET status = 'failed', completed_at = now(), error_message = $1
     WHERE status IN ('queued', 'running') RETURNING id`,
    [message]
  );

  // The parent operation above is now a terminal 'failed', but nothing else ever advances
  // cleanup_operation_items.status except this worker's own onItemSettled callback — which never
  // ran for whatever this crash interrupted. Left unfixed, those items stay stuck at 'pending'
  // forever, and POST /cleanup/:operationId/retry (which only rescopes status='failed' items) finds
  // nothing to retry, silently hiding that some of them may have already been permanently deleted
  // before the crash. Same "bulk-flip still-pending rows to a terminal status" pattern already used
  // by jobs/cloudSyncWorker.ts's cancelRemainingResources.
  let orphanedItemCount = 0;
  if (cleanups.rows.length > 0) {
    const orphanedItems = await query(
      `UPDATE cleanup_operation_items SET status = 'failed', error_message = $2, updated_at = now()
       WHERE cleanup_operation_id = ANY($1::uuid[]) AND status IN ('pending', 'processing')`,
      [cleanups.rows.map((r) => r.id), message]
    );
    orphanedItemCount = orphanedItems.rowCount ?? 0;
  }

  // Same "queued/running at boot cannot belong to this process" reasoning, applied to Data Dump
  // (jobs/dataDumpWorker.ts) — a workload task frozen mid-generation by a crash must not be left
  // silently stuck at 'running' forever with no way for an operator to tell it's actually dead.
  // Unlike Cleanup, a Data Dump operation is NOT flipped to a hard 'failed' here — it's flipped to
  // 'paused' instead, since (a) partially-generated data is a normal, safe, resumable state (not a
  // destructive-audit concern the way an interrupted cleanup is), and (b) the containers/batches
  // already recorded make it directly resumable via POST /:id/resume, which is the ordinary
  // operator recovery action for a stalled run, not a "start over" one.
  const dataDumpOps = await query<{ id: string }>(
    `UPDATE data_dump_operations SET status = 'paused', pause_requested_at = now()
     WHERE status IN ('queued', 'running') RETURNING id`
  );
  const dataDumpTasks = await query(
    `UPDATE data_dump_workload_tasks SET status = 'paused'
     WHERE status = 'running' AND operation_id = ANY($1::uuid[])`,
    [dataDumpOps.rows.map((r) => r.id)]
  );

  const total = syncJobs.rows.length + scans.rows.length + cleanups.rows.length + dataDumpOps.rows.length;
  if (total > 0) {
    console.warn(
      `[recovery] marked ${total} orphaned job(s) as failed/paused on startup (sync: ${syncJobs.rows.length}, scans: ${scans.rows.length}, cleanups: ${cleanups.rows.length}, cleanup items: ${orphanedItemCount}, data dump operations paused: ${dataDumpOps.rows.length}, tasks: ${dataDumpTasks.rowCount ?? 0})`
    );
  }
}
