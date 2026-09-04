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
  const cleanups = await query(
    `UPDATE cleanup_operations SET status = 'failed', completed_at = now(), error_message = $1
     WHERE status IN ('queued', 'running') RETURNING id`,
    [message]
  );

  const total = syncJobs.rows.length + scans.rows.length + cleanups.rows.length;
  if (total > 0) {
    console.warn(
      `[recovery] marked ${total} orphaned job(s) as failed on startup (sync: ${syncJobs.rows.length}, scans: ${scans.rows.length}, cleanups: ${cleanups.rows.length})`
    );
  }
}
