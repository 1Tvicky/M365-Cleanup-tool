import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import { getSiteById, getTeamById, getUserById, listAllUsers, searchSites, type BasicUser } from "../graph/cloudEnumeration.js";
import { filterRealMemberUsers } from "../services/dataDump/targetUsers.js";
import { connection as redis } from "./queue.js";
import type { DataDumpConfig, DataDumpWorkload, DataDumpWorkloadTaskStatus } from "../types/dataDump.js";
import { computeDataDumpPreview } from "../services/dataDump/preview.js";
import { readControlSignal } from "../services/dataDump/progress.js";
import { runOneDriveWorkload } from "../services/dataDump/runners/oneDriveRunner.js";
import { runSharePointWorkload } from "../services/dataDump/runners/sharePointRunner.js";
import { runTeamsWorkload } from "../services/dataDump/runners/teamsRunner.js";
import { runOutlookWorkload } from "../services/dataDump/runners/outlookRunner.js";
import type { RunnerContext, RunnerOutcome } from "../services/dataDump/runners/oneDriveRunner.js";

/**
 * Mirrors cleanupExecutionWorker.ts's overall shape (load operation by id from Postgres, never
 * trust the BullMQ payload beyond the id — spec §46) but for CREATING real M365 objects instead of
 * deleting them. Deliberately its own file, its own queue (jobs/queue.ts's dataDumpQueue), and its
 * own status vocabulary — never shares code with the Cleanup execution path beyond the generic
 * services/rateLimiter.ts throttling helper every workload/job in this app already reuses.
 */

interface OperationRow {
  id: string;
  tenant_id: string;
  workloads: DataDumpWorkload[];
  config: DataDumpConfig;
  label: string;
}

async function loadOperation(operationId: string): Promise<OperationRow> {
  const result = await query<OperationRow & { m365_tenant_id: string | null }>(
    `SELECT o.id, o.tenant_id, o.workloads, o.config, o.label, t.m365_tenant_id
     FROM data_dump_operations o
     JOIN tenants t ON t.id = o.tenant_id
     WHERE o.id = $1`,
    [operationId]
  );
  const row = result.rows[0];
  if (!row) throw new Error(`data_dump_operations ${operationId} not found`);
  return row;
}

async function getM365TenantId(tenantId: string): Promise<string> {
  const result = await query<{ m365_tenant_id: string | null }>(`SELECT m365_tenant_id FROM tenants WHERE id = $1`, [tenantId]);
  const id = result.rows[0]?.m365_tenant_id;
  if (!id) throw new Error(`Tenant ${tenantId} has no m365_tenant_id — Data Dump only supports Microsoft 365 tenants today`);
  return id;
}

interface TaskRow {
  id: string;
  status: DataDumpWorkloadTaskStatus;
}

async function getOrCreateWorkloadTask(operationId: string, tenantId: string, workload: DataDumpWorkload, requestedItems: number): Promise<TaskRow> {
  const existing = await query<TaskRow>(`SELECT id, status FROM data_dump_workload_tasks WHERE operation_id = $1 AND workload = $2`, [operationId, workload]);
  if (existing.rows[0]) return existing.rows[0];

  const connectionResult = await query<{ id: string }>(
    `SELECT id FROM connections WHERE tenant_id = $1 AND cloud_type = $2 AND status <> 'disconnected' ORDER BY connected_at DESC LIMIT 1`,
    [tenantId, workload]
  );
  const connectionId = connectionResult.rows[0]?.id;
  if (!connectionId) throw new Error(`No active ${workload} connection for tenant ${tenantId} — connect it first via Add Clouds`);

  const inserted = await query<TaskRow>(
    `INSERT INTO data_dump_workload_tasks (operation_id, connection_id, workload, requested_items)
     VALUES ($1, $2, $3, $4) RETURNING id, status`,
    [operationId, connectionId, workload, requestedItems]
  );
  return inserted.rows[0]!;
}

async function setTaskStatus(taskId: string, status: DataDumpWorkloadTaskStatus, errorMessage: string | null, terminal: boolean): Promise<void> {
  await query(
    `UPDATE data_dump_workload_tasks
     SET status = $2, error_message = $3, started_at = COALESCE(started_at, now()), completed_at = CASE WHEN $4 THEN now() ELSE completed_at END
     WHERE id = $1`,
    [taskId, status, errorMessage, terminal]
  );
}

function outcomeToTaskStatus(outcome: RunnerOutcome): DataDumpWorkloadTaskStatus {
  return outcome;
}

/** Resolves explicit ids (from the Select Resources step) via direct Graph lookups — never re-lists the whole tenant just to find a handful of already-known ids. Silently drops any id that no longer resolves (e.g. removed between selection and run) rather than failing the whole workload over one stale id. */
async function resolveUsers(client: Awaited<ReturnType<typeof graphClientForTenant>>, ids: string[]): Promise<{ id: string; upn: string; displayName: string }[]> {
  const users = await Promise.all(ids.map((id) => getUserById(client, id)));
  return users.filter((u): u is BasicUser => u !== null).map((u) => ({ id: u.id, upn: u.upn, displayName: u.displayName ?? u.upn }));
}

export const dataDumpWorker = new Worker(
  "data-dump-jobs",
  async (job) => {
    const { operationId } = job.data as { operationId: string };
    const operation = await loadOperation(operationId);
    const m365TenantId = await getM365TenantId(operation.tenant_id);
    const client = await graphClientForTenant(m365TenantId);

    await query(`UPDATE data_dump_operations SET status = 'running', started_at = COALESCE(started_at, now()) WHERE id = $1`, [operationId]);

    const preview = computeDataDumpPreview(operation.config.seed ?? 1, operation.workloads, operation.config);
    const requestedByWorkload = new Map(preview.workloads.map((w) => [w.workload, w.requestedObjects]));

    let anyFailed = false;
    let stoppedForPauseOrCancel: "paused" | "cancelled" | null = null;

    const isCancelled = async () => (await readControlSignal(operationId)).cancelled;
    const isPaused = async () => (await readControlSignal(operationId)).paused;

    for (const workload of operation.workloads) {
      if (stoppedForPauseOrCancel) break;
      if (await isCancelled()) {
        stoppedForPauseOrCancel = "cancelled";
        break;
      }
      if (await isPaused()) {
        stoppedForPauseOrCancel = "paused";
        break;
      }

      let task: TaskRow;
      try {
        task = await getOrCreateWorkloadTask(operationId, operation.tenant_id, workload, requestedByWorkload.get(workload) ?? 0);
      } catch (err) {
        anyFailed = true;
        console.error(`[data-dump] ${operationId}: failed to prepare ${workload} task`, err);
        continue;
      }

      if (task.status === "completed" || task.status === "skipped") continue;

      await setTaskStatus(task.id, "running", null, false);

      const ctx: RunnerContext = {
        client,
        operationId,
        taskId: task.id,
        operationSeed: operation.config.seed ?? 1,
        namingPrefix: operation.config.namingPrefix,
        dateRange: operation.config.dateRange,
        isCancelled,
        isPaused,
      };

      try {
        let outcome: RunnerOutcome;
        if (workload === "onedrive" && operation.config.onedrive) {
          const c = operation.config.onedrive;
          // Explicit selection (Select Users step) is honored exactly — no re-listing/slicing the
          // tenant once a real selection exists. Only the profile-driven quick-start path (no
          // selection at all) falls back to listAllUsers-and-slice.
          const users = c.selectedUserIds && c.selectedUserIds.length > 0
            ? await resolveUsers(client, c.selectedUserIds)
            : filterRealMemberUsers(await listAllUsers(client)).slice(0, c.userCount).map((u) => ({ id: u.id, upn: u.upn, displayName: u.displayName ?? u.upn }));
          outcome = await runOneDriveWorkload(ctx, users, c);
        } else if (workload === "sharepoint" && operation.config.sharepoint) {
          const c = operation.config.sharepoint;
          const sites = c.selectedSiteIds && c.selectedSiteIds.length > 0
            ? (await Promise.all(c.selectedSiteIds.map((id) => getSiteById(client, id)))).filter((s): s is NonNullable<typeof s> => s !== null).map((s) => ({ id: s.id, displayName: s.displayName }))
            : (await searchSites(client)).slice(0, c.siteCount).map((s) => ({ id: s.id, displayName: s.displayName }));
          const users = filterRealMemberUsers(await listAllUsers(client)).slice(0, 10).map((u) => ({ id: u.id, upn: u.upn, displayName: u.displayName ?? u.upn }));
          if (sites.length === 0 && !c.newSite) throw new Error("No existing SharePoint sites selected/found — select an existing site or configure a new one.");
          outcome = await runSharePointWorkload(ctx, sites, users, c);
        } else if (workload === "teams" && operation.config.teams) {
          const c = operation.config.teams;
          const selectedTeams = c.selectedTeamIds && c.selectedTeamIds.length > 0
            ? (await Promise.all(c.selectedTeamIds.map((id) => getTeamById(client, id)))).filter((t): t is NonNullable<typeof t> => t !== null).map((t) => ({ id: t.id, displayName: t.displayName }))
            : [];
          // The tenant-wide user pool used to staff new teams/channels — not the target of the
          // workload itself, so it's never sliced down to a "selection," just a reasonably large
          // real-member candidate pool for owner/member assignment.
          const candidateUsers = filterRealMemberUsers(await listAllUsers(client))
            .slice(0, Math.max(c.membersPerTeam, 25))
            .map((u) => ({ id: u.id, upn: u.upn, displayName: u.displayName ?? u.upn }));
          if (candidateUsers.length === 0 && (c.newTeam || selectedTeams.length === 0)) throw new Error("No existing tenant users found to own/staff generated Teams.");
          outcome = await runTeamsWorkload(ctx, candidateUsers, c, selectedTeams);
        } else if (workload === "outlook" && operation.config.outlook) {
          const c = operation.config.outlook;
          const users = c.selectedUserIds && c.selectedUserIds.length > 0
            ? await resolveUsers(client, c.selectedUserIds)
            : filterRealMemberUsers(await listAllUsers(client)).slice(0, c.userCount).map((u) => ({ id: u.id, upn: u.upn, displayName: u.displayName ?? u.upn }));
          outcome = await runOutlookWorkload(ctx, users, c);
        } else {
          outcome = "completed"; // workload selected without its config section — nothing to do, not an error
        }

        if (outcome === "paused" || outcome === "cancelled") {
          await setTaskStatus(task.id, outcome, null, false);
          stoppedForPauseOrCancel = outcome;
        } else {
          if (outcome === "completed_with_errors") anyFailed = true;
          await setTaskStatus(task.id, outcomeToTaskStatus(outcome), null, true);
        }
      } catch (err) {
        anyFailed = true;
        const message = err instanceof Error ? err.message : String(err);
        console.error(`[data-dump] ${operationId}: ${workload} workload failed`, err);
        await setTaskStatus(task.id, "failed", message, true);
      }
    }

    if (stoppedForPauseOrCancel === "cancelled") {
      await query(`UPDATE data_dump_operations SET status = 'cancelled', completed_at = now() WHERE id = $1`, [operationId]);
    } else if (stoppedForPauseOrCancel === "paused") {
      await query(`UPDATE data_dump_operations SET status = 'paused', pause_requested_at = COALESCE(pause_requested_at, now()) WHERE id = $1`, [operationId]);
    } else {
      const finalStatus = anyFailed ? "completed_with_errors" : "completed";
      await query(`UPDATE data_dump_operations SET status = $2, completed_at = now() WHERE id = $1`, [operationId, finalStatus]);
    }
  },
  { connection: redis, concurrency: 2 }
);
