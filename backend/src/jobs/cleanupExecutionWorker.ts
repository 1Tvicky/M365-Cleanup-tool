import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import { classifyDeleteError, deleteDriveItem, listDriveRootChildren, type DriveOwnerKind } from "../graph/cleanupDeletion.js";
import { runThrottled } from "../services/rateLimiter.js";
import { connection as redis } from "./queue.js";

/**
 * Mirrors jobs/cleaningScanWorker.ts's shape (same join-by-id pattern, same runThrottled usage,
 * same reauth-error branching, same connection_events logging), but for executing a confirmed
 * cleanup operation. Only 'onedrive_account'/'sharepoint_site' items are ever processed here —
 * 'channel'/'chat' items are created with status='unsupported' at manifest time (routes/cleaning.ts)
 * and never selected by the `status = 'pending'` query below, so this worker never attempts a Graph
 * call Microsoft doesn't support for this app's application-only permissions.
 *
 * All items in one operation share the same tenant (enforced at manifest-creation time), so only
 * one graphClientForTenant call is needed for the whole job, regardless of how many of the tenant's
 * up-to-3 connections (OneDrive/SharePoint/Teams) are actually touched.
 */

async function logConnectionEvent(event: string, connectionId: string, tenantId: string, detail: Record<string, unknown> = {}): Promise<void> {
  await query(`INSERT INTO connection_events (connection_id, tenant_id, event, detail) VALUES ($1, $2, $3, $4)`, [
    connectionId,
    tenantId,
    event,
    detail,
  ]);
}

function isReauthError(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  const code = String((err as { code?: string; body?: string })?.code ?? (err as { body?: string })?.body ?? "");
  return status === 401 || /InvalidAuthenticationToken|consent_required|invalid_grant|AuthenticationError/i.test(code);
}

async function isCancelled(operationId: string): Promise<boolean> {
  const result = await query<{ cancel_requested_at: string | null }>(`SELECT cancel_requested_at FROM cleanup_operations WHERE id = $1`, [
    operationId,
  ]);
  return result.rows[0]?.cancel_requested_at != null;
}

interface PendingItem {
  id: string;
  connection_id: string;
  resource_type: "onedrive_account" | "sharepoint_site";
  graph_ref: { userId?: string; siteId?: string };
}

/** Deletes every top-level file/folder inside the account's OneDrive or the site's default document library. Never touches the account or site itself. */
async function executeItem(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  item: PendingItem,
  operationId: string
): Promise<void> {
  const kind: DriveOwnerKind = item.resource_type === "onedrive_account" ? "user" : "site";
  const ownerId = (item.resource_type === "onedrive_account" ? item.graph_ref.userId : item.graph_ref.siteId)!;

  const children = await listDriveRootChildren(client, kind, ownerId);
  if (children.length === 0) return; // already empty — nothing to do, counts as success

  // Seed one row per file up front — so the report/live "recently removed" feed reflects the full
  // file list from the start, not just the ones that have finished so far.
  for (const child of children) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, child.name, child.id, child.size]
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, children.length]);

  let firstError: unknown = null;
  await runThrottled(children, (child) => deleteDriveItem(client, kind, ownerId, child.id), {
    isCancelled: () => isCancelled(operationId),
    batchSize: 10,
    onItemSettled: async (child, result) => {
      const fileStatus = result.ok ? result.value : "failed"; // "deleted" | "already_gone" | "failed"
      const errorMessage = result.ok ? null : classifyDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, child.id, fileStatus, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}

export const cleanupExecutionWorker = new Worker(
  "cleanup-execution-jobs",
  async (job) => {
    const { operationId } = job.data as { operationId: string };

    const opRow = await query<{ tenant_id: string; m365_tenant_id: string }>(
      `SELECT co.tenant_id, t.m365_tenant_id
       FROM cleanup_operations co
       JOIN tenants t ON t.id = co.tenant_id
       WHERE co.id = $1`,
      [operationId]
    );
    const info = opRow.rows[0];
    if (!info) throw new Error(`cleanup_operations ${operationId} not found`);

    const pending = await query<PendingItem>(
      `SELECT id, connection_id, resource_type, graph_ref
       FROM cleanup_operation_items
       WHERE cleanup_operation_id = $1 AND status = 'pending'`,
      [operationId]
    );

    await query(`UPDATE cleanup_operations SET status = 'running', started_at = now() WHERE id = $1`, [operationId]);
    const touchedConnections = [...new Set(pending.rows.map((r) => r.connection_id))];
    for (const connectionId of touchedConnections) {
      await logConnectionEvent("cleanup_started", connectionId, info.tenant_id, { operationId });
    }

    try {
      const client = await graphClientForTenant(info.m365_tenant_id);
      let successful = 0;
      let failed = 0;

      await runThrottled(pending.rows, (item) => executeItem(client, item, operationId), {
        isCancelled: () => isCancelled(operationId),
        batchSize: 3, // conservative — each item itself fans out into its own (throttled) per-file deletes; tune after the first live test run
        onItemSettled: async (item, result) => {
          if (result.ok) {
            successful++;
            await query(
              `UPDATE cleanup_operation_items SET status = 'completed', attempts = attempts + 1, completed_at = now(), updated_at = now() WHERE id = $1`,
              [item.id]
            );
          } else {
            failed++;
            const { code, message } = classifyDeleteError(result.error);
            await query(
              `UPDATE cleanup_operation_items
               SET status = 'failed', attempts = attempts + 1, completed_at = now(), updated_at = now(), error_code = $2, error_message = $3
               WHERE id = $1`,
              [item.id, code, message]
            );
          }
          await query(`UPDATE cleanup_operations SET processed_items = processed_items + 1 WHERE id = $1`, [operationId]);
        },
      });

      await query(`UPDATE cleanup_operations SET successful_items = $2, failed_items = $3 WHERE id = $1`, [operationId, successful, failed]);

      const cancelled = await isCancelled(operationId);
      const finalStatus = cancelled ? "cancelled" : failed > 0 ? "completed_with_errors" : "completed";
      await query(`UPDATE cleanup_operations SET status = $2, completed_at = now() WHERE id = $1`, [operationId, finalStatus]);
      for (const connectionId of touchedConnections) {
        await logConnectionEvent(
          finalStatus === "completed_with_errors" ? "cleanup_completed_with_errors" : finalStatus === "cancelled" ? "cleanup_cancelled" : "cleanup_completed",
          connectionId,
          info.tenant_id,
          { operationId, successful, failed }
        );
      }
    } catch (err) {
      if (isReauthError(err)) {
        for (const connectionId of touchedConnections) {
          await query(`UPDATE connections SET status = 'needs_reauth', last_error = $2 WHERE id = $1`, [connectionId, String(err)]);
        }
      }
      await query(`UPDATE cleanup_operations SET status = 'failed', completed_at = now(), error_message = $2 WHERE id = $1`, [
        operationId,
        String(err),
      ]);
      for (const connectionId of touchedConnections) {
        await logConnectionEvent("cleanup_failed", connectionId, info.tenant_id, { operationId, error: String(err) });
      }
      throw err;
    }
  },
  { connection: redis, concurrency: 3 }
);
