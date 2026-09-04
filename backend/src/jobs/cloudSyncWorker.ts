import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import {
  getSiteDriveQuota,
  getUserDriveQuota,
  getUserJoinedTeamsCount,
  listAllUsers,
  searchSites,
  type BasicUser,
  type SiteSummary,
} from "../graph/cloudEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import type { CloudType } from "../types/connections.js";
import { connection as redis } from "./queue.js";

interface ConnectionUserUpsert {
  graphUserId: string;
  upn: string;
  displayName: string | null;
  storageUsedBytes: number;
  itemCount: number;
  syncStatus: "synced" | "failed";
  errorMessage: string | null;
}

/**
 * Deletes connection_users rows left over from a previous sync whose subject (a user for
 * onedrive/teams, a site for sharepoint — connection_users doubles for both, scoped 1:1 by
 * connection_id since each connection is a single cloud_type) is no longer in the current
 * enumeration. Without this, someone excluded by a filter change (or removed/disabled in the
 * tenant) would linger forever, permanently inflating "not added" counts beyond the current
 * sync_jobs.total_users.
 */
async function pruneStaleConnectionUsers(connectionId: string, currentIds: string[]): Promise<void> {
  await query(`DELETE FROM connection_users WHERE connection_id = $1 AND NOT (graph_user_id = ANY($2::text[]))`, [
    connectionId,
    currentIds,
  ]);
}

async function upsertConnectionUser(connectionId: string, row: ConnectionUserUpsert): Promise<void> {
  await query(
    `INSERT INTO connection_users
       (connection_id, graph_user_id, upn, display_name, storage_used_bytes, item_count, sync_status, error_message, last_synced_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, now())
     ON CONFLICT (connection_id, graph_user_id) DO UPDATE SET
       upn = EXCLUDED.upn,
       display_name = EXCLUDED.display_name,
       storage_used_bytes = EXCLUDED.storage_used_bytes,
       item_count = EXCLUDED.item_count,
       sync_status = EXCLUDED.sync_status,
       error_message = EXCLUDED.error_message,
       last_synced_at = now()`,
    [connectionId, row.graphUserId, row.upn, row.displayName, row.storageUsedBytes, row.itemCount, row.syncStatus, row.errorMessage]
  );
}

async function logConnectionEvent(
  event: string,
  connectionId: string,
  tenantId: string,
  detail: Record<string, unknown> = {}
): Promise<void> {
  await query(
    `INSERT INTO connection_events (connection_id, tenant_id, event, detail) VALUES ($1, $2, $3, $4)`,
    [connectionId, tenantId, event, detail]
  );
}

/** 401s that specifically indicate consent was revoked / the delegated context is gone, not a transient per-item failure. */
function isReauthError(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  const code = String((err as { code?: string; body?: string })?.code ?? (err as { body?: string })?.body ?? "");
  return status === 401 || /InvalidAuthenticationToken|consent_required|invalid_grant|AuthenticationError/i.test(code);
}

async function isCancelled(syncJobId: string): Promise<boolean> {
  const result = await query<{ cancel_requested_at: string | null }>(
    `SELECT cancel_requested_at FROM sync_jobs WHERE id = $1`,
    [syncJobId]
  );
  return result.rows[0]?.cancel_requested_at != null;
}

async function syncOneDrive(client: Awaited<ReturnType<typeof graphClientForTenant>>, connectionId: string, syncJobId: string, users: BasicUser[]): Promise<number> {
  let failed = 0;
  await runThrottled(users, (user) => getUserDriveQuota(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    // A single-object drive-quota GET is one of the lightest calls this app makes to Graph, and
    // unlike Teams chat listing it isn't known to be throttled more aggressively tenant-wide — doubling
    // the default concurrency roughly halves wall-clock time for large tenants (hundreds of accounts).
    // Safe to raise further if 429s stay rare; the existing Retry-After backoff absorbs it either way.
    batchSize: 40,
    onItemSettled: async (user, result) => {
      // A user with no OneDrive provisioned (getUserDriveQuota returns null on 404) counts as
      // "not added", same as a real error — not a silent success with 0 bytes. This is what backs
      // the Manage Clouds "Users Not Added" figure (see docs/cloud-connections-api.md).
      if (result.ok && result.value !== null) {
        const quota = result.value;
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: quota.usedBytes,
          itemCount: quota.itemCount,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage: result.ok ? "No OneDrive provisioned for this user" : String(result.error),
        });
      }
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });
  return failed;
}

async function syncTeams(client: Awaited<ReturnType<typeof graphClientForTenant>>, connectionId: string, syncJobId: string, users: BasicUser[]): Promise<number> {
  let failed = 0;
  await runThrottled(users, (user) => getUserJoinedTeamsCount(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    onItemSettled: async (user, result) => {
      if (result.ok) {
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0, // not meaningful for Teams — see docs/graph-api-limitations.md
          itemCount: result.value,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage: String(result.error),
        });
      }
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });
  return failed;
}

/** Each `connection_users` row is a SITE, not a person, for SharePoint — see docs/graph-api-limitations.md. */
async function syncSharePoint(client: Awaited<ReturnType<typeof graphClientForTenant>>, connectionId: string, syncJobId: string, sites: SiteSummary[]): Promise<number> {
  let failed = 0;
  await runThrottled(sites, (site) => getSiteDriveQuota(client, site.id), {
    isCancelled: () => isCancelled(syncJobId),
    // Same reasoning as syncOneDrive's batchSize bump — this is often the largest resource count
    // (thousands of sites), so it benefits the most from higher concurrency.
    batchSize: 40,
    onItemSettled: async (site, result) => {
      if (result.ok && result.value !== null) {
        const quota = result.value;
        await upsertConnectionUser(connectionId, {
          graphUserId: site.id,
          upn: site.webUrl,
          displayName: site.displayName,
          storageUsedBytes: quota.usedBytes,
          itemCount: quota.itemCount,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        await upsertConnectionUser(connectionId, {
          graphUserId: site.id,
          upn: site.webUrl,
          displayName: site.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage: result.ok ? "No document library provisioned for this site" : String(result.error),
        });
      }
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });
  return failed;
}

export const cloudSyncWorker = new Worker(
  "cloud-sync-jobs",
  async (job) => {
    const { syncJobId } = job.data as { syncJobId: string };

    const row = await query<{
      connection_id: string;
      tenant_id: string;
      cloud_type: CloudType;
      m365_tenant_id: string;
    }>(
      `SELECT c.id AS connection_id, c.tenant_id, c.cloud_type, t.m365_tenant_id
       FROM sync_jobs sj
       JOIN connections c ON c.id = sj.connection_id
       JOIN tenants t ON t.id = c.tenant_id
       WHERE sj.id = $1`,
      [syncJobId]
    );
    const info = row.rows[0];
    if (!info) throw new Error(`sync_jobs ${syncJobId} not found`);

    // See jobs/cleaningScanWorker.ts's identical comment: processed_users must reset here too, or a
    // BullMQ stalled-job redelivery (previous attempt's worker died mid-run) keeps incrementing on
    // top of the dead attempt's count against a freshly-set total_users, and the sync never visibly
    // reaches 100%.
    await query(`UPDATE sync_jobs SET status = 'running', started_at = now(), processed_users = 0 WHERE id = $1`, [syncJobId]);
    await logConnectionEvent("job_started", info.connection_id, info.tenant_id, { syncJobId, cloudType: info.cloud_type });

    try {
      const client = await graphClientForTenant(info.m365_tenant_id);
      let failed = 0;

      if (info.cloud_type === "sharepoint") {
        const sites = await searchSites(client);
        await pruneStaleConnectionUsers(info.connection_id, sites.map((s) => s.id));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, sites.length]);
        failed = await syncSharePoint(client, info.connection_id, syncJobId, sites);
      } else {
        const users = await listAllUsers(client);
        await pruneStaleConnectionUsers(info.connection_id, users.map((u) => u.id));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, users.length]);
        failed = info.cloud_type === "onedrive"
          ? await syncOneDrive(client, info.connection_id, syncJobId, users)
          : await syncTeams(client, info.connection_id, syncJobId, users);
      }

      const cancelled = await isCancelled(syncJobId);
      const finalStatus = cancelled ? "cancelled" : failed > 0 ? "completed_with_errors" : "completed";

      await query(`UPDATE sync_jobs SET status = $2, finished_at = now() WHERE id = $1`, [syncJobId, finalStatus]);
      if (!cancelled) {
        await query(
          `UPDATE connections
           SET last_synced_at = now(),
               status = CASE WHEN status IN ('connecting', 'error') THEN 'active' ELSE status END,
               last_error = NULL
           WHERE id = $1`,
          [info.connection_id]
        );
      }
      await logConnectionEvent("job_finished", info.connection_id, info.tenant_id, { syncJobId, status: finalStatus, failed });
    } catch (err) {
      // A tenant-wide auth failure (revoked consent, expired app-only grant) is a distinct
      // connection state, not a stalled progress bar — requirement #5 of the connections spec.
      if (isReauthError(err)) {
        await query(
          `UPDATE connections SET status = 'needs_reauth', last_error = $2 WHERE id = $1`,
          [info.connection_id, String(err)]
        );
        await logConnectionEvent("reauth_required", info.connection_id, info.tenant_id, { syncJobId, error: String(err) });
      } else {
        await query(`UPDATE connections SET status = 'error', last_error = $2 WHERE id = $1`, [info.connection_id, String(err)]);
      }
      await query(
        `UPDATE sync_jobs SET status = 'failed', finished_at = now(), error_log = error_log || $2::jsonb WHERE id = $1`,
        [syncJobId, JSON.stringify([{ message: String(err), at: new Date().toISOString() }])]
      );
      throw err;
    }
  },
  { connection: redis, concurrency: 5 }
);
