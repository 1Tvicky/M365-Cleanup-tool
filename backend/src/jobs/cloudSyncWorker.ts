import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import {
  getSiteDriveQuota,
  getUserCalendarEventCount,
  getUserContactCount,
  getUserDriveQuota,
  getUserMailSummary,
  listAllTeams,
  listAllUsers,
  listChannels,
  searchSites,
  type BasicUser,
  type SiteSummary,
  type TeamSummary,
} from "../graph/cloudEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import { cloudProvider, type CloudType } from "../types/connections.js";
import { connection as redis } from "./queue.js";
import { isGoogleReauthError, listAllGoogleDomainUsers, syncGoogleMyDrive } from "./googleDriveSync.js";
import { listAllSharedDrivesForAdmin, syncSharedDrives } from "./googleSharedDriveSync.js";
import { listAllChatSpacesForAdmin, syncGoogleChatSpaces } from "./googleChatSync.js";
import { syncGmail } from "./gmailSync.js";

/**
 * Maps a resource's own Graph id (user id / site id / team id) to its sync_job_resources.id, for a
 * resource-scoped sync run. `null` means "this is a legacy/tenant-wide run" (no selection was made,
 * or this sync_jobs row predates this feature) — every sync* function below skips all
 * sync_job_resources bookkeeping in that case and behaves exactly as it always has.
 */
export type ResourceRowMap = Map<string, string> | null;

/** Bulk-marks every selected resource 'processing' right before its pass starts — batched per pass rather than per individual Graph call (which run concurrently within a batch anyway, see rateLimiter.ts), so this doesn't need to reach into runThrottled's per-item timing. Exported for jobs/googleDriveSync.ts, which mirrors this same resource-scoped-sync bookkeeping. */
export async function markProcessing(resourceRows: ResourceRowMap): Promise<void> {
  if (!resourceRows || resourceRows.size === 0) return;
  await query(`UPDATE sync_job_resources SET status = 'processing', started_at = now() WHERE id = ANY($1::uuid[])`, [[...resourceRows.values()]]);
}

/** Records one resource's outcome for a resource-scoped run — a no-op for a legacy/tenant-wide run (resourceRows is null) or a resource id with no matching row (shouldn't happen — every id here came from the same requested set the rows were seeded from). Exported — see markProcessing. */
export async function settleResource(resourceRows: ResourceRowMap, graphResourceId: string, ok: boolean, errorMessage: string | null): Promise<void> {
  const rowId = resourceRows?.get(graphResourceId);
  if (!rowId) return;
  await query(`UPDATE sync_job_resources SET status = $2, error_message = $3, completed_at = now(), updated_at = now() WHERE id = $1`, [
    rowId,
    ok ? "completed" : "failed",
    errorMessage,
  ]);
}

/** Bulk-flips any resources a cancelled run never got to (still 'pending' or 'processing' — a batch mid-flight when cancellation was noticed) to 'cancelled', mirroring how sync_jobs itself already records cancellation. No-op for a legacy/tenant-wide run. */
async function cancelRemainingResources(syncJobId: string): Promise<void> {
  await query(
    `UPDATE sync_job_resources SET status = 'cancelled', completed_at = now(), updated_at = now()
     WHERE sync_job_id = $1 AND status IN ('pending', 'processing')`,
    [syncJobId]
  );
}

export interface ConnectionUserUpsert {
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
export async function pruneStaleConnectionUsers(connectionId: string, currentIds: string[]): Promise<void> {
  await query(`DELETE FROM connection_users WHERE connection_id = $1 AND NOT (graph_user_id = ANY($2::text[]))`, [
    connectionId,
    currentIds,
  ]);
}

export async function upsertConnectionUser(connectionId: string, row: ConnectionUserUpsert): Promise<void> {
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

// Calendar and Contacts get their own prune/upsert pair, deliberately not a shared
// resource-parameterized helper — mirrors connection_users' pair but targets its own dedicated
// table (connection_outlook_calendars). See docs/azure-ad-app-registration.md's Outlook isolation
// note and migrations/011_outlook_calendar_contacts.sql's comment for why these tables aren't
// folded into connection_users.
async function pruneStaleOutlookCalendars(connectionId: string, currentIds: string[]): Promise<void> {
  await query(`DELETE FROM connection_outlook_calendars WHERE connection_id = $1 AND NOT (graph_user_id = ANY($2::text[]))`, [
    connectionId,
    currentIds,
  ]);
}

async function upsertOutlookCalendar(connectionId: string, row: ConnectionUserUpsert): Promise<void> {
  await query(
    `INSERT INTO connection_outlook_calendars
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

// Same pairing again, targeting connection_outlook_contacts — kept separate from the calendar pair
// above for the same reason: no function here branches on "which Outlook resource."
async function pruneStaleOutlookContacts(connectionId: string, currentIds: string[]): Promise<void> {
  await query(`DELETE FROM connection_outlook_contacts WHERE connection_id = $1 AND NOT (graph_user_id = ANY($2::text[]))`, [
    connectionId,
    currentIds,
  ]);
}

async function upsertOutlookContact(connectionId: string, row: ConnectionUserUpsert): Promise<void> {
  await query(
    `INSERT INTO connection_outlook_contacts
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

export async function logConnectionEvent(
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

export async function isCancelled(syncJobId: string): Promise<boolean> {
  const result = await query<{ cancel_requested_at: string | null }>(
    `SELECT cancel_requested_at FROM sync_jobs WHERE id = $1`,
    [syncJobId]
  );
  return result.rows[0]?.cancel_requested_at != null;
}

async function syncOneDrive(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  syncJobId: string,
  users: BasicUser[],
  resourceRows: ResourceRowMap
): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  await runThrottled(users, (user) => getUserDriveQuota(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "OneDrive",
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
        await settleResource(resourceRows, user.id, true, null);
      } else {
        failed++;
        const errorMessage = result.ok ? "No OneDrive provisioned for this user" : String(result.error);
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        await settleResource(resourceRows, user.id, false, errorMessage);
      }
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });
  return failed;
}

/**
 * Syncs actual Teams (not per-user joined-team counts, which is what this used to do — see
 * migrations/012_sync_job_resources.sql). "Sync a team" = enumerate its channels via the existing
 * listChannels (already used by the separate Cleaning module's structure scan) and store the count;
 * the selected team is the sync boundary, matching Outlook's mailbox-is-the-boundary and Cleanup's
 * own per-resource granularity elsewhere in this app.
 */
async function syncTeams(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  syncJobId: string,
  teams: TeamSummary[],
  resourceRows: ResourceRowMap
): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  await runThrottled(teams, (team) => listChannels(client, team.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "Teams",
    onItemSettled: async (team, result) => {
      if (result.ok) {
        await upsertConnectionUser(connectionId, {
          graphUserId: team.id,
          upn: team.displayName, // Teams has no natural secondary identifier — see connection_users' "columns mean different things per type" comment
          displayName: team.displayName,
          storageUsedBytes: 0, // not meaningful for Teams — see docs/graph-api-limitations.md
          itemCount: result.value.length,
          syncStatus: "synced",
          errorMessage: null,
        });
        await settleResource(resourceRows, team.id, true, null);
      } else {
        failed++;
        const errorMessage = String(result.error);
        await upsertConnectionUser(connectionId, {
          graphUserId: team.id,
          upn: team.displayName,
          displayName: team.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        await settleResource(resourceRows, team.id, false, errorMessage);
      }
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });
  return failed;
}

/**
 * Runs Mail, then Calendar, then Contacts as three separate, sequential runThrottled passes over
 * the same user list — never a single pass that Promise.all's all three Graph calls together per
 * item. Keeps them genuinely separate Graph operations (own label, own failure handling, own
 * table), the same way jobs/cleaningScanWorker.ts's runStructureScan already runs two separate
 * passes (teams→channels, users→chats) inside one job rather than interleaving them.
 *
 * Bundled into this one sync_jobs row (not three) per product decision: Outlook's existing "Sync
 * Now" stays one button. sync_jobs.total_users/processed_users are shared with routes/
 * cloudConnections.ts's GET /manage (Manage Clouds' "X out of Y Mailboxes"), which — like every
 * other cloud type — expects processed_users to reach total_users exactly once per real user, not
 * once per (user, phase). So total_users stays users.length (set by the caller, unchanged), and
 * only the last pass (Contacts) increments processed_users — Mail/Calendar still do their own full
 * work and DB writes, they just don't touch the shared counter, so it can't overshoot total_users
 * or corrupt Manage Clouds' math the way an earlier version of this function did (confirmed live:
 * it showed "0 out of 1,185 Mailboxes" instead of the real 395).
 */
async function syncOutlook(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  syncJobId: string,
  users: BasicUser[],
  resourceRows: ResourceRowMap
): Promise<number> {
  let failed = 0;
  // The selected mailbox is the sync boundary (matching Teams' selected-team boundary) even though
  // it internally runs three separate Graph passes (Mail, Calendar, Contacts) — so sync_job_resources
  // is marked processing once up front, and settled once at the end below based on whether ANY of
  // the three passes failed for that user, not once per pass.
  const failedAny = new Map<string, string>(); // userId -> first error message seen for that user, across all three passes
  await markProcessing(resourceRows);

  // connection_users is already pruned by the caller (shared with onedrive/teams) before syncOutlook runs.
  await runThrottled(users, (user) => getUserMailSummary(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "Outlook-Mail",
    // Same reasoning as syncOneDrive's batchSize bump — a mail-folder listing is a comparably light call.
    batchSize: 40,
    onItemSettled: async (user, result) => {
      if (result.ok && result.value !== null) {
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0, // not meaningful for Outlook — see graph/cloudEnumeration.ts's getUserMailSummary
          itemCount: result.value.itemCount,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        const errorMessage = result.ok ? "No mailbox provisioned for this user" : String(result.error);
        await upsertConnectionUser(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        if (!failedAny.has(user.id)) failedAny.set(user.id, errorMessage);
      }
      // Deliberately no processed_users increment here — see this function's doc comment.
    },
  });

  await pruneStaleOutlookCalendars(connectionId, users.map((u) => u.id));
  await runThrottled(users, (user) => getUserCalendarEventCount(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "Outlook-Calendar",
    // Unlike Mail's single mailFolders call, this is enumerate-calendars-then-paginate-events-per-
    // calendar — confirmed by live testing to throttle far more aggressively than a flat call (the
    // same lesson cleaningScanWorker.ts already learned for Teams chat listing, batchSize: 5) — 40
    // concurrent workers drove the sync into a near-standstill of repeated timeouts/429s.
    batchSize: 8,
    onItemSettled: async (user, result) => {
      if (result.ok && result.value !== null) {
        await upsertOutlookCalendar(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: result.value.itemCount,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        const errorMessage = result.ok ? "No mailbox provisioned for this user" : String(result.error);
        await upsertOutlookCalendar(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        if (!failedAny.has(user.id)) failedAny.set(user.id, errorMessage);
      }
      // Deliberately no processed_users increment here either — see this function's doc comment.
    },
  });

  await pruneStaleOutlookContacts(connectionId, users.map((u) => u.id));
  await runThrottled(users, (user) => getUserContactCount(client, user.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "Outlook-Contacts",
    // Same reasoning as the Calendar pass above — a folder-BFS-then-paginate-per-folder call is
    // heavier than Mail's flat call, so it gets the same conservative concurrency rather than
    // waiting to hit the same throttling wall Calendar just did.
    batchSize: 8,
    onItemSettled: async (user, result) => {
      if (result.ok && result.value !== null) {
        await upsertOutlookContact(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: result.value.itemCount,
          syncStatus: "synced",
          errorMessage: null,
        });
      } else {
        failed++;
        const errorMessage = result.ok ? "No mailbox provisioned for this user" : String(result.error);
        await upsertOutlookContact(connectionId, {
          graphUserId: user.id,
          upn: user.upn,
          displayName: user.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        if (!failedAny.has(user.id)) failedAny.set(user.id, errorMessage);
      }
      // The one place this function increments the shared counter — see this function's doc
      // comment for why Mail/Calendar deliberately don't.
      await query(`UPDATE sync_jobs SET processed_users = processed_users + 1 WHERE id = $1`, [syncJobId]);
    },
  });

  // Settle each selected mailbox once, here, based on whether ANY of the three passes above failed
  // for it — never once per pass, which would overwrite an earlier pass's outcome with a later one.
  for (const user of users) {
    await settleResource(resourceRows, user.id, !failedAny.has(user.id), failedAny.get(user.id) ?? null);
  }

  return failed;
}

/** Each `connection_users` row is a SITE, not a person, for SharePoint — see docs/graph-api-limitations.md. */
async function syncSharePoint(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  syncJobId: string,
  sites: SiteSummary[],
  resourceRows: ResourceRowMap
): Promise<number> {
  let failed = 0;
  await markProcessing(resourceRows);
  await runThrottled(sites, (site) => getSiteDriveQuota(client, site.id), {
    isCancelled: () => isCancelled(syncJobId),
    label: "SharePoint",
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
        await settleResource(resourceRows, site.id, true, null);
      } else {
        failed++;
        const errorMessage = result.ok ? "No document library provisioned for this site" : String(result.error);
        await upsertConnectionUser(connectionId, {
          graphUserId: site.id,
          upn: site.webUrl,
          displayName: site.displayName,
          storageUsedBytes: 0,
          itemCount: 0,
          syncStatus: "failed",
          errorMessage,
        });
        await settleResource(resourceRows, site.id, false, errorMessage);
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
      m365_tenant_id: string | null;
      admin_upn: string;
    }>(
      `SELECT c.id AS connection_id, c.tenant_id, c.cloud_type, t.m365_tenant_id, c.admin_upn
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

    // A resource-scoped run (POST /:id/resync with resourceIds) has already-validated rows here,
    // seeded by the route — an empty result means either no selection was made (today's plain
    // "Resync" icon) or this sync_jobs row predates this feature entirely; both cases run the exact
    // tenant-wide path below, unchanged, which is also what keeps routes/cleaning.ts's "Sync Now"
    // and the initial connect flow (neither ever populates sync_job_resources) working as before.
    const selectedResources = await query<{
      id: string;
      graph_resource_id: string;
      display_name: string;
      secondary: string | null;
    }>(`SELECT id, graph_resource_id, display_name, secondary FROM sync_job_resources WHERE sync_job_id = $1`, [syncJobId]);
    const isScoped = selectedResources.rows.length > 0;

    try {
      let failed = 0;

      if (cloudProvider(info.cloud_type) === "google") {
        // Google's per-item impersonated-client model (built inside each sync* function's own
        // runThrottled callback) means there's no single shared client to build here, unlike every
        // M365 branch below. Each Google cloud_type has its own resource shape/table, so this
        // dispatches on cloud_type once per branch rather than trying to force a shared shape.
        switch (info.cloud_type) {
          case "shared_drive": {
            if (isScoped) {
              const resourceRows: ResourceRowMap = new Map(selectedResources.rows.map((r) => [r.graph_resource_id, r.id]));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, selectedResources.rows.length]);
              const drives = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, name: r.display_name }));
              failed = await syncSharedDrives(info.admin_upn, info.connection_id, syncJobId, drives, resourceRows);
            } else {
              const drives = await listAllSharedDrivesForAdmin(info.admin_upn);
              await pruneStaleConnectionUsers(info.connection_id, drives.map((d) => d.id));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, drives.length]);
              failed = await syncSharedDrives(info.admin_upn, info.connection_id, syncJobId, drives, null);
            }
            break;
          }
          case "google_chat": {
            if (isScoped) {
              const resourceRows: ResourceRowMap = new Map(selectedResources.rows.map((r) => [r.graph_resource_id, r.id]));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, selectedResources.rows.length]);
              const spaces = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, displayName: r.display_name }));
              failed = await syncGoogleChatSpaces(info.admin_upn, info.connection_id, syncJobId, spaces, resourceRows);
            } else {
              const spaces = await listAllChatSpacesForAdmin(info.admin_upn);
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, spaces.length]);
              failed = await syncGoogleChatSpaces(info.admin_upn, info.connection_id, syncJobId, spaces, null);
            }
            break;
          }
          case "gmail": {
            if (isScoped) {
              const resourceRows: ResourceRowMap = new Map(selectedResources.rows.map((r) => [r.graph_resource_id, r.id]));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, selectedResources.rows.length]);
              const users = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, email: r.secondary ?? "", displayName: r.display_name }));
              failed = await syncGmail(info.connection_id, syncJobId, users, resourceRows);
            } else {
              const users = await listAllGoogleDomainUsers(info.admin_upn);
              await pruneStaleConnectionUsers(info.connection_id, users.map((u) => u.id));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, users.length]);
              failed = await syncGmail(info.connection_id, syncJobId, users, null);
            }
            break;
          }
          default: {
            // google_my_drive
            if (isScoped) {
              const resourceRows: ResourceRowMap = new Map(selectedResources.rows.map((r) => [r.graph_resource_id, r.id]));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, selectedResources.rows.length]);
              const users = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, email: r.secondary ?? "", displayName: r.display_name }));
              failed = await syncGoogleMyDrive(info.connection_id, syncJobId, users, resourceRows);
            } else {
              const users = await listAllGoogleDomainUsers(info.admin_upn);
              await pruneStaleConnectionUsers(info.connection_id, users.map((u) => u.id));
              await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, users.length]);
              failed = await syncGoogleMyDrive(info.connection_id, syncJobId, users, null);
            }
          }
        }

        const cancelled = await isCancelled(syncJobId);
        const finalStatus = cancelled ? "cancelled" : failed > 0 ? "completed_with_errors" : "completed";
        if (cancelled) await cancelRemainingResources(syncJobId);
        await query(`UPDATE sync_jobs SET status = $2, finished_at = now() WHERE id = $1`, [syncJobId, finalStatus]);
        if (!cancelled) {
          await query(
            `UPDATE connections SET last_synced_at = now(), status = CASE WHEN status IN ('connecting', 'error') THEN 'active' ELSE status END, last_error = NULL WHERE id = $1`,
            [info.connection_id]
          );
        }
        await logConnectionEvent("job_finished", info.connection_id, info.tenant_id, { syncJobId, status: finalStatus, failed });
        return;
      }

      const client = await graphClientForTenant(info.m365_tenant_id!);

      if (isScoped) {
        // Never re-enumerates the tenant — the selected resources' identifying fields were already
        // validated and snapshotted by the route at selection time, so this just reshapes them into
        // the same BasicUser/SiteSummary/TeamSummary the sync* functions already expect. No prune*
        // call here either: pruning means "remove anything not in this full listing," which is only
        // meaningful for a run that actually touched the full tenant.
        const resourceRows: ResourceRowMap = new Map(selectedResources.rows.map((r) => [r.graph_resource_id, r.id]));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, selectedResources.rows.length]);

        if (info.cloud_type === "sharepoint") {
          const sites: SiteSummary[] = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, displayName: r.display_name, webUrl: r.secondary ?? "" }));
          failed = await syncSharePoint(client, info.connection_id, syncJobId, sites, resourceRows);
        } else if (info.cloud_type === "teams") {
          const teams: TeamSummary[] = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, displayName: r.display_name }));
          failed = await syncTeams(client, info.connection_id, syncJobId, teams, resourceRows);
        } else {
          const users: BasicUser[] = selectedResources.rows.map((r) => ({ id: r.graph_resource_id, upn: r.secondary ?? "", displayName: r.display_name }));
          failed = info.cloud_type === "onedrive" ? await syncOneDrive(client, info.connection_id, syncJobId, users, resourceRows) : await syncOutlook(client, info.connection_id, syncJobId, users, resourceRows);
        }
      } else if (info.cloud_type === "sharepoint") {
        const sites = await searchSites(client);
        await pruneStaleConnectionUsers(info.connection_id, sites.map((s) => s.id));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, sites.length]);
        failed = await syncSharePoint(client, info.connection_id, syncJobId, sites, null);
      } else if (info.cloud_type === "teams") {
        const teams = await listAllTeams(client);
        await pruneStaleConnectionUsers(info.connection_id, teams.map((t) => t.id));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, teams.length]);
        failed = await syncTeams(client, info.connection_id, syncJobId, teams, null);
      } else {
        const users = await listAllUsers(client);
        await pruneStaleConnectionUsers(info.connection_id, users.map((u) => u.id));
        await query(`UPDATE sync_jobs SET total_users = $2 WHERE id = $1`, [syncJobId, users.length]);
        failed = info.cloud_type === "onedrive" ? await syncOneDrive(client, info.connection_id, syncJobId, users, null) : await syncOutlook(client, info.connection_id, syncJobId, users, null);
      }

      const cancelled = await isCancelled(syncJobId);
      const finalStatus = cancelled ? "cancelled" : failed > 0 ? "completed_with_errors" : "completed";
      if (cancelled) await cancelRemainingResources(syncJobId);

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
      // A tenant-wide auth failure (revoked consent, expired app-only grant / revoked domain-wide
      // delegation) is a distinct connection state, not a stalled progress bar — requirement #5 of
      // the connections spec. Google's error shape differs from Graph's, hence the separate check.
      if (cloudProvider(info.cloud_type) === "google" ? isGoogleReauthError(err) : isReauthError(err)) {
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
      // Whatever resources hadn't settled yet when the whole job crashed (auth failure, unexpected
      // exception) must not be left at 'pending'/'processing' forever — distinct from cancellation,
      // this reflects the job dying out from under them, not an operator-requested stop.
      await query(
        `UPDATE sync_job_resources SET status = 'failed', error_message = COALESCE(error_message, $2), completed_at = now(), updated_at = now()
         WHERE sync_job_id = $1 AND status IN ('pending', 'processing')`,
        [syncJobId, String(err)]
      );
      throw err;
    }
  },
  { connection: redis, concurrency: 5 }
);
