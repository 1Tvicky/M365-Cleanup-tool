import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import {
  countChannelMessages,
  countChatMessages,
  listAllTeams,
  listAllUsers,
  listChannels,
  listUserChats,
  type ChatSummary,
} from "../graph/cloudEnumeration.js";
import { runThrottled } from "../services/rateLimiter.js";
import { connection as redis } from "./queue.js";

/**
 * Mirrors jobs/cloudSyncWorker.ts's structure exactly (same join pattern, same runThrottled
 * usage, same reauth-error branching, same connection_events logging) for the Cleaning module's
 * two scan types:
 *  - teams_structure: discover actual teams/channels + every 1:1/group chat (cheap — no message
 *    bodies), so the Cleaning UI has something to show immediately.
 *  - message_counts: the expensive part — Graph has no count endpoint, so this pages through every
 *    channel/chat's messages. Runs separately, on demand, never blocking structure discovery.
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

/** ChannelMessage.Read.All missing/not consented yet — distinct from a transient failure so the UI can say "not available" instead of "failed", per docs/azure-ad-app-registration.md. */
function isMissingPermissionError(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  const code = String((err as { code?: string; body?: string })?.code ?? (err as { body?: string })?.body ?? "");
  return status === 403 || /Authorization_RequestDenied|AccessDenied/i.test(code);
}

async function isCancelled(scanId: string): Promise<boolean> {
  const result = await query<{ cancel_requested_at: string | null }>(`SELECT cancel_requested_at FROM cleaning_scans WHERE id = $1`, [scanId]);
  return result.rows[0]?.cancel_requested_at != null;
}

async function runStructureScan(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  scanId: string
): Promise<number> {
  let failed = 0;
  const teams = await listAllTeams(client);
  const users = await listAllUsers(client);
  await query(`UPDATE cleaning_scans SET total_items = $2 WHERE id = $1`, [scanId, teams.length + users.length]);

  const channelIds: string[] = [];
  await runThrottled(teams, (team) => listChannels(client, team.id), {
    isCancelled: () => isCancelled(scanId),
    onItemSettled: async (team, result) => {
      if (result.ok) {
        for (const channel of result.value) {
          channelIds.push(channel.id);
          await query(
            `INSERT INTO cleaning_channels (connection_id, team_id, team_name, channel_id, channel_name)
             VALUES ($1, $2, $3, $4, $5)
             ON CONFLICT (connection_id, channel_id) DO UPDATE SET team_name = EXCLUDED.team_name, channel_name = EXCLUDED.channel_name`,
            [connectionId, team.id, team.displayName, channel.id, channel.displayName]
          );
        }
      } else {
        failed++;
      }
      await query(`UPDATE cleaning_scans SET processed_items = processed_items + 1 WHERE id = $1`, [scanId]);
    },
  });

  const chatsById = new Map<string, ChatSummary>();
  await runThrottled(users, (user) => listUserChats(client, user.id), {
    isCancelled: () => isCancelled(scanId),
    // Teams chat listing throttles far more aggressively than most Graph resources (observed:
    // roughly 10 requests/10s tenant-wide) — a smaller batch means fewer wasted, immediately-
    // throttled attempts, not a change to the total number of users still to enumerate.
    batchSize: 5,
    onItemSettled: async (_user, result) => {
      if (result.ok) {
        for (const chat of result.value) chatsById.set(chat.id, chat);
      } else {
        failed++;
      }
      await query(`UPDATE cleaning_scans SET processed_items = processed_items + 1 WHERE id = $1`, [scanId]);
    },
  });

  const chatIds = [...chatsById.keys()];
  for (const chat of chatsById.values()) {
    await query(
      `INSERT INTO cleaning_chats (connection_id, chat_id, chat_type, participants, last_message_at)
       VALUES ($1, $2, $3, $4, $5)
       ON CONFLICT (connection_id, chat_id) DO UPDATE SET
         chat_type = EXCLUDED.chat_type, participants = EXCLUDED.participants, last_message_at = EXCLUDED.last_message_at`,
      [connectionId, chat.id, chat.chatType, JSON.stringify(chat.participants), chat.lastUpdatedDateTime]
    );
  }

  // Anything no longer present (channel deleted, chat no longer visible to any enumerated user) shouldn't linger.
  await query(`DELETE FROM cleaning_channels WHERE connection_id = $1 AND NOT (channel_id = ANY($2::text[]))`, [connectionId, channelIds]);
  await query(`DELETE FROM cleaning_chats WHERE connection_id = $1 AND NOT (chat_id = ANY($2::text[]))`, [connectionId, chatIds]);

  return failed;
}

async function runMessageCountScan(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  connectionId: string,
  scanId: string
): Promise<number> {
  let failed = 0;

  const pendingChannels = await query<{ id: string; team_id: string; channel_id: string }>(
    `SELECT id, team_id, channel_id FROM cleaning_channels WHERE connection_id = $1 AND count_status IN ('pending', 'failed')`,
    [connectionId]
  );
  const pendingChats = await query<{ id: string; chat_id: string }>(
    `SELECT id, chat_id FROM cleaning_chats WHERE connection_id = $1 AND count_status IN ('pending', 'failed')`,
    [connectionId]
  );

  await query(`UPDATE cleaning_scans SET total_items = $2 WHERE id = $1`, [scanId, pendingChannels.rows.length + pendingChats.rows.length]);
  await query(`UPDATE cleaning_channels SET count_status = 'calculating' WHERE id = ANY($1::uuid[])`, [pendingChannels.rows.map((r) => r.id)]);
  await query(`UPDATE cleaning_chats SET count_status = 'calculating' WHERE id = ANY($1::uuid[])`, [pendingChats.rows.map((r) => r.id)]);

  await runThrottled(pendingChannels.rows, (row) => countChannelMessages(client, row.team_id, row.channel_id), {
    isCancelled: () => isCancelled(scanId),
    batchSize: 10, // paginating full message+reply history per channel is heavier than a plain list call
    onItemSettled: async (row, result) => {
      if (result.ok) {
        await query(
          `UPDATE cleaning_channels SET message_count = $2, count_status = 'completed', error_message = NULL, last_counted_at = now() WHERE id = $1`,
          [row.id, result.value]
        );
      } else {
        failed++;
        const message = isMissingPermissionError(result.error)
          ? "Message counting isn't available yet for this connection — reconnect Teams after ChannelMessage.Read.All is granted."
          : String(result.error);
        await query(`UPDATE cleaning_channels SET count_status = 'failed', error_message = $2 WHERE id = $1`, [row.id, message]);
      }
      await query(`UPDATE cleaning_scans SET processed_items = processed_items + 1 WHERE id = $1`, [scanId]);
    },
  });

  await runThrottled(pendingChats.rows, (row) => countChatMessages(client, row.chat_id), {
    isCancelled: () => isCancelled(scanId),
    batchSize: 5, // same chats infrastructure as listUserChats — same tight throttle
    onItemSettled: async (row, result) => {
      if (result.ok) {
        await query(
          `UPDATE cleaning_chats SET message_count = $2, count_status = 'completed', error_message = NULL, last_counted_at = now() WHERE id = $1`,
          [row.id, result.value]
        );
      } else {
        failed++;
        await query(`UPDATE cleaning_chats SET count_status = 'failed', error_message = $2 WHERE id = $1`, [row.id, String(result.error)]);
      }
      await query(`UPDATE cleaning_scans SET processed_items = processed_items + 1 WHERE id = $1`, [scanId]);
    },
  });

  return failed;
}

export const cleaningScanWorker = new Worker(
  "cleaning-scan-jobs",
  async (job) => {
    const { scanId } = job.data as { scanId: string };

    const row = await query<{ connection_id: string; tenant_id: string; scan_type: "teams_structure" | "message_counts"; m365_tenant_id: string }>(
      `SELECT cs.connection_id, c.tenant_id, cs.scan_type, t.m365_tenant_id
       FROM cleaning_scans cs
       JOIN connections c ON c.id = cs.connection_id
       JOIN tenants t ON t.id = c.tenant_id
       WHERE cs.id = $1`,
      [scanId]
    );
    const info = row.rows[0];
    if (!info) throw new Error(`cleaning_scans ${scanId} not found`);

    await query(`UPDATE cleaning_scans SET status = 'running', started_at = now() WHERE id = $1`, [scanId]);
    await logConnectionEvent("cleaning_scan_started", info.connection_id, info.tenant_id, { scanId, scanType: info.scan_type });

    try {
      const client = await graphClientForTenant(info.m365_tenant_id);
      const failed =
        info.scan_type === "teams_structure"
          ? await runStructureScan(client, info.connection_id, scanId)
          : await runMessageCountScan(client, info.connection_id, scanId);

      const cancelled = await isCancelled(scanId);
      const finalStatus = cancelled ? "cancelled" : failed > 0 ? "completed_with_errors" : "completed";
      await query(`UPDATE cleaning_scans SET status = $2, finished_at = now() WHERE id = $1`, [scanId, finalStatus]);
      await logConnectionEvent("cleaning_scan_finished", info.connection_id, info.tenant_id, { scanId, status: finalStatus, failed });
    } catch (err) {
      if (isReauthError(err)) {
        await query(`UPDATE connections SET status = 'needs_reauth', last_error = $2 WHERE id = $1`, [info.connection_id, String(err)]);
      }
      await query(
        `UPDATE cleaning_scans SET status = 'failed', finished_at = now(), error_log = error_log || $2::jsonb WHERE id = $1`,
        [scanId, JSON.stringify([{ message: String(err), at: new Date().toISOString() }])]
      );
      throw err;
    }
  },
  { connection: redis, concurrency: 3 }
);
