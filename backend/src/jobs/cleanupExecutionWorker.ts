import { Worker } from "bullmq";
import { query } from "../db/pool.js";
import { graphClientForTenant } from "../graph/client.js";
import {
  classifyDeleteError,
  deleteCalendarEvent,
  deleteContact,
  deleteDriveItem,
  deleteMessage,
  listCalendarEvents,
  listDriveRootChildren,
  listFolderContacts,
  listFolderMessages,
  listMailFoldersRecursive,
  listUserCalendars,
  listUserContactFoldersRecursive,
  type DriveOwnerKind,
} from "../graph/cleanupDeletion.js";
import { runThrottled } from "../services/rateLimiter.js";
import { connection as redis } from "./queue.js";
import { classifyGoogleDeleteError } from "../graph/googleDriveDeletion.js";
import { classifySharedDriveDeleteError } from "../graph/googleSharedDriveDeletion.js";
import { executeGoogleMyDriveItem, isGoogleReauthCleanupError } from "./googleDriveCleanupExecution.js";
import { executeSharedDriveItem } from "./googleSharedDriveCleanupExecution.js";

/**
 * Mirrors jobs/cleaningScanWorker.ts's shape (same join-by-id pattern, same runThrottled usage,
 * same reauth-error branching, same connection_events logging), but for executing a confirmed
 * cleanup operation. Only 'onedrive_account'/'sharepoint_site'/'outlook_mailbox' items are ever
 * processed here — 'channel'/'chat' items are created with status='unsupported' at manifest time
 * (routes/cleaning.ts) and never selected by the `status = 'pending'` query below, so this worker
 * never attempts a Graph call Microsoft doesn't support for this app's application-only permissions.
 *
 * All items in one operation share the same tenant (enforced at manifest-creation time), so only
 * one graphClientForTenant call is needed for the whole job, regardless of how many of the tenant's
 * up-to-4 connections (OneDrive/SharePoint/Teams/Outlook) are actually touched.
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

export interface PendingItem {
  id: string;
  connection_id: string;
  resource_type:
    | "onedrive_account"
    | "sharepoint_site"
    | "outlook_mailbox"
    | "outlook_calendar"
    | "outlook_contacts"
    | "google_my_drive_account"
    | "shared_drive";
  graph_ref: { userId?: string; siteId?: string; userEmail?: string; driveId?: string };
  /** This item's own connection's admin_upn — only populated/used for Google items where the impersonation subject isn't the item's own graph_ref (e.g. shared_drive, which has no owning user). */
  admin_upn?: string;
}

/** Deletes every top-level file/folder inside the account's OneDrive or the site's default document library. Never touches the account or site itself. */
async function executeItem(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  item: PendingItem,
  operationId: string,
  permanent: boolean
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
  await runThrottled(children, (child) => deleteDriveItem(client, kind, ownerId, child.id, permanent), {
    isCancelled: () => isCancelled(operationId),
    label: item.resource_type === "onedrive_account" ? "OneDrive" : "SharePoint",
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

/**
 * Deletes every message in every mail folder (all depths) for the mailbox — never the folders
 * themselves, since distinguished/well-known ones (Inbox, Sent Items, etc.) can't be deleted, only
 * emptied; see graph/cleanupDeletion.ts's comment on this. Same per-file (per-message) audit
 * pattern as executeItem above — one cleanup_operation_item_files row per message, keyed by
 * message.size for the "data cleared" report.
 */
async function executeMailboxItem(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  item: PendingItem,
  operationId: string,
  permanent: boolean
): Promise<void> {
  const userId = item.graph_ref.userId!;

  const folderIds = await listMailFoldersRecursive(client, userId);
  const messages: { folderId: string; message: { id: string; subject: string; size: number } }[] = [];
  for (const folderId of folderIds) {
    for (const message of await listFolderMessages(client, userId, folderId)) {
      messages.push({ folderId, message });
    }
  }
  if (messages.length === 0) return; // already empty — nothing to do, counts as success

  for (const { message } of messages) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, message.subject, message.id, message.size]
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, messages.length]);

  let firstError: unknown = null;
  await runThrottled(messages, ({ message }) => deleteMessage(client, userId, message.id, permanent), {
    isCancelled: () => isCancelled(operationId),
    label: "Outlook-Mail",
    batchSize: 10,
    onItemSettled: async ({ message }, result) => {
      const fileStatus = result.ok ? result.value : "failed"; // "deleted" | "already_gone" | "failed"
      const errorMessage = result.ok ? null : classifyDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, message.id, fileStatus, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}

/**
 * Deletes every event across every calendar the user owns — never the calendars themselves. Own
 * function, own Graph calls, deliberately not merged with executeMailboxItem — see the Outlook
 * isolation note in docs/azure-ad-app-registration.md.
 */
async function executeCalendarItem(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  item: PendingItem,
  operationId: string
): Promise<void> {
  const userId = item.graph_ref.userId!;

  const calendarIds = await listUserCalendars(client, userId);
  const events: { id: string; subject: string }[] = [];
  for (const calendarId of calendarIds) {
    events.push(...(await listCalendarEvents(client, userId, calendarId)));
  }
  if (events.length === 0) return; // already empty — nothing to do, counts as success

  for (const event of events) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, event.subject, event.id, 0] // events carry no reliable size field
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, events.length]);

  let firstError: unknown = null;
  await runThrottled(events, (event) => deleteCalendarEvent(client, userId, event.id), {
    isCancelled: () => isCancelled(operationId),
    label: "Outlook-Calendar",
    batchSize: 10,
    onItemSettled: async (event, result) => {
      const fileStatus = result.ok ? result.value : "failed";
      const errorMessage = result.ok ? null : classifyDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, event.id, fileStatus, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}

/**
 * Deletes every contact across every contact folder the user has (plus the root "Contacts"
 * collection) — never the folders themselves. Own function, own Graph calls, deliberately not
 * merged with executeMailboxItem/executeCalendarItem.
 */
async function executeContactItem(
  client: Awaited<ReturnType<typeof graphClientForTenant>>,
  item: PendingItem,
  operationId: string
): Promise<void> {
  const userId = item.graph_ref.userId!;

  const folderIds = await listUserContactFoldersRecursive(client, userId);
  const contacts: { id: string; displayName: string }[] = [];
  contacts.push(...(await listFolderContacts(client, userId, null))); // root "Contacts" collection
  for (const folderId of folderIds) {
    contacts.push(...(await listFolderContacts(client, userId, folderId)));
  }
  if (contacts.length === 0) return; // already empty — nothing to do, counts as success

  for (const contact of contacts) {
    await query(
      `INSERT INTO cleanup_operation_item_files (cleanup_operation_item_id, file_name, graph_item_id, file_size_bytes) VALUES ($1, $2, $3, $4)`,
      [item.id, contact.displayName, contact.id, 0] // contacts carry no reliable size field
    );
  }
  await query(`UPDATE cleanup_operation_items SET files_total = $2 WHERE id = $1`, [item.id, contacts.length]);

  let firstError: unknown = null;
  await runThrottled(contacts, (contact) => deleteContact(client, userId, contact.id), {
    isCancelled: () => isCancelled(operationId),
    label: "Outlook-Contacts",
    batchSize: 10,
    onItemSettled: async (contact, result) => {
      const fileStatus = result.ok ? result.value : "failed";
      const errorMessage = result.ok ? null : classifyDeleteError(result.error).message;
      await query(
        `UPDATE cleanup_operation_item_files SET status = $3, error_message = $4, completed_at = now()
         WHERE cleanup_operation_item_id = $1 AND graph_item_id = $2`,
        [item.id, contact.id, fileStatus, errorMessage]
      );
      await query(`UPDATE cleanup_operation_items SET files_completed = files_completed + 1 WHERE id = $1`, [item.id]);
      if (!result.ok && !firstError) firstError = result.error;
    },
  });
  if (firstError) throw firstError;
}

/**
 * Routes a pending item to the execution path for its resource type — executeItem only ever handles
 * onedrive_account/sharepoint_site, so Outlook's three resource kinds each get their own explicit
 * branch here rather than a shared switch-by-parameter helper. `permanent` is never consulted for
 * Calendar/Contacts — those two stay on plain soft delete regardless of the operation's
 * deletion_mode (out of scope for permanent deletion; see graph/cleanupDeletion.ts).
 */
function executeAnyItem(client: Awaited<ReturnType<typeof graphClientForTenant>>, item: PendingItem, operationId: string, permanent: boolean): Promise<void> {
  if (item.resource_type === "outlook_mailbox") return executeMailboxItem(client, item, operationId, permanent);
  if (item.resource_type === "outlook_calendar") return executeCalendarItem(client, item, operationId);
  if (item.resource_type === "outlook_contacts") return executeContactItem(client, item, operationId);
  return executeItem(client, item, operationId, permanent);
}

export const cleanupExecutionWorker = new Worker(
  "cleanup-execution-jobs",
  async (job) => {
    const { operationId } = job.data as { operationId: string };

    const opRow = await query<{ tenant_id: string; m365_tenant_id: string | null; google_customer_id: string | null; deletion_mode: "recycle_bin" | "permanent" }>(
      `SELECT co.tenant_id, t.m365_tenant_id, t.google_customer_id, co.deletion_mode
       FROM cleanup_operations co
       JOIN tenants t ON t.id = co.tenant_id
       WHERE co.id = $1`,
      [operationId]
    );
    const info = opRow.rows[0];
    if (!info) throw new Error(`cleanup_operations ${operationId} not found`);
    const permanent = info.deletion_mode === "permanent";

    const pending = await query<PendingItem>(
      `SELECT coi.id, coi.connection_id, coi.resource_type, coi.graph_ref, c.admin_upn
       FROM cleanup_operation_items coi
       JOIN connections c ON c.id = coi.connection_id
       WHERE coi.cleanup_operation_id = $1 AND coi.status = 'pending'`,
      [operationId]
    );

    // Same reasoning as jobs/cleaningScanWorker.ts's identical reset: a BullMQ stalled-job
    // redelivery must not keep incrementing processed_items on top of a dead attempt's count.
    await query(`UPDATE cleanup_operations SET status = 'running', started_at = now(), processed_items = 0 WHERE id = $1`, [operationId]);
    const touchedConnections = [...new Set(pending.rows.map((r) => r.connection_id))];
    for (const connectionId of touchedConnections) {
      await logConnectionEvent("cleanup_started", connectionId, info.tenant_id, { operationId });
    }

    const isGoogle = info.google_customer_id !== null;

    try {
      let successful = 0;
      let failed = 0;

      // Every item in one cleanup_operations row belongs to exactly one tenant, and every tenant is
      // exactly one vendor (M365 xor Google — see migrations/014_google_my_drive.sql) — but a
      // Google tenant can still mix google_my_drive_account and shared_drive items in one
      // operation (the same way an M365 tenant already mixes OneDrive/SharePoint/Outlook), so the
      // Google side dispatches per item's own resource_type, not once for the whole job. Neither
      // Google path shares a single client the way Graph does — each builds its own impersonated
      // client per item (My Drive: the item's own user; Shared Drives: the item's connection's
      // admin, since a Shared Drive has no owning user).
      let execute: (item: PendingItem) => Promise<void>;
      if (isGoogle) {
        execute = (item) =>
          item.resource_type === "shared_drive"
            ? executeSharedDriveItem(item.admin_upn!, item, operationId, permanent)
            : executeGoogleMyDriveItem(item, operationId, permanent);
      } else {
        const client = await graphClientForTenant(info.m365_tenant_id!);
        execute = (item) => executeAnyItem(client, item, operationId, permanent);
      }

      await runThrottled(pending.rows, execute, {
        isCancelled: () => isCancelled(operationId),
        batchSize: 3, // conservative — each item itself fans out into its own (throttled) per-file deletes; tune after the first live test run
        // executeAnyItem is a whole list-everything/seed-DB/delete-everything pipeline per item, not
        // one atomic Graph call — a mailbox with a few thousand messages finishes in seconds, but
        // one with millions of messages (observed: a real mailbox with 3.3M items) can never finish
        // inside the default 30s call timeout no matter how healthy Graph is, so every such item was
        // guaranteed to fail with a false "Graph call exceeded 30000ms". Every actual Graph call
        // inside executeAnyItem already gets its own correctly-scoped timeout via its own nested
        // runThrottled call (e.g. deleteMessage/deleteCalendarEvent/deleteContact below) — this
        // outer layer doesn't need one too, and imposing one here only forces large resources to
        // fail outright instead of just taking proportionally longer.
        callTimeoutMs: null,
        onItemSettled: async (item, result) => {
          if (result.ok) {
            successful++;
            await query(
              `UPDATE cleanup_operation_items SET status = 'completed', attempts = attempts + 1, completed_at = now(), updated_at = now() WHERE id = $1`,
              [item.id]
            );
          } else {
            failed++;
            const { code, message } = isGoogle
              ? item.resource_type === "shared_drive"
                ? classifySharedDriveDeleteError(result.error)
                : classifyGoogleDeleteError(result.error)
              : classifyDeleteError(result.error);
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
      if (isGoogle ? isGoogleReauthCleanupError(err) : isReauthError(err)) {
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
