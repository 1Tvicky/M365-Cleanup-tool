import type { gmail_v1 } from "googleapis";
import { runThrottled } from "../services/rateLimiter.js";

/**
 * Gmail message deletion — deliberately shaped differently from every other Google/M365 delete
 * module in this codebase. Every other one does "list everything first, then delete the full
 * list" (fine for dozens/thousands of files/folders); Gmail mailboxes can hold millions of
 * messages, so listing everything first would mean holding millions of ids in memory before a
 * single delete happens. Instead, deleteAllMailboxMessages paginates and deletes in the SAME loop —
 * each page (up to 500 ids, Gmail's own list cap) is deleted before the next page is even
 * requested, so memory stays bounded regardless of mailbox size. Per the same reasoning, this
 * module tracks only aggregate counters, never one DB row per message — a million-message mailbox
 * would otherwise mean a million cleanup_operation_item_files rows for one cleanup item.
 *
 * Verified against the Gmail API reference (not guessed): messages.delete is a genuine immediate
 * permanent delete ("cannot be undone"); messages.trash is the recycle-bin-equivalent soft delete —
 * same recycle_bin/permanent mapping already used everywhere else in this app. messages.batchDelete
 * accepts up to 1,000 ids per call and is itself always permanent (skips trash) — used only in
 * permanent mode, since trash has no batch equivalent.
 */

export interface MailboxDeleteProgress {
  requested: number;
  completed: number;
  failed: number;
}

export interface MailboxDeleteResult extends MailboxDeleteProgress {
  cancelled: boolean;
}

const LIST_PAGE_SIZE = 500; // Gmail's own documented max for messages.list
const BATCH_DELETE_CHUNK_SIZE = 1000; // Gmail's own documented max ids per batchDelete call

/**
 * Deletes every message in the mailbox — never the mailbox, the user, or the account. `onProgress`
 * is called after each page settles so the caller can persist aggregate progress (files_total/
 * files_completed on the cleanup_operation_items row) without this module knowing about the DB.
 */
export async function deleteAllMailboxMessages(
  gmail: gmail_v1.Gmail,
  userEmail: string,
  permanent: boolean,
  opts: { isCancelled?: () => boolean | Promise<boolean>; onProgress?: (progress: MailboxDeleteProgress) => void | Promise<void> } = {}
): Promise<MailboxDeleteResult> {
  let requested = 0;
  let completed = 0;
  let failed = 0;
  let pageToken: string | undefined;
  let cancelled = false;

  do {
    if (await opts.isCancelled?.()) {
      cancelled = true;
      break;
    }

    const listRes = await gmail.users.messages.list({ userId: userEmail, maxResults: LIST_PAGE_SIZE, pageToken });
    const ids = (listRes.data.messages ?? []).map((m) => m.id!).filter(Boolean);
    pageToken = listRes.data.nextPageToken ?? undefined;
    if (ids.length === 0) continue;
    requested += ids.length;

    if (permanent) {
      // One batchDelete call per page (<=500 ids, under the 1000-id cap) — far fewer API calls
      // than one delete() per message, which matters at million-message scale.
      try {
        await gmail.users.messages.batchDelete({ userId: userEmail, requestBody: { ids } });
        completed += ids.length;
      } catch (err) {
        // batchDelete is all-or-nothing per call; a failed chunk means none of its ids are
        // confirmed deleted — retry-by-resync (re-running cleanup) picks up whatever's left,
        // since the next run's messages.list simply won't return anything already gone.
        failed += ids.length;
        console.warn(`[cleanup.permanent_delete] failed resourceType=gmail_mailbox batchSize=${ids.length}`);
      }
    } else {
      // No batch-trash exists — one trash() call per message, throttled like every other
      // per-item soft delete in this app.
      await runThrottled(ids, (id) => gmail.users.messages.trash({ userId: userEmail, id }), {
        isCancelled: opts.isCancelled,
        label: "Gmail",
        batchSize: 20,
        onItemSettled: (_id, result) => {
          if (result.ok) completed++;
          else failed++;
        },
      });
    }

    await opts.onProgress?.({ requested, completed, failed });
  } while (pageToken);

  return { requested, completed, failed, cancelled };
}
