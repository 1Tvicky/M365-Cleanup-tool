import type { drive_v3 } from "googleapis";

/**
 * Google Drive calls specific to the Cleaning module's cleanup/deletion execution phase — the
 * Google equivalent of graph/cleanupDeletion.ts. Only My Drive is implemented here; Shared Drives
 * would need `supportsAllDrives: true` plus an `organizer` role check and are out of scope for this
 * pass (see docs/google-workspace-integration.md).
 *
 * Verified against the current Drive API v3 docs: `files.delete` "permanently deletes a file owned
 * by the user without moving it to the trash" — a genuine, immediate permanent delete, the direct
 * analog of Graph's permanentDelete. `files.update({trashed: true})` is the recycle-bin-equivalent
 * soft delete (30-day auto-purge, restorable). Which one runs is chosen per cleanup operation on the
 * confirmation screen (cleanup_operations.deletion_mode), same as the Graph side — never hardcoded.
 */

export type DriveItemDeleteResult = "deleted" | "already_gone";

function isNotFoundError(err: unknown): boolean {
  return (err as { code?: number })?.code === 404;
}

/**
 * A 404 means the file is already gone under this id — never touched, separately deleted, or
 * removed by a prior partial run — treated as success so retries stay idempotent, same convention
 * as deleteDriveItem in cleanupDeletion.ts.
 */
export async function deleteMyDriveItem(drive: drive_v3.Drive, fileId: string, permanent: boolean): Promise<DriveItemDeleteResult> {
  try {
    if (permanent) {
      await drive.files.delete({ fileId });
      console.log(`[cleanup.permanent_delete] completed resourceType=google_my_drive_account`);
    } else {
      await drive.files.update({ fileId, requestBody: { trashed: true } });
    }
    return "deleted";
  } catch (err) {
    if (isNotFoundError(err)) return "already_gone";
    if (permanent) console.warn(`[cleanup.permanent_delete] failed resourceType=google_my_drive_account`);
    throw err;
  }
}

/**
 * Google's client library surfaces errors with a numeric `.code` and `.errors[0].reason` rather
 * than Graph's `.statusCode` — a distinct mapping, not a shared function with classifyDeleteError.
 */
export function classifyGoogleDeleteError(err: unknown): { code: string; message: string } {
  const code = (err as { code?: number })?.code;
  const reason = (err as { errors?: { reason?: string }[] })?.errors?.[0]?.reason;
  if (code === 403) {
    return { code: "INSUFFICIENT_PERMISSION", message: "This connection doesn't currently have permission to delete this data." };
  }
  if (code === 429) {
    return { code: "RATE_LIMITED", message: "Google rate-limited this request. It can be retried." };
  }
  return { code: reason ?? String(code ?? "UNKNOWN"), message: String((err as { message?: string })?.message ?? err) };
}
