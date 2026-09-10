import type { drive_v3 } from "googleapis";

/**
 * Google Shared Drive item deletion — the Google equivalent of graph/googleDriveDeletion.ts, with
 * supportsAllDrives:true added since files.update/files.delete otherwise silently fail to see
 * shared-drive content. Deliberately never deletes the shared drive itself (drives.delete is a
 * separate API call this module never makes) — default cleanup behavior is "delete contents,
 * preserve the drive," per the original spec.
 */

export type DriveItemDeleteResult = "deleted" | "already_gone";

function isNotFoundError(err: unknown): boolean {
  return (err as { code?: number })?.code === 404;
}

export async function deleteSharedDriveItem(drive: drive_v3.Drive, fileId: string, permanent: boolean): Promise<DriveItemDeleteResult> {
  try {
    if (permanent) {
      await drive.files.delete({ fileId, supportsAllDrives: true });
      console.log(`[cleanup.permanent_delete] completed resourceType=shared_drive`);
    } else {
      await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
    }
    return "deleted";
  } catch (err) {
    if (isNotFoundError(err)) return "already_gone";
    if (permanent) console.warn(`[cleanup.permanent_delete] failed resourceType=shared_drive`);
    throw err;
  }
}

/** Same mapping as classifyGoogleDeleteError in googleDriveDeletion.ts — duplicated rather than shared since these are deliberately separate, workload-isolated modules (see docs/google-workspace-integration.md). */
export function classifySharedDriveDeleteError(err: unknown): { code: string; message: string } {
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
