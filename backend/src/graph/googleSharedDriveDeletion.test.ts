import { describe, expect, it, vi } from "vitest";
import type { drive_v3 } from "googleapis";
import { classifySharedDriveDeleteError, deleteSharedDriveItem } from "./googleSharedDriveDeletion.js";

function fakeDrive(opts: { deleteImpl?: (args: unknown) => Promise<unknown>; updateImpl?: (args: unknown) => Promise<unknown> }) {
  const deleteFn = vi.fn(opts.deleteImpl ?? (() => Promise.resolve({})));
  const updateFn = vi.fn(opts.updateImpl ?? (() => Promise.resolve({})));
  const drive = { files: { delete: deleteFn, update: updateFn } } as unknown as drive_v3.Drive;
  return { drive, deleteFn, updateFn };
}

describe("deleteSharedDriveItem", () => {
  it("permanent=true calls files.delete with supportsAllDrives:true", async () => {
    const { drive, deleteFn, updateFn } = fakeDrive({});
    const result = await deleteSharedDriveItem(drive, "file-1", true);
    expect(result).toBe("deleted");
    expect(deleteFn).toHaveBeenCalledWith({ fileId: "file-1", supportsAllDrives: true });
    expect(updateFn).not.toHaveBeenCalled();
  });

  it("permanent=false calls files.update with trashed:true and supportsAllDrives:true", async () => {
    const { drive, deleteFn, updateFn } = fakeDrive({});
    const result = await deleteSharedDriveItem(drive, "file-1", false);
    expect(result).toBe("deleted");
    expect(updateFn).toHaveBeenCalledWith({ fileId: "file-1", supportsAllDrives: true, requestBody: { trashed: true } });
    expect(deleteFn).not.toHaveBeenCalled();
  });

  it("treats a 404 as already_gone", async () => {
    const notFound = Object.assign(new Error("not found"), { code: 404 });
    const { drive } = fakeDrive({ deleteImpl: () => Promise.reject(notFound) });
    expect(await deleteSharedDriveItem(drive, "file-1", true)).toBe("already_gone");
  });

  it("never calls drives.delete — this module has no way to delete the shared drive itself", async () => {
    const { drive } = fakeDrive({});
    expect((drive as unknown as { drives?: unknown }).drives).toBeUndefined();
  });
});

describe("classifySharedDriveDeleteError", () => {
  it("classifies 403 as an insufficient-permission failure", () => {
    expect(classifySharedDriveDeleteError({ code: 403 }).code).toBe("INSUFFICIENT_PERMISSION");
  });

  it("classifies 429 as retryable rate limiting", () => {
    expect(classifySharedDriveDeleteError({ code: 429 }).code).toBe("RATE_LIMITED");
  });
});
