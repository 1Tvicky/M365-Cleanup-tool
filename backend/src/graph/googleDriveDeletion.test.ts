import { describe, expect, it, vi } from "vitest";
import type { drive_v3 } from "googleapis";
import { classifyGoogleDeleteError, deleteMyDriveItem } from "./googleDriveDeletion.js";

function fakeDrive(opts: { deleteImpl?: (args: unknown) => Promise<unknown>; updateImpl?: (args: unknown) => Promise<unknown> }) {
  const deleteFn = vi.fn(opts.deleteImpl ?? (() => Promise.resolve({})));
  const updateFn = vi.fn(opts.updateImpl ?? (() => Promise.resolve({})));
  const drive = { files: { delete: deleteFn, update: updateFn } } as unknown as drive_v3.Drive;
  return { drive, deleteFn, updateFn };
}

describe("deleteMyDriveItem", () => {
  it("permanent=true calls files.delete, not files.update", async () => {
    const { drive, deleteFn, updateFn } = fakeDrive({});
    const result = await deleteMyDriveItem(drive, "file-1", true);
    expect(result).toBe("deleted");
    expect(deleteFn).toHaveBeenCalledWith({ fileId: "file-1" });
    expect(updateFn).not.toHaveBeenCalled();
  });

  it("permanent=false calls files.update with trashed:true, not files.delete", async () => {
    const { drive, deleteFn, updateFn } = fakeDrive({});
    const result = await deleteMyDriveItem(drive, "file-1", false);
    expect(result).toBe("deleted");
    expect(updateFn).toHaveBeenCalledWith({ fileId: "file-1", requestBody: { trashed: true } });
    expect(deleteFn).not.toHaveBeenCalled();
  });

  it("treats a 404 as already_gone for both permanent and recycle-bin mode", async () => {
    const notFound = Object.assign(new Error("not found"), { code: 404 });
    const { drive: permanentDrive } = fakeDrive({ deleteImpl: () => Promise.reject(notFound) });
    expect(await deleteMyDriveItem(permanentDrive, "file-1", true)).toBe("already_gone");

    const { drive: softDrive } = fakeDrive({ updateImpl: () => Promise.reject(notFound) });
    expect(await deleteMyDriveItem(softDrive, "file-1", false)).toBe("already_gone");
  });

  it("propagates a non-404 error unchanged", async () => {
    const forbidden = Object.assign(new Error("forbidden"), { code: 403 });
    const { drive } = fakeDrive({ deleteImpl: () => Promise.reject(forbidden) });
    await expect(deleteMyDriveItem(drive, "file-1", true)).rejects.toThrow("forbidden");
  });
});

describe("classifyGoogleDeleteError", () => {
  it("classifies 403 as an insufficient-permission failure", () => {
    const result = classifyGoogleDeleteError({ code: 403 });
    expect(result.code).toBe("INSUFFICIENT_PERMISSION");
    expect(result.message).toMatch(/permission/i);
  });

  it("classifies 429 as retryable rate limiting", () => {
    const result = classifyGoogleDeleteError({ code: 429 });
    expect(result.code).toBe("RATE_LIMITED");
    expect(result.message).toMatch(/retried/i);
  });

  it("falls back to the reason/code and message for anything else", () => {
    const result = classifyGoogleDeleteError({ code: 500, message: "Internal error", errors: [{ reason: "backendError" }] });
    expect(result.code).toBe("backendError");
    expect(result.message).toBe("Internal error");
  });

  it("handles an error with no code at all", () => {
    const result = classifyGoogleDeleteError(new Error("network failure"));
    expect(result.code).toBe("UNKNOWN");
    expect(result.message).toBe("network failure");
  });
});
