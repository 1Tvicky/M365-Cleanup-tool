import { describe, expect, it, vi } from "vitest";
import type { drive_v3 } from "googleapis";
import { getSharedDriveById, getSharedDriveUsage, listAllSharedDrives, listSharedDriveRootItems } from "./googleSharedDriveEnumeration.js";

describe("listAllSharedDrives", () => {
  it("paginates with useDomainAdminAccess:true so it returns every drive in the domain, not just the caller's", async () => {
    const list = vi
      .fn()
      .mockResolvedValueOnce({ data: { drives: [{ id: "1", name: "Marketing" }], nextPageToken: "p2" } })
      .mockResolvedValueOnce({ data: { drives: [{ id: "2", name: "Sales" }] } });
    const drive = { drives: { list } } as unknown as drive_v3.Drive;

    const drives = await listAllSharedDrives(drive);

    expect(drives).toEqual([
      { id: "1", name: "Marketing" },
      { id: "2", name: "Sales" },
    ]);
    expect(list).toHaveBeenNthCalledWith(1, { useDomainAdminAccess: true, pageSize: 100, pageToken: undefined });
    expect(list).toHaveBeenNthCalledWith(2, { useDomainAdminAccess: true, pageSize: 100, pageToken: "p2" });
  });
});

describe("getSharedDriveById", () => {
  it("returns null on a not-found-shaped error", async () => {
    const get = vi.fn().mockRejectedValue({ code: 404 });
    const drive = { drives: { get } } as unknown as drive_v3.Drive;
    expect(await getSharedDriveById(drive, "missing")).toBeNull();
  });
});

describe("listSharedDriveRootItems", () => {
  it("scopes the query to the drive's own id as parent, with corpora/driveId/supportsAllDrives set", async () => {
    const list = vi.fn().mockResolvedValueOnce({ data: { files: [{ id: "f1", name: "doc.txt", size: "100" }] } });
    const drive = { files: { list } } as unknown as drive_v3.Drive;

    const items = await listSharedDriveRootItems(drive, "drive-1");

    expect(items).toEqual([{ id: "f1", name: "doc.txt", sizeBytes: 100 }]);
    expect(list).toHaveBeenCalledWith(
      expect.objectContaining({
        q: "'drive-1' in parents and trashed = false",
        corpora: "drive",
        driveId: "drive-1",
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
      })
    );
  });
});

describe("getSharedDriveUsage", () => {
  it("sums root item sizes as a best-effort, non-recursive total", async () => {
    const list = vi.fn().mockResolvedValueOnce({ data: { files: [{ id: "1", size: "10" }, { id: "2", size: "20" }] } });
    const drive = { files: { list } } as unknown as drive_v3.Drive;
    expect(await getSharedDriveUsage(drive, "drive-1")).toEqual({ itemCount: 2, usedBytes: 30 });
  });
});
