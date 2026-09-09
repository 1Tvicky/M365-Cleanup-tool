import { describe, expect, it, vi } from "vitest";
import type { admin_directory_v1, drive_v3 } from "googleapis";
import { getGoogleUserById, getUserDriveUsage, listDomainUsers, listRootDriveItems } from "./googleDriveEnumeration.js";

describe("listDomainUsers", () => {
  it("paginates until nextPageToken is absent", async () => {
    const list = vi
      .fn()
      .mockResolvedValueOnce({ data: { users: [{ id: "1", primaryEmail: "a@x.com", name: { fullName: "A" } }], nextPageToken: "p2" } })
      .mockResolvedValueOnce({ data: { users: [{ id: "2", primaryEmail: "b@x.com", name: { fullName: "B" } }] } });
    const directory = { users: { list } } as unknown as admin_directory_v1.Admin;

    const users = await listDomainUsers(directory, "x.com");

    expect(users).toEqual([
      { id: "1", email: "a@x.com", displayName: "A" },
      { id: "2", email: "b@x.com", displayName: "B" },
    ]);
    expect(list).toHaveBeenCalledTimes(2);
    expect(list).toHaveBeenNthCalledWith(1, { domain: "x.com", maxResults: 500, pageToken: undefined });
    expect(list).toHaveBeenNthCalledWith(2, { domain: "x.com", maxResults: 500, pageToken: "p2" });
  });

  it("skips a user row missing id/primaryEmail rather than throwing", async () => {
    const list = vi.fn().mockResolvedValueOnce({ data: { users: [{ id: "1" }, { id: "2", primaryEmail: "b@x.com" }] } });
    const directory = { users: { list } } as unknown as admin_directory_v1.Admin;
    const users = await listDomainUsers(directory, "x.com");
    expect(users).toEqual([{ id: "2", email: "b@x.com", displayName: null }]);
  });
});

describe("getGoogleUserById", () => {
  it("returns null on a not-found-shaped error (404 or 400), never throws", async () => {
    const get = vi.fn().mockRejectedValue({ code: 404 });
    const directory = { users: { get } } as unknown as admin_directory_v1.Admin;
    expect(await getGoogleUserById(directory, "missing")).toBeNull();
  });

  it("propagates a non-not-found error", async () => {
    const get = vi.fn().mockRejectedValue({ code: 500 });
    const directory = { users: { get } } as unknown as admin_directory_v1.Admin;
    await expect(getGoogleUserById(directory, "user-1")).rejects.toBeTruthy();
  });
});

describe("getUserDriveUsage", () => {
  it("combines about.get's storageQuota with a paginated root-child count", async () => {
    const get = vi.fn().mockResolvedValue({ data: { storageQuota: { usage: "12345" } } });
    const list = vi
      .fn()
      .mockResolvedValueOnce({ data: { files: [{ id: "1" }, { id: "2" }], nextPageToken: "p2" } })
      .mockResolvedValueOnce({ data: { files: [{ id: "3" }] } });
    const drive = { about: { get }, files: { list } } as unknown as drive_v3.Drive;

    const usage = await getUserDriveUsage(drive);

    expect(usage).toEqual({ usedBytes: 12345, itemCount: 3 });
  });

  it("returns null when the user has no Drive provisioned (404)", async () => {
    const get = vi.fn().mockRejectedValue({ code: 404 });
    const drive = { about: { get }, files: { list: vi.fn() } } as unknown as drive_v3.Drive;
    expect(await getUserDriveUsage(drive)).toBeNull();
  });
});

describe("listRootDriveItems", () => {
  it("paginates and defaults missing size to 0 (native Google Docs/Sheets have no bytes)", async () => {
    const list = vi
      .fn()
      .mockResolvedValueOnce({ data: { files: [{ id: "1", name: "doc", size: undefined }], nextPageToken: "p2" } })
      .mockResolvedValueOnce({ data: { files: [{ id: "2", name: "file.pdf", size: "2048" }] } });
    const drive = { files: { list } } as unknown as drive_v3.Drive;

    const items = await listRootDriveItems(drive);

    expect(items).toEqual([
      { id: "1", name: "doc", sizeBytes: 0 },
      { id: "2", name: "file.pdf", sizeBytes: 2048 },
    ]);
  });
});
