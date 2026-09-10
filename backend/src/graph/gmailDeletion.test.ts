import { describe, expect, it, vi } from "vitest";
import type { gmail_v1 } from "googleapis";
import { deleteAllMailboxMessages } from "./gmailDeletion.js";

function fakeGmail(pages: { id: string }[][]) {
  let pageIndex = 0;
  const list = vi.fn(() => {
    const ids = pages[pageIndex] ?? [];
    pageIndex++;
    return Promise.resolve({ data: { messages: ids, nextPageToken: pageIndex < pages.length ? `page-${pageIndex}` : undefined } });
  });
  const batchDelete = vi.fn(() => Promise.resolve({}));
  const trash = vi.fn(() => Promise.resolve({}));
  const gmail = { users: { messages: { list, batchDelete, trash } } } as unknown as gmail_v1.Gmail;
  return { gmail, list, batchDelete, trash };
}

describe("deleteAllMailboxMessages", () => {
  it("never lists a second page before the first page's deletes are issued (memory-bounded streaming)", async () => {
    const calls: string[] = [];
    const gmail = {
      users: {
        messages: {
          list: vi.fn((args: { pageToken?: string }) => {
            calls.push(`list:${args.pageToken ?? "first"}`);
            const isFirst = !args.pageToken;
            return Promise.resolve({
              data: { messages: [{ id: isFirst ? "a" : "b" }], nextPageToken: isFirst ? "page-2" : undefined },
            });
          }),
          batchDelete: vi.fn((args: { requestBody: { ids: string[] } }) => {
            calls.push(`delete:${args.requestBody.ids.join(",")}`);
            return Promise.resolve({});
          }),
          trash: vi.fn(),
        },
      },
    } as unknown as gmail_v1.Gmail;

    await deleteAllMailboxMessages(gmail, "user@example.com", true);

    expect(calls).toEqual(["list:first", "delete:a", "list:page-2", "delete:b"]);
  });

  it("permanent=true uses batchDelete once per page, not one delete() call per message", async () => {
    const { gmail, batchDelete, trash } = fakeGmail([[{ id: "1" }, { id: "2" }, { id: "3" }]]);
    const result = await deleteAllMailboxMessages(gmail, "user@example.com", true);
    expect(batchDelete).toHaveBeenCalledTimes(1);
    expect(batchDelete).toHaveBeenCalledWith({ userId: "user@example.com", requestBody: { ids: ["1", "2", "3"] } });
    expect(trash).not.toHaveBeenCalled();
    expect(result).toEqual({ requested: 3, completed: 3, failed: 0, cancelled: false });
  });

  it("permanent=false calls trash() per message (no batch-trash exists)", async () => {
    const { gmail, batchDelete, trash } = fakeGmail([[{ id: "1" }, { id: "2" }]]);
    const result = await deleteAllMailboxMessages(gmail, "user@example.com", false);
    expect(trash).toHaveBeenCalledTimes(2);
    expect(trash).toHaveBeenCalledWith({ userId: "user@example.com", id: "1" });
    expect(batchDelete).not.toHaveBeenCalled();
    expect(result.completed).toBe(2);
  });

  it("stops pagination and reports cancelled when isCancelled becomes true", async () => {
    const { gmail, list } = fakeGmail([[{ id: "1" }], [{ id: "2" }]]);
    const result = await deleteAllMailboxMessages(gmail, "user@example.com", true, { isCancelled: () => true });
    expect(result.cancelled).toBe(true);
    expect(result.requested).toBe(0);
    expect(list).not.toHaveBeenCalled();
  });

  it("tracks a failed batchDelete chunk as failed, not thrown, so later pages still process", async () => {
    let call = 0;
    const gmail = {
      users: {
        messages: {
          list: vi.fn((args: { pageToken?: string }) => {
            const isFirst = !args.pageToken;
            return Promise.resolve({ data: { messages: [{ id: isFirst ? "a" : "b" }], nextPageToken: isFirst ? "p2" : undefined } });
          }),
          batchDelete: vi.fn(() => {
            call++;
            return call === 1 ? Promise.reject(new Error("boom")) : Promise.resolve({});
          }),
          trash: vi.fn(),
        },
      },
    } as unknown as gmail_v1.Gmail;

    const result = await deleteAllMailboxMessages(gmail, "user@example.com", true);
    expect(result).toEqual({ requested: 2, completed: 1, failed: 1, cancelled: false });
  });
});
