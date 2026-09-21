import { describe, expect, it, vi, afterEach } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import { createFolder, uploadFile } from "./uploadHelper.js";

/** Builds a fake Graph Client whose `.api(path)` returns a fluent chain recording every call. */
function fakeClient(responses: { post?: unknown; put?: unknown; patch?: unknown; get?: unknown } = {}) {
  const calls: { method: string; path: string; body?: unknown; header?: Record<string, string> }[] = [];
  const api = (path: string) => {
    const headers: Record<string, string> = {};
    const chain = {
      header: (name: string, value: string) => {
        headers[name] = value;
        return chain;
      },
      post: vi.fn(async (body?: unknown) => {
        calls.push({ method: "POST", path, body, header: headers });
        return responses.post;
      }),
      put: vi.fn(async (body?: unknown) => {
        calls.push({ method: "PUT", path, body, header: headers });
        return responses.put;
      }),
      patch: vi.fn(async (body?: unknown) => {
        calls.push({ method: "PATCH", path, body, header: headers });
        return responses.patch;
      }),
      get: vi.fn(async () => {
        calls.push({ method: "GET", path, header: headers });
        return responses.get;
      }),
    };
    return chain;
  };
  return { client: { api } as unknown as Client, calls };
}

describe("createFolder", () => {
  it("posts to root/children when there is no parent", async () => {
    const { client, calls } = fakeClient({ post: { id: "f1", name: "Docs", webUrl: "https://x" } });
    const result = await createFolder(client, "/users/u1/drive", null, "Docs");
    expect(calls[0]!.path).toBe("/users/u1/drive/root/children");
    expect(result).toEqual({ id: "f1", name: "Docs", webUrl: "https://x" });
  });

  it("posts to items/{parentId}/children when a parent is given, with conflictBehavior=rename (never overwrite/duplicate silently)", async () => {
    const { client, calls } = fakeClient({ post: { id: "f2", name: "Sub", webUrl: "https://x" } });
    await createFolder(client, "/users/u1/drive", "parent-1", "Sub");
    expect(calls[0]!.path).toBe("/users/u1/drive/items/parent-1/children");
    expect((calls[0]!.body as any)["@microsoft.graph.conflictBehavior"]).toBe("rename");
  });
});

describe("uploadFile", () => {
  const smallContent = Buffer.alloc(1024, 1);
  const largeContent = Buffer.alloc(5 * 1024 * 1024, 2); // > 4 MiB Graph simple-upload threshold

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("uses a single PUT :content call for a file under Graph's 4 MiB simple-upload threshold", async () => {
    const { client, calls } = fakeClient({ put: { id: "file1", name: "a.txt", webUrl: "https://x" } });
    const result = await uploadFile(client, "/users/u1/drive", "parent-1", "a.txt", smallContent, "text/plain");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.method).toBe("PUT");
    expect(calls[0]!.path).toContain(":/content");
    expect(result.sizeBytes).toBe(smallContent.length);
  });

  it("uses a resumable upload session (never a single PUT of the whole buffer) for a file over the 4 MiB threshold", async () => {
    const fetchMock = vi.fn(async (_url: string, init: RequestInit) => {
      const range = (init.headers as Record<string, string>)["Content-Range"];
      const isLast = range?.endsWith(`/${largeContent.length}`) && range.split("-")[1]?.split("/")[0] === String(largeContent.length - 1);
      return {
        ok: true,
        status: isLast ? 201 : 202,
        json: async () => ({ id: "file2", name: "big.bin", webUrl: "https://x" }),
        text: async () => "",
      } as Response;
    });
    vi.stubGlobal("fetch", fetchMock);

    const { client, calls } = fakeClient({ post: { uploadUrl: "https://upload.example/session" } });
    const result = await uploadFile(client, "/users/u1/drive", "parent-1", "big.bin", largeContent, "application/octet-stream");

    expect(calls.some((c) => c.path.includes("createUploadSession"))).toBe(true);
    expect(calls.some((c) => c.method === "PUT" && c.path.includes(":/content"))).toBe(false); // never the small-file path
    expect(fetchMock.mock.calls.length).toBeGreaterThan(1); // chunked, not one request
    expect(result.sizeBytes).toBe(largeContent.length);
  });

  it("PATCHes fileSystemInfo when historical timestamps are provided (never mutates the just-uploaded content itself)", async () => {
    const { client, calls } = fakeClient({ put: { id: "file3", name: "a.txt", webUrl: "https://x" }, patch: { id: "file3" } });
    await uploadFile(client, "/users/u1/drive", "parent-1", "a.txt", smallContent, "text/plain", {
      createdDateTime: "2024-01-01T00:00:00Z",
      lastModifiedDateTime: "2024-01-01T00:00:00Z",
    });
    const patchCall = calls.find((c) => c.method === "PATCH");
    expect(patchCall).toBeDefined();
    expect((patchCall!.body as any).fileSystemInfo.createdDateTime).toBe("2024-01-01T00:00:00Z");
  });

  it("never calls PATCH when no fileSystemInfo is given", async () => {
    const { client, calls } = fakeClient({ put: { id: "file4", name: "a.txt", webUrl: "https://x" } });
    await uploadFile(client, "/users/u1/drive", "parent-1", "a.txt", smallContent, "text/plain");
    expect(calls.some((c) => c.method === "PATCH")).toBe(false);
  });
});
