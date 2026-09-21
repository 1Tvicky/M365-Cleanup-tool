import { describe, expect, it, vi } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import { ApiError } from "../../types/index.js";
import { addChannelMember, addTeamMember, createChannel, createTeam, postChannelMessageUnsupported } from "./teamsCreation.js";

function fakeClient(handlers: { post?: (path: string, body?: unknown) => Promise<unknown> }) {
  const calls: { path: string; body?: unknown }[] = [];
  const api = (path: string) => ({
    post: vi.fn(async (body?: unknown) => {
      calls.push({ path, body });
      if (!handlers.post) throw new Error(`no handler for ${path}`);
      return handlers.post(path, body);
    }),
    put: vi.fn(async (body?: unknown) => {
      calls.push({ path, body });
      return handlers.post ? handlers.post(path, body) : undefined;
    }),
  });
  return { client: { api } as unknown as Client, calls };
}

describe("postChannelMessageUnsupported", () => {
  it("always throws a clear 501 rather than silently no-opping (spec: never simulate an unsupported Graph operation as successful)", () => {
    expect(() => postChannelMessageUnsupported()).toThrow(ApiError);
    try {
      postChannelMessageUnsupported();
    } catch (err) {
      expect(err).toBeInstanceOf(ApiError);
      expect((err as ApiError).status).toBe(501);
      expect((err as ApiError).code).toBe("UNSUPPORTED_OPERATION");
    }
  });
});

describe("createTeam", () => {
  it("creates a group with exactly the given owner, then provisions a Team on top of it", async () => {
    let putCalled = false;
    const { client, calls } = fakeClient({
      post: async (path) => {
        if (path === "/groups") return { id: "group-1" };
        throw new Error(`unexpected POST ${path}`);
      },
    });
    // override put separately since createTeamFromGroup uses PUT /groups/{id}/team
    (client.api as any) = (path: string) => {
      if (path === "/groups") {
        return { post: vi.fn(async (body: unknown) => { calls.push({ path, body }); return { id: "group-1" }; }) };
      }
      if (path === "/groups/group-1/team") {
        return { put: vi.fn(async () => { putCalled = true; return {}; }) };
      }
      throw new Error(`unexpected path ${path}`);
    };

    const result = await createTeam(client, "CF-Demo Team", "cfdemoteam", "owner-1");
    expect(result).toEqual({ groupId: "group-1", displayName: "CF-Demo Team" });
    expect(putCalled).toBe(true);
    const groupCall = calls.find((c) => c.path === "/groups");
    expect((groupCall!.body as any)["owners@odata.bind"]).toEqual(["https://graph.microsoft.com/v1.0/users/owner-1"]);
  });

  it("retries PUT .../team on a 404 (group-replication-lag, per Microsoft's own documented advice) instead of failing immediately", async () => {
    vi.useFakeTimers();
    let attempts = 0;
    const client = {
      api: (path: string) => {
        if (path === "/groups") return { post: vi.fn(async () => ({ id: "group-2" })) };
        if (path === "/groups/group-2/team") {
          return {
            put: vi.fn(async () => {
              attempts++;
              if (attempts < 2) {
                const err: any = new Error("not found");
                err.statusCode = 404;
                throw err;
              }
              return {};
            }),
          };
        }
        throw new Error(`unexpected path ${path}`);
      },
    } as unknown as Client;

    const promise = createTeam(client, "Team", "team", "owner-1");
    await vi.advanceTimersByTimeAsync(10_000);
    await promise;
    expect(attempts).toBe(2);
    vi.useRealTimers();
  });
});

describe("createChannel / addTeamMember / addChannelMember", () => {
  it("createChannel defaults to a standard channel with the given displayName", async () => {
    const { client, calls } = fakeClient({ post: async () => ({ id: "chan-1", displayName: "General" }) });
    const channel = await createChannel(client, "team-1", "General");
    expect(channel).toEqual({ id: "chan-1", displayName: "General", membershipType: "standard" });
    expect(calls[0]!.path).toBe("/teams/team-1/channels");
    expect((calls[0]!.body as any).membershipType).toBe("standard");
  });

  it("createChannel with membershipType 'private' seeds an explicit owner member (only private channels support an explicit membership list)", async () => {
    const { client, calls } = fakeClient({ post: async () => ({ id: "chan-2", displayName: "Leads Only" }) });
    const channel = await createChannel(client, "team-1", "Leads Only", { membershipType: "private", ownerUserId: "owner-1" });
    expect(channel.membershipType).toBe("private");
    const body = calls[0]!.body as any;
    expect(body.membershipType).toBe("private");
    expect(body.members[0]["user@odata.bind"]).toBe("https://graph.microsoft.com/v1.0/users/owner-1");
  });

  it("addTeamMember binds an existing user, never creating a new/external identity", async () => {
    const { client, calls } = fakeClient({ post: async () => ({}) });
    await addTeamMember(client, "team-1", "user-1");
    const body = calls[0]!.body as any;
    expect(body["user@odata.bind"]).toBe("https://graph.microsoft.com/v1.0/users/user-1");
  });

  it("addChannelMember binds an existing user to a specific channel, not the team as a whole", async () => {
    const { client, calls } = fakeClient({ post: async () => ({}) });
    await addChannelMember(client, "team-1", "chan-2", "user-2");
    expect(calls[0]!.path).toBe("/teams/team-1/channels/chan-2/members");
    expect((calls[0]!.body as any)["user@odata.bind"]).toBe("https://graph.microsoft.com/v1.0/users/user-2");
  });
});
