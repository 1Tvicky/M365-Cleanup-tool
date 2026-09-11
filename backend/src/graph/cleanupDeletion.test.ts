import { describe, expect, it } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import { classifyDeleteError, deleteCalendarEvent, deleteContact, deleteDriveItem, deleteMessage, deleteTeam, deleteTeamsChannel } from "./cleanupDeletion.js";

describe("classifyDeleteError", () => {
  it("classifies 403 as an insufficient-permission failure", () => {
    const result = classifyDeleteError({ statusCode: 403 });
    expect(result.code).toBe("INSUFFICIENT_PERMISSION");
    expect(result.message).toMatch(/permission/i);
  });

  it("classifies 409 as a retryable conflict", () => {
    const result = classifyDeleteError({ statusCode: 409 });
    expect(result.code).toBe("CONFLICT");
    expect(result.message).toMatch(/retried/i);
  });

  it("falls back to the status code and message for anything else", () => {
    const result = classifyDeleteError({ statusCode: 500, message: "Internal server error" });
    expect(result.code).toBe("500");
    expect(result.message).toBe("Internal server error");
  });

  it("handles an error with no status code at all", () => {
    const result = classifyDeleteError(new Error("network failure"));
    expect(result.code).toBe("UNKNOWN");
    expect(result.message).toBe("network failure");
  });
});

/**
 * A minimal stand-in for the Graph SDK's fluent `client.api(path).method()` builder — records every
 * call so tests can assert on the exact path/verb/body used, and resolves/rejects per `respond`.
 */
function fakeClient(respond: (path: string) => Promise<unknown>) {
  const calls: { path: string; verb: "get" | "post" | "delete"; body?: unknown }[] = [];
  const client = {
    api: (path: string) => ({
      get: () => {
        calls.push({ path, verb: "get" });
        return respond(path);
      },
      post: (body: unknown) => {
        calls.push({ path, verb: "post", body });
        return respond(path);
      },
      delete: () => {
        calls.push({ path, verb: "delete" });
        return respond(path);
      },
    }),
  } as unknown as Client;
  return { client, calls };
}

describe("deleteDriveItem", () => {
  it("permanent=true calls permanentDelete (POST), not plain DELETE, with no request body", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteDriveItem(client, "user", "user-1", "item-1", true);

    expect(result).toBe("deleted");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.verb).toBe("post");
    expect(calls[0]!.path).toBe("/users/user-1/drive/items/item-1/permanentDelete");
    expect(calls[0]!.body).toBeUndefined();
  });

  it("permanent=false still calls plain DELETE (recycle-bin mode, chosen on the confirmation screen)", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteDriveItem(client, "user", "user-1", "item-1", false);

    expect(result).toBe("deleted");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/users/user-1/drive/items/item-1");
    expect(calls[0]!.path).not.toMatch(/permanentDelete/);
  });

  it("uses the site path for a SharePoint drive item", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    await deleteDriveItem(client, "site", "site-1", "item-2", true);
    expect(calls[0]!.path).toBe("/sites/site-1/drive/items/item-2/permanentDelete");
  });

  it("treats a 404 as already_gone, not a failure, in either mode", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 404 }));
    expect(await deleteDriveItem(client, "user", "user-1", "item-1", true)).toBe("already_gone");
    expect(await deleteDriveItem(client, "user", "user-1", "item-1", false)).toBe("already_gone");
  });

  it("propagates a non-404 error unchanged, still classifiable", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 403 }));
    await expect(deleteDriveItem(client, "user", "user-1", "item-1", true)).rejects.toMatchObject({ statusCode: 403 });
  });
});

describe("deleteMessage", () => {
  it("permanent=true calls permanentDelete (POST), not plain DELETE, with no request body", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteMessage(client, "user-1", "message-1", true);

    expect(result).toBe("deleted");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.verb).toBe("post");
    expect(calls[0]!.path).toBe("/users/user-1/messages/message-1/permanentDelete");
    expect(calls[0]!.body).toBeUndefined();
  });

  it("permanent=false still calls plain DELETE (recycle-bin mode)", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteMessage(client, "user-1", "message-1", false);

    expect(result).toBe("deleted");
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/users/user-1/messages/message-1");
    expect(calls[0]!.path).not.toMatch(/permanentDelete/);
  });

  it("treats a 404 as already_gone in either mode (covers: never touched, already in Deleted Items separately, or already purged by a prior partial run)", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 404 }));
    expect(await deleteMessage(client, "user-1", "message-1", true)).toBe("already_gone");
    expect(await deleteMessage(client, "user-1", "message-1", false)).toBe("already_gone");
  });

  it("propagates a non-404 error unchanged", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 403 }));
    await expect(deleteMessage(client, "user-1", "message-1", true)).rejects.toMatchObject({ statusCode: 403 });
  });
});

/**
 * Regression guard: Calendar/Contacts are deliberately out of scope for permanent deletion (only
 * OneDrive, SharePoint, and Outlook mail are). These four functions are textually near-identical —
 * this is the one thing that would catch a copy-paste change accidentally widening scope later.
 */
describe("deleteCalendarEvent / deleteContact stay on plain soft delete", () => {
  it("deleteCalendarEvent still calls plain DELETE, never permanentDelete", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    await deleteCalendarEvent(client, "user-1", "event-1");
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/users/user-1/events/event-1");
    expect(calls[0]!.path).not.toMatch(/permanentDelete/);
  });

  it("deleteContact still calls plain DELETE, never permanentDelete", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    await deleteContact(client, "user-1", "contact-1");
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/users/user-1/contacts/contact-1");
    expect(calls[0]!.path).not.toMatch(/permanentDelete/);
  });
});

describe("deleteTeamsChannel", () => {
  it("calls DELETE /teams/{teamId}/channels/{channelId} — the channel delete API, not a message-delete path", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteTeamsChannel(client, "team-1", "channel-1");

    expect(result).toBe("deleted");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/teams/team-1/channels/channel-1");
  });

  it("never touches a different team's channel with the same channel id", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    await deleteTeamsChannel(client, "team-A", "channel-1");
    expect(calls[0]!.path).toBe("/teams/team-A/channels/channel-1");
    expect(calls[0]!.path).not.toContain("team-B");
  });

  it("treats a 404 (already deleted) as already_gone, not a failure", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 404 }));
    expect(await deleteTeamsChannel(client, "team-1", "channel-1")).toBe("already_gone");
  });

  it("propagates a non-404 error (e.g. insufficient permission) unchanged, still classifiable", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 403 }));
    await expect(deleteTeamsChannel(client, "team-1", "channel-1")).rejects.toMatchObject({ statusCode: 403 });
    expect(classifyDeleteError({ statusCode: 403 }).code).toBe("INSUFFICIENT_PERMISSION");
  });

  it("rejects with an invalid-resource-style error for a malformed/invalid channel id, and it's still classifiable", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 400, message: "Invalid channel id" }));
    await expect(deleteTeamsChannel(client, "team-1", "not-a-real-channel")).rejects.toMatchObject({ statusCode: 400 });
  });
});

describe("deleteTeam", () => {
  it("calls DELETE /groups/{id} — deletes the Team's backing M365 Group, never a channel path", async () => {
    const { client, calls } = fakeClient(() => Promise.resolve());
    const result = await deleteTeam(client, "team-1");

    expect(result).toBe("deleted");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.verb).toBe("delete");
    expect(calls[0]!.path).toBe("/groups/team-1");
    expect(calls[0]!.path).not.toMatch(/channels/);
  });

  it("treats a 404 (already deleted) as already_gone, not a failure", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 404 }));
    expect(await deleteTeam(client, "team-1")).toBe("already_gone");
  });

  it("propagates a non-404 error (e.g. insufficient permission) unchanged, still classifiable", async () => {
    const { client } = fakeClient(() => Promise.reject({ statusCode: 403 }));
    await expect(deleteTeam(client, "team-1")).rejects.toMatchObject({ statusCode: 403 });
    expect(classifyDeleteError({ statusCode: 403 }).code).toBe("INSUFFICIENT_PERMISSION");
  });
});
