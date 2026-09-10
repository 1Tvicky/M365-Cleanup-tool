import { describe, expect, it, vi } from "vitest";
import type { chat_v1 } from "googleapis";
import { getChatSpaceMembership, listAllChatSpaces, listSpaceMessages } from "./googleChatEnumeration.js";

describe("listAllChatSpaces", () => {
  it("searches with useAdminAccess:true and spaceType='SPACE' — every real space in the domain, not import-mode or the caller's own memberships", async () => {
    const search = vi.fn().mockResolvedValueOnce({ data: { spaces: [{ name: "spaces/AAA", displayName: "Marketing" }] } });
    const chat = { spaces: { search } } as unknown as chat_v1.Chat;

    const spaces = await listAllChatSpaces(chat);

    expect(spaces).toEqual([{ id: "AAA", displayName: "Marketing" }]);
    expect(search).toHaveBeenCalledWith(expect.objectContaining({ useAdminAccess: true, query: expect.stringContaining("spaceType = 'SPACE'") }));
  });
});

describe("getChatSpaceMembership", () => {
  it("counts only HUMAN members and picks the first one as the impersonation candidate", async () => {
    const list = vi.fn().mockResolvedValueOnce({
      data: {
        memberships: [{ member: { name: "users/111", type: "HUMAN" } }, { member: { name: "users/222", type: "HUMAN" } }],
      },
    });
    const chat = { spaces: { members: { list } } } as unknown as chat_v1.Chat;

    const membership = await getChatSpaceMembership(chat, "AAA");

    expect(membership).toEqual({ memberCount: 2, aHumanMemberId: "111" });
    expect(list).toHaveBeenCalledWith(expect.objectContaining({ parent: "spaces/AAA", useAdminAccess: true }));
  });

  it("returns aHumanMemberId: null when the space has no human members", async () => {
    const list = vi.fn().mockResolvedValueOnce({ data: { memberships: [] } });
    const chat = { spaces: { members: { list } } } as unknown as chat_v1.Chat;
    expect(await getChatSpaceMembership(chat, "AAA")).toEqual({ memberCount: 0, aHumanMemberId: null });
  });
});

describe("listSpaceMessages", () => {
  it("paginates and returns each message's full resource name", async () => {
    const list = vi.fn().mockResolvedValueOnce({ data: { messages: [{ name: "spaces/AAA/messages/1" }] } });
    const chat = { spaces: { messages: { list } } } as unknown as chat_v1.Chat;
    expect(await listSpaceMessages(chat, "AAA")).toEqual([{ name: "spaces/AAA/messages/1" }]);
  });
});
