import { describe, expect, it, vi } from "vitest";
import type { chat_v1 } from "googleapis";
import { classifyChatDeleteError, deleteChatMessage, deleteChatSpace } from "./googleChatDeletion.js";

describe("deleteChatMessage", () => {
  it("calls messages.delete with force:true so threaded replies are removed too", async () => {
    const del = vi.fn(() => Promise.resolve({}));
    const chat = { spaces: { messages: { delete: del } } } as unknown as chat_v1.Chat;
    const result = await deleteChatMessage(chat, "spaces/s1/messages/m1");
    expect(result).toBe("deleted");
    expect(del).toHaveBeenCalledWith({ name: "spaces/s1/messages/m1", force: true });
  });

  it("treats a 404 as already_gone", async () => {
    const del = vi.fn(() => Promise.reject({ code: 404 }));
    const chat = { spaces: { messages: { delete: del } } } as unknown as chat_v1.Chat;
    expect(await deleteChatMessage(chat, "spaces/s1/messages/m1")).toBe("already_gone");
  });

  it("propagates a non-404 error unchanged", async () => {
    const del = vi.fn(() => Promise.reject({ code: 403 }));
    const chat = { spaces: { messages: { delete: del } } } as unknown as chat_v1.Chat;
    await expect(deleteChatMessage(chat, "spaces/s1/messages/m1")).rejects.toBeTruthy();
  });
});

describe("classifyChatDeleteError", () => {
  it("classifies 403 as an insufficient-permission failure", () => {
    expect(classifyChatDeleteError({ code: 403 }).code).toBe("INSUFFICIENT_PERMISSION");
  });
});

describe("deleteChatSpace", () => {
  it("calls spaces.delete with useAdminAccess:true — the Space delete API, not message deletion", async () => {
    const del = vi.fn(() => Promise.resolve({}));
    const messagesDel = vi.fn();
    const chat = { spaces: { delete: del, messages: { delete: messagesDel } } } as unknown as chat_v1.Chat;
    const result = await deleteChatSpace(chat, "s1");

    expect(result).toBe("deleted");
    expect(del).toHaveBeenCalledWith({ name: "spaces/s1", useAdminAccess: true });
    expect(messagesDel).not.toHaveBeenCalled();
  });

  it("treats a 404 (space not found / already deleted) as already_gone, not a failure", async () => {
    const del = vi.fn(() => Promise.reject({ code: 404 }));
    const chat = { spaces: { delete: del } } as unknown as chat_v1.Chat;
    expect(await deleteChatSpace(chat, "s1")).toBe("already_gone");
  });

  it("propagates a permission-denied error unchanged, and it classifies as insufficient permission", async () => {
    const del = vi.fn(() => Promise.reject({ code: 403 }));
    const chat = { spaces: { delete: del } } as unknown as chat_v1.Chat;
    await expect(deleteChatSpace(chat, "s1")).rejects.toMatchObject({ code: 403 });
    expect(classifyChatDeleteError({ code: 403 }).code).toBe("INSUFFICIENT_PERMISSION");
  });

  it("propagates a rate-limit error unchanged, and it classifies as retryable", async () => {
    const del = vi.fn(() => Promise.reject({ code: 429 }));
    const chat = { spaces: { delete: del } } as unknown as chat_v1.Chat;
    await expect(deleteChatSpace(chat, "s1")).rejects.toMatchObject({ code: 429 });
    expect(classifyChatDeleteError({ code: 429 }).code).toBe("RATE_LIMITED");
  });

  it("rejects for an invalid space id, still classifiable rather than silently swallowed", async () => {
    const del = vi.fn(() => Promise.reject({ code: 400, message: "Invalid space id" }));
    const chat = { spaces: { delete: del } } as unknown as chat_v1.Chat;
    await expect(deleteChatSpace(chat, "not-a-real-space")).rejects.toMatchObject({ code: 400 });
  });
});
