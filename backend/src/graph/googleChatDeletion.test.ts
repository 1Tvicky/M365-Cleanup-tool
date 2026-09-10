import { describe, expect, it, vi } from "vitest";
import type { chat_v1 } from "googleapis";
import { classifyChatDeleteError, deleteChatMessage } from "./googleChatDeletion.js";

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
