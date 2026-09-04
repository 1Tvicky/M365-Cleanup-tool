import { describe, expect, it } from "vitest";
import { classifyDeleteError } from "./cleanupDeletion.js";

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
