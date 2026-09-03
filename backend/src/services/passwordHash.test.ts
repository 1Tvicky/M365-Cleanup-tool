import { describe, expect, it } from "vitest";
import { hashPassword, verifyPassword } from "./passwordHash.js";

describe("hashPassword / verifyPassword", () => {
  it("round-trips a valid password (LOGIN-P-001)", async () => {
    const hash = await hashPassword("Correct-Horse-Battery-1");
    await expect(verifyPassword("Correct-Horse-Battery-1", hash)).resolves.toBe(true);
  });

  it("rejects the wrong password against a real hash (LOGIN-PWD-002/003)", async () => {
    const hash = await hashPassword("Correct-Horse-Battery-1");
    await expect(verifyPassword("wrong-guess", hash)).resolves.toBe(false);
  });

  it("accepts a valid special-character password (LOGIN-P-012 / LOGIN-PWD-008)", async () => {
    const password = "p@$$w0rd!#%&*()";
    const hash = await hashPassword(password);
    await expect(verifyPassword(password, hash)).resolves.toBe(true);
  });

  it("handles an extremely long password without throwing (LOGIN-PWD-007)", async () => {
    const password = "a".repeat(128);
    const hash = await hashPassword(password);
    await expect(verifyPassword(password, hash)).resolves.toBe(true);
  });

  it("returns false (not a throw) when the account has no local password — SSO-only operator (LOGIN-SEC-014 timing safety)", async () => {
    await expect(verifyPassword("anything", null)).resolves.toBe(false);
  });
});
