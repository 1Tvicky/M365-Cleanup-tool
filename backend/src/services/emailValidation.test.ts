import { describe, expect, it } from "vitest";
import { isValidLoginEmail, normalizeEmail } from "./emailValidation.js";

// Each case cites the test-case ID it satisfies from CloudFuze_Login_Page_Test_Cases.docx —
// see docs/login-test-case-coverage.md for the full mapping.
describe("isValidLoginEmail", () => {
  it.each([
    ["test@example.com", "LOGIN-E-001 valid email"],
    ["user@mail.example.com", "LOGIN-E-002 email with subdomain"],
    ["user123@example.com", "LOGIN-E-003 email with numbers"],
    ["user+test@example.com", "LOGIN-E-004 email with plus sign"],
  ])("accepts %s (%s)", (email) => {
    expect(isValidLoginEmail(email)).toBe(true);
  });

  it.each([
    ["", "LOGIN-E-005 empty email"],
    ["test", "LOGIN-E-006 invalid email (no @, no domain)"],
    ["testexample.com", "LOGIN-E-007 missing @"],
    ["test@", "LOGIN-E-008 missing domain"],
    ["@example.com", "LOGIN-E-009 missing username"],
    ["test@@example.com", "LOGIN-E-010 multiple @ symbols"],
    ["test..user@example.com", "LOGIN-E-011 consecutive dots"],
    ["test user@example.com", "LOGIN-E-012 space inside email"],
    ["!@#$%^", "LOGIN-E-014 special characters only"],
    ["test@example.com\nBcc: attacker@evil.com", "LOGIN-E-019 multiline input"],
    ["<script>alert(1)</script>@example.com", "LOGIN-E-016/017 HTML/XSS payload"],
    ["' OR '1'='1@example.com", "LOGIN-E-018 SQL injection payload"],
  ])("rejects %s (%s)", (email) => {
    expect(isValidLoginEmail(email)).toBe(false);
  });

  it("rejects an email exceeding the supported length (LOGIN-E-015)", () => {
    const huge = `${"a".repeat(250)}@example.com`;
    expect(isValidLoginEmail(huge)).toBe(false);
  });

  it("does not support unicode local parts in v1 (LOGIN-E-020 — documented, not silently mishandled)", () => {
    expect(isValidLoginEmail("üser@example.com")).toBe(false);
  });
});

describe("normalizeEmail", () => {
  it("trims surrounding whitespace (LOGIN-E-013 / LOGIN-P-013)", () => {
    expect(normalizeEmail("  test@example.com  ")).toBe("test@example.com");
  });

  it("only-whitespace input normalizes to empty, which isValidLoginEmail then rejects (LOGIN-E-013)", () => {
    expect(normalizeEmail("   ")).toBe("");
    expect(isValidLoginEmail(normalizeEmail("   "))).toBe(false);
  });

  it("lowercases so lookups are case-insensitive (LOGIN-P-003/004/005)", () => {
    expect(normalizeEmail("Test.User@EXAMPLE.com")).toBe("test.user@example.com");
  });
});
