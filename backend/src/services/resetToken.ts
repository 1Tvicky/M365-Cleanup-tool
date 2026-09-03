import { createHash, randomBytes } from "node:crypto";

/** Raw token goes only in the emailed reset link; the DB stores just its SHA-256 hash (schema.sql password_reset_tokens). */
export function generateResetToken(): { raw: string; hash: string } {
  const raw = randomBytes(32).toString("base64url");
  return { raw, hash: hashResetToken(raw) };
}

export function hashResetToken(raw: string): string {
  return createHash("sha256").update(raw).digest("hex");
}
