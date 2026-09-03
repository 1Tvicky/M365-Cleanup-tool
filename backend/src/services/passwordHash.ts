import bcrypt from "bcryptjs";

const SALT_ROUNDS = 12;

export async function hashPassword(plaintext: string): Promise<string> {
  return bcrypt.hash(plaintext, SALT_ROUNDS);
}

export async function verifyPassword(plaintext: string, hash: string | null): Promise<boolean> {
  // Runs a real bcrypt compare against a fixed dummy hash even when the account has no password
  // (SSO-only operator) or doesn't exist, so response time doesn't leak which case occurred
  // (LOGIN-SEC-014 user enumeration via timing).
  const target = hash ?? "$2a$12$XwwZ49AGCq2L02ovhHrwPev2elX0kJIhYkkCWoA6ckXWDA8/zxt9S";
  const ok = await bcrypt.compare(plaintext, target);
  return hash !== null && ok;
}
