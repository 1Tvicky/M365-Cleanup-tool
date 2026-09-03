import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";

/**
 * Envelope encryption for any secret that must be persisted at rest (e.g. a future delegated
 * refresh token — application-permission tokens per docs/azure-ad-app-registration.md §5 are
 * never persisted). The data key below should come from a KMS/Key Vault in production; this
 * module only handles the AES-GCM operation, not key custody.
 */

const ALGO = "aes-256-gcm";

export function encryptSecret(plaintext: string, dataKey: Buffer): string {
  const iv = randomBytes(12);
  const cipher = createCipheriv(ALGO, dataKey, iv);
  const encrypted = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const authTag = cipher.getAuthTag();
  return Buffer.concat([iv, authTag, encrypted]).toString("base64");
}

export function decryptSecret(ciphertext: string, dataKey: Buffer): string {
  const raw = Buffer.from(ciphertext, "base64");
  const iv = raw.subarray(0, 12);
  const authTag = raw.subarray(12, 28);
  const encrypted = raw.subarray(28);
  const decipher = createDecipheriv(ALGO, dataKey, iv);
  decipher.setAuthTag(authTag);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString("utf8");
}
