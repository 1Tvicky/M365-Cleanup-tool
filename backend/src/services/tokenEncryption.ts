import { decryptSecret, encryptSecret } from "./encryption.js";
import { config } from "../config/index.js";

/**
 * Field-level encryption for `connections.encrypted_refresh_token`. The data key here comes from
 * an env var for local dev; in production this must be a KMS-managed key (AWS KMS / Azure Key
 * Vault) fetched at startup, never a static env value — swap `getDataKey()`'s implementation, the
 * call sites don't need to change.
 */
function getDataKey(): Buffer {
  const raw = config.tokenEncryption.key;
  if (!raw) {
    throw new Error("TOKEN_ENCRYPTION_KEY is not configured — cannot encrypt/decrypt connection tokens");
  }
  const key = Buffer.from(raw, "base64");
  if (key.length !== 32) {
    throw new Error("TOKEN_ENCRYPTION_KEY must decode to exactly 32 bytes (AES-256)");
  }
  return key;
}

export function encryptToken(plaintext: string): string {
  return encryptSecret(plaintext, getDataKey());
}

export function decryptToken(ciphertext: string): string {
  return decryptSecret(ciphertext, getDataKey());
}
