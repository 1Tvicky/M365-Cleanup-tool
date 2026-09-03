// Deliberately stricter than a permissive HTML5-style email regex: single '@', no
// leading/trailing/consecutive dots in the local part, no whitespace or control characters
// anywhere. This is what rejects the malformed-email family of test cases (LOGIN-E-006..013) —
// see docs/login-test-case-coverage.md.
const STRICT_EMAIL = /^(?!\.)(?!.*\.\.)[A-Za-z0-9._%+-]+(?<!\.)@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$/;

export const EMAIL_MAX_LENGTH = 254; // RFC 5321 §4.5.3.1.3

/** Trims surrounding whitespace (LOGIN-E-013) and lowercases for case-insensitive lookup (LOGIN-P-003/004/005). */
export function normalizeEmail(raw: string): string {
  return raw.trim().toLowerCase();
}

export function isValidLoginEmail(raw: string): boolean {
  if (raw.length === 0 || raw.length > EMAIL_MAX_LENGTH) return false;
  if (/[\r\n]/.test(raw)) return false;
  return STRICT_EMAIL.test(raw);
}
