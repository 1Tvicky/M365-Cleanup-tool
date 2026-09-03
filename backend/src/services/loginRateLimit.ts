import { connection } from "../jobs/queue.js";

/**
 * Fixed-window Redis counter — used for the coarse, IP-scoped brute-force guard on /auth/login
 * and /auth/forgot-password (LOGIN-SEC-005). Per-account lockout (locked account, failed count)
 * is tracked durably on the operators row instead, since that's user-facing account state, not
 * just a rate-limit signal — see routes/session.ts.
 */
export async function hitRateLimit(
  key: string,
  opts: { limit: number; windowSeconds: number }
): Promise<{ allowed: boolean; retryAfterSeconds: number }> {
  const redisKey = `ratelimit:${key}`;
  const count = await connection.incr(redisKey);
  if (count === 1) {
    await connection.expire(redisKey, opts.windowSeconds);
  }
  if (count <= opts.limit) {
    return { allowed: true, retryAfterSeconds: 0 };
  }
  const ttl = await connection.ttl(redisKey);
  return { allowed: false, retryAfterSeconds: Math.max(ttl, 1) };
}
