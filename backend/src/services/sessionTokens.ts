import { randomUUID } from "node:crypto";
import jwt from "jsonwebtoken";
import { config } from "../config/index.js";
import { connection } from "../jobs/queue.js";

export interface SessionTokenPayload {
  sub: string; // operator id
  email: string;
  jti: string;
}

/** Issues a fresh session JWT with a new jti every time — never reuses a pre-auth token (LOGIN-SEC-011 session fixation). */
export function issueSessionToken(operatorId: string, email: string): { token: string; jti: string } {
  const jti = randomUUID();
  const token = jwt.sign({ email, jti } satisfies Omit<SessionTokenPayload, "sub">, config.session.jwtSecret, {
    subject: operatorId,
    issuer: config.session.issuer,
    expiresIn: `${config.session.ttlMinutes}m`,
  });
  return { token, jti };
}

export function verifySessionToken(token: string): SessionTokenPayload & { iat: number; exp: number } {
  const payload = jwt.verify(token, config.session.jwtSecret, { issuer: config.session.issuer }) as jwt.JwtPayload;
  return {
    sub: payload.sub as string,
    email: payload.email as string,
    jti: payload.jti as string,
    iat: payload.iat as number,
    exp: payload.exp as number,
  };
}

/** Logout (LOGIN-SES-002/003) — blocks the specific token's jti until it would have expired anyway. */
export async function revokeToken(jti: string, expiresAt: number): Promise<void> {
  const ttlSeconds = Math.max(expiresAt - Math.floor(Date.now() / 1000), 1);
  await connection.set(`revoked-jti:${jti}`, "1", "EX", ttlSeconds);
}

export async function isTokenRevoked(jti: string): Promise<boolean> {
  const value = await connection.get(`revoked-jti:${jti}`);
  return value !== null;
}
