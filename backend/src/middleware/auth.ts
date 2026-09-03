import type { NextFunction, Request, Response } from "express";
import { config } from "../config/index.js";
import { query } from "../db/pool.js";
import { isTokenRevoked, verifySessionToken } from "../services/sessionTokens.js";
import { ApiError } from "../types/index.js";

export interface SessionClaims {
  operatorId: string;
  email: string;
  jti: string;
}

declare global {
  // eslint-disable-next-line @typescript-eslint/no-namespace
  namespace Express {
    interface Request {
      session?: SessionClaims;
    }
  }
}

function extractToken(req: Request): string | null {
  const cookieToken = req.cookies?.[config.session.cookieName];
  if (cookieToken) return cookieToken;

  // Bearer header stays supported for non-browser callers (scripts, CI) — browsers should rely
  // on the httpOnly cookie so the token is never reachable from page JS (LOGIN-SEC-007).
  const header = req.headers.authorization;
  if (header?.startsWith("Bearer ")) return header.slice("Bearer ".length);

  return null;
}

/**
 * Verifies the CloudFuze operator session — distinct from Graph tokens, which are per-tenant and
 * never exposed to the frontend. Beyond signature/expiry, also rejects a token that was
 * explicitly logged out (jti revocation, LOGIN-SES-002/003) or issued before the account's most
 * recent password change (credentials_valid_after, LOGIN-FP-013 / LOGIN-SEC-011).
 */
export async function requireSession(req: Request, _res: Response, next: NextFunction): Promise<void> {
  try {
    const token = extractToken(req);
    if (!token) {
      throw new ApiError(401, "UNAUTHENTICATED", "Missing session token");
    }

    let claims: ReturnType<typeof verifySessionToken>;
    try {
      claims = verifySessionToken(token);
    } catch {
      throw new ApiError(401, "UNAUTHENTICATED", "Invalid or expired session token");
    }

    if (await isTokenRevoked(claims.jti)) {
      throw new ApiError(401, "UNAUTHENTICATED", "Session has been logged out");
    }

    const result = await query<{ credentials_valid_after: string; status: string }>(
      `SELECT credentials_valid_after, status FROM operators WHERE id = $1`,
      [claims.sub]
    );
    const operator = result.rows[0];
    if (!operator || operator.status !== "active") {
      throw new ApiError(401, "UNAUTHENTICATED", "Account is no longer active");
    }
    if (claims.iat * 1000 < new Date(operator.credentials_valid_after).getTime()) {
      throw new ApiError(401, "UNAUTHENTICATED", "Session was issued before the most recent password change");
    }

    req.session = { operatorId: claims.sub, email: claims.email, jti: claims.jti };
    next();
  } catch (err) {
    next(err);
  }
}
