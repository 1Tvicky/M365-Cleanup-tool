import { randomBytes, timingSafeEqual } from "node:crypto";
import type { NextFunction, Request, Response } from "express";
import { config } from "../config/index.js";
import { ApiError } from "../types/index.js";

const CSRF_COOKIE = "cf_csrf";
const CSRF_HEADER = "x-csrf-token";

/** Mints the CSRF cookie (readable by JS) that the frontend must echo back in a header on state-changing requests. */
export function issueCsrfCookie(res: Response): void {
  const token = randomBytes(24).toString("base64url");
  res.cookie(CSRF_COOKIE, token, {
    httpOnly: false, // must be readable by the frontend to echo back — this cookie carries no auth power on its own
    secure: config.nodeEnv === "production",
    sameSite: "lax",
    domain: config.session.cookieDomain,
    maxAge: 4 * 60 * 60 * 1000,
  });
}

/**
 * Double-submit CSRF check (LOGIN-SEC-012) — a cross-site form/script can trigger the request and
 * carry the session cookie automatically, but cannot read the CSRF cookie to also set the header,
 * since that cookie isn't sent to other origins and the header must be set explicitly by JS.
 */
export function requireCsrf(req: Request, _res: Response, next: NextFunction): void {
  const cookieToken = req.cookies?.[CSRF_COOKIE];
  const headerToken = req.header(CSRF_HEADER);

  if (!cookieToken || !headerToken || cookieToken.length !== headerToken.length) {
    throw new ApiError(403, "CSRF_VALIDATION_FAILED", "Missing or invalid CSRF token");
  }
  const match = timingSafeEqual(Buffer.from(cookieToken), Buffer.from(headerToken));
  if (!match) {
    throw new ApiError(403, "CSRF_VALIDATION_FAILED", "Missing or invalid CSRF token");
  }
  next();
}

/**
 * Prevents open-redirect via a manipulated `redirect` query param (LOGIN-SEC-013) — only a
 * same-origin relative path is accepted; anything else (protocol-relative //host, absolute URL)
 * falls back to the default.
 */
export function safeRedirectPath(candidate: unknown, fallback = "/"): string {
  if (typeof candidate !== "string" || !candidate.startsWith("/") || candidate.startsWith("//")) {
    return fallback;
  }
  return candidate;
}
