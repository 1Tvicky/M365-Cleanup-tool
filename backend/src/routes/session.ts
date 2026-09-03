import { randomBytes } from "node:crypto";
import { Router } from "express";
import { z } from "zod";
import { config } from "../config/index.js";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { issueCsrfCookie, requireCsrf, safeRedirectPath } from "../middleware/csrf.js";
import { hashPassword, verifyPassword } from "../services/passwordHash.js";
import { EMAIL_MAX_LENGTH, isValidLoginEmail, normalizeEmail } from "../services/emailValidation.js";
import { hitRateLimit } from "../services/loginRateLimit.js";
import { generateResetToken, hashResetToken } from "../services/resetToken.js";
import { issueSessionToken, revokeToken, verifySessionToken } from "../services/sessionTokens.js";
import { getGoogleAuthUrl, resolveGoogleIdentity } from "../services/googleAuth.js";
import { getMicrosoftAuthUrl, resolveMicrosoftIdentity } from "../services/microsoftUserAuth.js";
import { ApiError } from "../types/index.js";

export const sessionRouter = Router();

const emailField = z
  .string()
  .max(EMAIL_MAX_LENGTH * 2, "Email is too long") // generous pre-trim cap; isValidLoginEmail enforces the real limit post-trim
  .transform(normalizeEmail)
  .refine((v) => v.length > 0, "Email is required")
  .refine(isValidLoginEmail, "Enter a valid email address");

const passwordField = z.string().min(1, "Password is required").max(config.password.maxLength, "Password is too long");

interface OperatorRow {
  id: string;
  email: string;
  display_name: string;
  password_hash: string | null;
  status: "active" | "unverified" | "disabled" | "deleted";
  failed_login_count: number;
  locked_until: string | null;
}

async function findOperatorByEmail(email: string): Promise<OperatorRow | null> {
  const result = await query<OperatorRow>(
    `SELECT id, email, display_name, password_hash, status, failed_login_count, locked_until
     FROM operators WHERE email = $1`,
    [email]
  );
  return result.rows[0] ?? null;
}

async function logAuthEvent(
  event: string,
  email: string,
  operatorId: string | null,
  req: { ip?: string; headers: Record<string, unknown> }
): Promise<void> {
  await query(
    `INSERT INTO auth_events (operator_id, email, event, ip_address, user_agent) VALUES ($1, $2, $3, $4, $5)`,
    [operatorId, email, event, req.ip ?? null, String(req.headers["user-agent"] ?? "")]
  );
}

function setSessionCookie(res: import("express").Response, token: string): void {
  res.cookie(config.session.cookieName, token, {
    httpOnly: true, // never readable by page JS — the frontend never sees the token (LOGIN-SEC-007)
    secure: config.nodeEnv === "production",
    sameSite: "lax", // must survive the top-level redirect back from Google/Microsoft's own login page
    domain: config.session.cookieDomain,
    path: "/",
    maxAge: config.session.ttlMinutes * 60 * 1000,
  });
  issueCsrfCookie(res);
}

/** Bootstraps the CSRF cookie before the frontend renders the login form. */
sessionRouter.get("/csrf", (_req, res) => {
  issueCsrfCookie(res);
  res.status(204).end();
});

sessionRouter.get("/me", requireSession, asyncHandler(async (req, res) => {
  const result = await query<{ id: string; email: string; display_name: string }>(
    `SELECT id, email, display_name FROM operators WHERE id = $1`,
    [req.session!.operatorId]
  );
  const operator = result.rows[0];
  if (!operator) throw new ApiError(401, "UNAUTHENTICATED", "Account no longer exists");
  res.json({ id: operator.id, email: operator.email, displayName: operator.display_name });
}));

const loginSchema = z.object({ email: emailField, password: passwordField });

/**
 * LOGIN-P-*, LOGIN-E-*, LOGIN-PWD-*, LOGIN-AUTH-*, LOGIN-SEC-001/002/005/014 — see
 * docs/login-test-case-coverage.md for the full test-ID-to-code mapping.
 */
sessionRouter.post("/login", asyncHandler(async (req, res) => {
  const parsed = loginSchema.safeParse(req.body);
  if (!parsed.success) {
    throw new ApiError(400, "VALIDATION_ERROR", "Invalid email or password format", parsed.error.flatten());
  }
  const { email, password } = parsed.data;

  const ipLimit = await hitRateLimit(`login:ip:${req.ip}`, { limit: 30, windowSeconds: 300 });
  const emailLimit = await hitRateLimit(`login:email:${email}`, { limit: 10, windowSeconds: 300 });
  if (!ipLimit.allowed || !emailLimit.allowed) {
    const retryAfter = Math.max(ipLimit.retryAfterSeconds, emailLimit.retryAfterSeconds);
    res.setHeader("Retry-After", String(retryAfter));
    throw new ApiError(429, "RATE_LIMITED", "Too many login attempts — try again shortly", { retryAfterSeconds: retryAfter });
  }

  const operator = await findOperatorByEmail(email);

  // Locked accounts are rejected before touching the password at all, and with the same message
  // regardless of whether the supplied password would have been correct.
  if (operator?.locked_until && new Date(operator.locked_until).getTime() > Date.now()) {
    await logAuthEvent("login_locked", email, operator.id, req);
    throw new ApiError(403, "ACCOUNT_LOCKED", "This account is temporarily locked due to repeated failed sign-in attempts. Try again later.");
  }

  // Always runs a real bcrypt compare (against a dummy hash when the account doesn't exist or has
  // no local password) so response timing can't be used to enumerate valid emails.
  const passwordOk = await verifyPassword(password, operator?.password_hash ?? null);

  if (!operator || !passwordOk) {
    if (operator) {
      const nextCount = operator.failed_login_count + 1;
      const lock = nextCount >= config.login.maxFailedAttempts;
      await query(
        `UPDATE operators SET failed_login_count = $2,
                locked_until = CASE WHEN $3 THEN now() + ($4 || ' minutes')::interval ELSE locked_until END
         WHERE id = $1`,
        [operator.id, nextCount, lock, config.login.lockoutMinutes]
      );
      await logAuthEvent(lock ? "login_locked" : "login_failed", email, operator.id, req);
    } else {
      await logAuthEvent("login_failed", email, null, req);
    }
    // Deliberately identical whether the email doesn't exist or the password was wrong.
    throw new ApiError(401, "INVALID_CREDENTIALS", "Invalid email or password");
  }

  if (operator.status !== "active") {
    const message =
      operator.status === "unverified"
        ? "Please verify your email address before logging in."
        : "This account is not available. Contact your administrator.";
    await logAuthEvent("login_failed", email, operator.id, req);
    throw new ApiError(403, "ACCOUNT_NOT_ACTIVE", message);
  }

  await query(`UPDATE operators SET failed_login_count = 0, locked_until = NULL WHERE id = $1`, [operator.id]);
  const { token } = issueSessionToken(operator.id, operator.email);
  setSessionCookie(res, token);
  await logAuthEvent("login_success", email, operator.id, req);

  res.json({ id: operator.id, email: operator.email, displayName: operator.display_name });
}));

/** LOGIN-SES-002/003 — best-effort revoke; a missing/expired token is still a successful logout from the client's perspective. */
sessionRouter.post("/logout", requireCsrf, asyncHandler(async (req, res) => {
  const token = req.cookies?.[config.session.cookieName];
  if (token) {
    try {
      const claims = verifySessionToken(token);
      await revokeToken(claims.jti, claims.exp);
      await logAuthEvent("logout", claims.email, claims.sub, req);
    } catch {
      // token already invalid/expired — nothing to revoke
    }
  }
  res.clearCookie(config.session.cookieName, { path: "/", domain: config.session.cookieDomain });
  res.status(204).end();
}));

const forgotPasswordSchema = z.object({ email: emailField });

/** LOGIN-FP-003..006 — always returns the same generic response so the caller can't learn whether an email is registered. */
sessionRouter.post("/forgot-password", requireCsrf, asyncHandler(async (req, res) => {
  const parsed = forgotPasswordSchema.safeParse(req.body);
  if (!parsed.success) {
    throw new ApiError(400, "VALIDATION_ERROR", "Enter a valid email address", parsed.error.flatten());
  }
  const { email } = parsed.data;

  const limit = await hitRateLimit(`forgot-password:${email}`, { limit: 3, windowSeconds: 3600 });
  if (!limit.allowed) {
    res.setHeader("Retry-After", String(limit.retryAfterSeconds));
    throw new ApiError(429, "RATE_LIMITED", "Too many reset requests — try again later");
  }

  const operator = await findOperatorByEmail(email);
  if (operator && operator.status !== "deleted") {
    const { raw, hash } = generateResetToken();
    await query(
      `INSERT INTO password_reset_tokens (operator_id, token_hash, expires_at)
       VALUES ($1, $2, now() + ($3 || ' minutes')::interval)`,
      [operator.id, hash, config.passwordReset.tokenTtlMinutes]
    );
    await logAuthEvent("password_reset_requested", email, operator.id, req);

    // Real delivery is an integration point (SES/SendGrid/etc), out of scope here — logging the
    // link keeps the flow runnable end-to-end in dev without a mail provider configured.
    console.log(`[password reset] ${email}: /reset-password?token=${raw}`);
  }

  res.status(202).json({ message: "If an account exists for that email, a reset link has been sent." });
}));

const resetPasswordSchema = z.object({
  token: z.string().min(1),
  newPassword: z
    .string()
    .min(config.password.minLength, `Password must be at least ${config.password.minLength} characters`)
    .max(config.password.maxLength, "Password is too long"),
});

/** LOGIN-FP-008..013 — single-use, expiring token; completing a reset invalidates every existing session for the account. */
sessionRouter.post("/reset-password", requireCsrf, asyncHandler(async (req, res) => {
  const parsed = resetPasswordSchema.safeParse(req.body);
  if (!parsed.success) {
    throw new ApiError(400, "VALIDATION_ERROR", "Invalid reset request", parsed.error.flatten());
  }
  const { token, newPassword } = parsed.data;

  const tokenHash = hashResetToken(token);
  const result = await query<{ id: string; operator_id: string; expires_at: string; used_at: string | null }>(
    `SELECT id, operator_id, expires_at, used_at FROM password_reset_tokens WHERE token_hash = $1`,
    [tokenHash]
  );
  const row = result.rows[0];
  const valid = row && !row.used_at && new Date(row.expires_at).getTime() > Date.now();
  if (!valid) {
    throw new ApiError(400, "INVALID_RESET_TOKEN", "This reset link is invalid or has expired");
  }

  const passwordHash = await hashPassword(newPassword);
  await query(
    `UPDATE operators
     SET password_hash = $2, credentials_valid_after = now(), failed_login_count = 0, locked_until = NULL,
         status = CASE WHEN status = 'unverified' THEN 'active' ELSE status END
     WHERE id = $1`,
    [row.operator_id, passwordHash]
  );
  await query(`UPDATE password_reset_tokens SET used_at = now() WHERE id = $1`, [row.id]);

  const operatorRow = await query<{ email: string }>(`SELECT email FROM operators WHERE id = $1`, [row.operator_id]);
  await logAuthEvent("password_reset_completed", operatorRow.rows[0]?.email ?? "", row.operator_id, req);

  res.json({ message: "Password updated. Please log in with your new password." });
}));

/* --- SSO: Google and Office 365. Both follow the same shape: GET /provider redirects to the
   provider's consent screen with a CSRF-bound `state`; GET /provider/callback exchanges the code,
   resolves a verified identity, and either finds or provisions the matching operator. --- */

const OAUTH_STATE_COOKIE = "cf_oauth_state";

function issueOAuthState(res: import("express").Response, redirectTo: string): string {
  const state = randomBytes(16).toString("base64url");
  res.cookie(OAUTH_STATE_COOKIE, JSON.stringify({ state, redirectTo }), {
    httpOnly: true,
    secure: config.nodeEnv === "production",
    sameSite: "lax",
    maxAge: 10 * 60 * 1000,
  });
  return state;
}

function consumeOAuthState(req: import("express").Request, res: import("express").Response, receivedState: unknown): string {
  const raw = req.cookies?.[OAUTH_STATE_COOKIE];
  res.clearCookie(OAUTH_STATE_COOKIE);
  if (!raw) throw new ApiError(400, "OAUTH_STATE_MISMATCH", "OAuth session expired, please try again");
  const parsed = JSON.parse(raw) as { state: string; redirectTo: string };
  // Prevents a forged callback from logging a victim into an attacker-controlled session (LOGIN-SEC-012 applied to OAuth).
  if (parsed.state !== receivedState) {
    throw new ApiError(400, "OAUTH_STATE_MISMATCH", "OAuth session expired, please try again");
  }
  return parsed.redirectTo;
}

function isAllowedEmailDomain(email: string): boolean {
  if (config.oauth.allowedEmailDomains.length === 0) return true;
  const domain = email.split("@")[1]?.toLowerCase();
  return config.oauth.allowedEmailDomains.includes(domain ?? "");
}

async function findOrProvisionSsoOperator(email: string, displayName: string): Promise<OperatorRow> {
  const existing = await findOperatorByEmail(email);
  if (existing) return existing;

  const result = await query<OperatorRow>(
    `INSERT INTO operators (email, display_name, status)
     VALUES ($1, $2, 'active')
     RETURNING id, email, display_name, password_hash, status, failed_login_count, locked_until`,
    [email, displayName]
  );
  return result.rows[0]!;
}

sessionRouter.get("/google", (req, res) => {
  const redirectTo = safeRedirectPath(req.query.redirect);
  const state = issueOAuthState(res, redirectTo);
  try {
    res.redirect(getGoogleAuthUrl(state));
  } catch {
    // Not configured — send the browser back to our own login page with a friendly message
    // instead of either a raw JSON error or (worse) Google's own confusing error screen.
    res.redirect("/login?error=google_not_configured");
  }
});

sessionRouter.get("/google/callback", asyncHandler(async (req, res) => {
  const redirectTo = consumeOAuthState(req, res, req.query.state);

  // User denied consent, or Google reported an error — return to login rather than a raw crash.
  if (req.query.error || typeof req.query.code !== "string") {
    return res.redirect("/login?error=google_cancelled");
  }

  try {
    const identity = await resolveGoogleIdentity(req.query.code);
    if (!identity.emailVerified || !isAllowedEmailDomain(identity.email)) {
      await logAuthEvent("oauth_login_failed", identity.email, null, req);
      return res.redirect("/login?error=not_authorized");
    }
    const operator = await findOrProvisionSsoOperator(identity.email, identity.displayName);
    if (operator.status !== "active") {
      return res.redirect("/login?error=account_not_active");
    }
    const { token } = issueSessionToken(operator.id, operator.email);
    setSessionCookie(res, token);
    await logAuthEvent("oauth_login_success", operator.email, operator.id, req);
    res.redirect(redirectTo);
  } catch (err) {
    console.error("Google OAuth callback failed", err);
    res.redirect("/login?error=google_unavailable");
  }
}));

sessionRouter.get("/office365", asyncHandler(async (req, res) => {
  const redirectTo = safeRedirectPath(req.query.redirect);
  const state = issueOAuthState(res, redirectTo);
  try {
    res.redirect(await getMicrosoftAuthUrl(state));
  } catch {
    res.redirect("/login?error=office365_not_configured");
  }
}));

sessionRouter.get("/office365/callback", asyncHandler(async (req, res) => {
  const redirectTo = consumeOAuthState(req, res, req.query.state);

  if (req.query.error || typeof req.query.code !== "string") {
    return res.redirect("/login?error=office365_cancelled");
  }

  try {
    // MFA (LOGIN-O365-007..009) is enforced by Azure AD's own hosted sign-in page during this
    // redirect via Conditional Access — this callback only ever runs after MFA has already
    // succeeded, so there is no separate MFA step to implement here.
    const identity = await resolveMicrosoftIdentity(req.query.code);
    if (!isAllowedEmailDomain(identity.email)) {
      await logAuthEvent("oauth_login_failed", identity.email, null, req);
      return res.redirect("/login?error=not_authorized");
    }
    const operator = await findOrProvisionSsoOperator(identity.email, identity.displayName);
    if (operator.status !== "active") {
      return res.redirect("/login?error=account_not_active");
    }
    const { token } = issueSessionToken(operator.id, operator.email);
    setSessionCookie(res, token);
    await logAuthEvent("oauth_login_success", operator.email, operator.id, req);
    res.redirect(redirectTo);
  } catch (err) {
    console.error("Office 365 OAuth callback failed", err);
    res.redirect("/login?error=office365_unavailable");
  }
}));
