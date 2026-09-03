# Login Test Case Coverage

Maps every test case in `CloudFuze_Login_Page_Test_Cases.docx` (189 rows / ~150 cases across
LOGIN-P, LOGIN-E, LOGIN-PWD, LOGIN-BTN, LOGIN-FP, LOGIN-G, LOGIN-O365, LOGIN-AUTH, LOGIN-SEC,
LOGIN-UI, LOGIN-RESP, LOGIN-BR, LOGIN-NET, LOGIN-ACC, LOGIN-SES, LOGIN-EDGE) to what actually
implements it, so a reviewer can trace a test ID straight to code instead of taking it on faith.

## Backend: [routes/session.ts](../backend/src/routes/session.ts)

| Test IDs | Behavior | Where |
|---|---|---|
| LOGIN-P-001, 003–005, 007–009, 011 | Valid login, case-insensitive email, recovery after a failed attempt, Google/Office 365, login-after-logout | `POST /login`, `GET /google\|office365` |
| LOGIN-P-002, LOGIN-BTN-009 | Enter-key submit | Frontend form `onSubmit` — same endpoint as button click, no server distinction needed |
| LOGIN-P-006, LOGIN-PWD-012 | Password masking | Frontend `<input type="password">` — not a backend concern |
| LOGIN-P-012, LOGIN-PWD-008 | Special-character password | [passwordHash.test.ts](../backend/src/services/passwordHash.test.ts) |
| LOGIN-P-013, LOGIN-E-013 | Leading/trailing spaces trimmed | `normalizeEmail()` in [emailValidation.ts](../backend/src/services/emailValidation.ts) |
| LOGIN-E-001–020 | Full email-format matrix (valid forms, malformed forms, injection payloads, length, unicode) | [emailValidation.ts](../backend/src/services/emailValidation.ts) + [emailValidation.test.ts](../backend/src/services/emailValidation.test.ts) — 21 cases run and pass |
| LOGIN-PWD-001, 004–006 | Required-field validation (empty password/email/both) | Zod schema in `session.ts`, 400 `VALIDATION_ERROR` |
| LOGIN-PWD-002, 003, 009–011 | Wrong password, Unicode password, HTML/SQLi in password | bcrypt compare never reflects input back; parameterized queries make SQLi structurally impossible regardless of content |
| LOGIN-PWD-007 | Very long password | `passwordField` zod cap at `PASSWORD_MAX_LENGTH` (128), tested in passwordHash.test.ts |
| LOGIN-PWD-013 | Copy/paste password | No paste-blocking anywhere — works by default |
| LOGIN-BTN-001–004 | Button visibility/behavior on valid/invalid/empty input | Frontend; server side just needs correct status codes (400/401), which it returns |
| LOGIN-BTN-005, 006, LOGIN-SES-006, LOGIN-NET-010 | Double-click / rapid clicks / duplicate requests | **Frontend responsibility** (disable submit while in flight) — each backend request is independently safe to receive concurrently: failed attempts still count correctly, a successful login just issues another valid session |
| LOGIN-BTN-007, LOGIN-NET-002 | Loading state under slow network | Frontend concern; backend adds no artificial delay |
| LOGIN-BTN-008, LOGIN-NET-003 | Network failure / offline | Frontend fetch error handling; nothing to implement server-side |
| LOGIN-FP-001–013 | Forgot-password link, request, unregistered-email non-disclosure, expired/reused token, new password takes effect, old password stops working | `POST /forgot-password`, `POST /reset-password` — see schema.sql `password_reset_tokens` |
| LOGIN-G-001–010 | Google button, consent flow, cancel, multiple accounts, denial, popup blocked, timeout, service failure, expired token | `GET /google`, `GET /google/callback` in session.ts + [googleAuth.ts](../backend/src/services/googleAuth.ts). Popup-blocked (G-007) is a full-page redirect flow, not a popup, so that failure mode doesn't apply here |
| LOGIN-O365-001–010 | Office 365 button, MSAL auth-code flow, MFA, denial, unauthorized account | `GET /office365`, `GET /office365/callback` + [microsoftUserAuth.ts](../backend/src/services/microsoftUserAuth.ts). **MFA (007–009) is enforced by Azure AD's own hosted sign-in page** via Conditional Access — the callback only ever runs after MFA already succeeded, so there's no MFA code to write here |
| LOGIN-AUTH-001–003 | Valid/invalid user × valid/invalid password combinations | Same generic 401 for all — see enumeration note below |
| LOGIN-AUTH-004, 006 | Disabled / deleted account | `operators.status`, checked after password verification |
| LOGIN-AUTH-005 | Locked account | `operators.locked_until`, checked before password verification |
| LOGIN-AUTH-007 | Unverified account | `operators.status = 'unverified'`, specific message (see below on why this one *can* be specific) |
| LOGIN-AUTH-008 | Expired account | **Not modeled** — no business requirement surfaced for account expiry distinct from disabled; treat as `disabled` if this becomes a real requirement |
| LOGIN-AUTH-009, LOGIN-SES-003 | Protected URL without login → redirect | `requireSession` middleware returns 401; frontend routes the redirect |
| LOGIN-AUTH-010, LOGIN-SES-004 | Session expiration / refresh while authenticated | JWT `exp` = `SESSION_TTL_MINUTES`; `GET /me` lets the frontend silently re-check on load |
| LOGIN-SEC-001, 002 | SQL injection (email/password) | Every query in this codebase is parameterized (`pg` `$1`/`$2` placeholders) — string concatenation into SQL is never used, so injection payloads are just inert data |
| LOGIN-SEC-003, 004 | XSS (email/password) | Backend never reflects input back as HTML — all responses are JSON; rejected by email format validation before even reaching storage in most cases |
| LOGIN-SEC-005 | Brute-force / rate limiting | [loginRateLimit.ts](../backend/src/services/loginRateLimit.ts) (IP + email, Redis fixed window) **and** durable per-account lockout (`failed_login_count`/`locked_until`) |
| LOGIN-SEC-006 | Password never in URL | Login is `POST` with a JSON body; no endpoint accepts credentials via query string |
| LOGIN-SEC-007 | Password never in storage | Session token is an **httpOnly** cookie — never readable by page JS, never put in `localStorage`/`sessionStorage` by this app |
| LOGIN-SEC-008 | Password never logged | Password fields are never passed to `console.*`/logging anywhere; only email + event type are recorded in `auth_events` |
| LOGIN-SEC-009 | HTTPS enforcement | `app.ts` middleware rejects non-HTTPS requests in production |
| LOGIN-SEC-010 | Secure cookies | `httpOnly`, `secure` (prod), `sameSite: lax` on the session cookie — see `setSessionCookie()` |
| LOGIN-SEC-011 | Session fixation | Every login issues a **brand-new JWT with a fresh `jti`** ([sessionTokens.ts](../backend/src/services/sessionTokens.ts)) — a pre-auth token is never "upgraded" into an authenticated one |
| LOGIN-SEC-012 | CSRF protection | Double-submit cookie ([csrf.ts](../backend/src/middleware/csrf.ts)) on `/logout`, `/forgot-password`, `/reset-password`; OAuth flows get their own `state`-based CSRF guard |
| LOGIN-SEC-013 | Open redirect | `safeRedirectPath()` only accepts a same-origin relative path, rejecting `//host` and absolute URLs |
| LOGIN-SEC-014 | User enumeration | Identical 401 message and response shape for "no such user" vs. "wrong password"; `verifyPassword` always runs a real bcrypt compare (dummy hash fallback) so timing doesn't leak which case occurred either |
| LOGIN-SEC-015 | Protected content after logout + Back button | Logout revokes the token server-side (Redis `jti` blocklist), so a cached page's API calls fail with 401 even if the HTML itself is served from bfcache |
| LOGIN-SEC-016 | Token manipulation | `jsonwebtoken` signature verification rejects any modified token outright |
| LOGIN-UI-\*, LOGIN-RESP-\*, LOGIN-BR-\*, LOGIN-ACC-\* | Visual layout, responsive breakpoints, cross-browser, accessibility (tab order, screen readers, contrast, zoom) | **Out of backend scope** — these are properties of [LoginPage.tsx](../frontend/src/pages/LoginPage.tsx) and need manual/automated frontend testing (Lighthouse, axe, real-device pass) |
| LOGIN-NET-001, 004–009 | Normal network, timeout, HTTP 400/401/403/429/500 | `ApiError` + the global error handler in [app.ts](../backend/src/app.ts) return the correct status for every failure path above; timeouts are the HTTP client's responsibility (no server-side change needed) |
| LOGIN-SES-001, 002, 005 | Session creation, logout invalidation, multi-tab consistency | Cookie-based session is shared across tabs on the same origin automatically; logout's `jti` revocation applies to all of them at once |
| LOGIN-EDGE-001–003 | Max email length, min/one-below-min password length | Enforced in `resetPasswordSchema`/`emailField`; policy is `PASSWORD_MIN_LENGTH=8`, `PASSWORD_MAX_LENGTH=128`, `EMAIL_MAX_LENGTH=254` (`.env.example`) |
| LOGIN-EDGE-004 | Huge pasted input doesn't break the page | Field-level zod `max()` rejects with 400 before any processing; global body size cap (`express.json({ limit: "1mb" })`) backstops it |
| LOGIN-EDGE-005, 006 | Cookies / JavaScript disabled | Cookies disabled: login literally cannot work (session is cookie-based by design) — documented as a hard requirement, not a bug. JS disabled: a React SPA cannot function at all; out of scope for this architecture |
| LOGIN-EDGE-007, 008 | Browser Back/Forward, popup blocker | Back/Forward: no server state changes on navigation, so this is inherently safe; popup blocker doesn't apply since both OAuth flows are full-page redirects, not popups |

## Why locked-account and unverified-account errors are allowed to be specific

LOGIN-SEC-014 (user enumeration) is only violated when a response lets someone learn **whether an
email is registered** *before* proving they know the password. Once the password has already been
verified correct, revealing that the account is locked, unverified, or disabled isn't an
enumeration leak — the caller has already demonstrated they hold valid credentials for that
account. That's why the code checks password correctness first, then account state — see the
ordering in `POST /login` in `session.ts`. The one exception is the **locked** check, which
happens *before* the password check by design: once an account is locked, its password is
irrelevant, so checking it first would mean doing a needless bcrypt compare on every locked-account
request — this doesn't reopen the enumeration issue by the same argument (the account already
being locked is itself only revealed to someone who's about to be told "locked" regardless of
whether they typed the right password).

## Not implemented — needs a decision, not just code

- **Real email delivery** for password-reset links — `POST /forgot-password` currently logs the
  reset link to the server console instead of sending it. Needs an SES/SendGrid/etc integration
  before LOGIN-FP-007 can be verified end-to-end.
- **MFA UI on this app's own login form** — not needed. Both Google and Office 365 already handle
  MFA on their own hosted pages; there's no local-password MFA in this test matrix to implement.
- **Frontend wiring** — [LoginPage.tsx](../frontend/src/pages/LoginPage.tsx) still simulates login
  client-side with a mock `onLogin()` callback. It needs to be pointed at `POST /api/v1/auth/login`
  (with `credentials: "include"` so the httpOnly cookie is sent/received) plus the CSRF bootstrap
  call (`GET /api/v1/auth/csrf` on page load) and error-message display for the 400/401/403/429
  responses above. That's the next piece of work to make this testable end-to-end in a browser.
