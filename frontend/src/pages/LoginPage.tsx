import { useEffect, useState, type FormEvent } from "react";
import { BrandMark } from "../components/layout/BrandMark";
import {
  GoogleIcon,
  OfficeRibbonIcon,
  OneDriveOutlineIcon,
  SharePointOutlineIcon,
  TeamsOutlineIcon,
} from "../components/clouds/CloudIcons";
import { ApiClientError } from "../api/client";
import { bootstrapCsrf, forgotPassword, googleLoginHref, login, office365LoginHref, type OperatorSummary } from "../api/auth";

const FEATURES = [
  "Delete migrated Teams, OneDrive & SharePoint data safely",
  "Every deletion backed by a mandatory pre-delete export",
  "Full audit trail for every cleanup job, exportable as CSV",
];

// Maps the `?error=` code the backend redirects back with after a Google/Office 365 attempt
// (backend/src/routes/session.ts) to copy a user can actually act on.
const OAUTH_ERROR_MESSAGES: Record<string, string> = {
  google_cancelled: "Google sign-in was cancelled.",
  office365_cancelled: "Office 365 sign-in was cancelled.",
  not_authorized: "That account isn't authorized to sign in here.",
  account_not_active: "This account isn't active yet. Contact your administrator.",
  google_unavailable: "Google sign-in is temporarily unavailable. Try again shortly.",
  office365_unavailable: "Office 365 sign-in is temporarily unavailable. Try again shortly.",
  google_not_configured: "Google sign-in isn't set up for this deployment yet — use your email and password instead.",
  office365_not_configured: "Office 365 sign-in isn't set up for this deployment yet — use your email and password instead.",
};

// Pure read only — no side effects. React 18 StrictMode double-invokes useState's lazy
// initializer in dev; an initializer that also *mutates* location (e.g. via replaceState) would
// have its own first call's cleanup wipe the query string before the second invocation reads it,
// silently losing the error every time. Cleanup happens separately, in an effect (see below).
function readOAuthErrorCode(): string | null {
  return new URLSearchParams(window.location.search).get("error");
}

/** Same "pure, read-only lazy initializer" reasoning as readOAuthErrorCode above. */
function readMode(): "login" | "forgot" {
  return new URLSearchParams(window.location.search).get("mode") === "forgot" ? "forgot" : "login";
}

/**
 * Matches CloudFuze's existing split-panel login screen (layout, colors, form elements), adapted
 * to this tool's scope: copy talks about cleanup, not migration, and the provider row at the
 * bottom of the blue panel only shows the three M365 workloads this tool touches.
 *
 * Wired to the real backend (backend/src/routes/session.ts) — see docs/login-test-case-coverage.md.
 */
export function LoginPage({ onLogin }: { onLogin: (operator: OperatorSummary) => void }) {
  const [mode, setMode] = useState<"login" | "forgot">(readMode);
  const [oauthErrorCode] = useState(readOAuthErrorCode);
  const oauthError = oauthErrorCode ? (OAUTH_ERROR_MESSAGES[oauthErrorCode] ?? "Sign-in didn't complete. Please try again.") : null;

  useEffect(() => {
    if (oauthErrorCode) {
      // Strip just the error param (keep `mode`/anything else) so a refresh doesn't keep
      // re-showing the same error. Runs here (not in the lazy initializer above) so StrictMode's
      // double-invoke can't race the read against it.
      const params = new URLSearchParams(window.location.search);
      params.delete("error");
      const search = params.toString();
      window.history.replaceState({}, "", `${window.location.pathname}${search ? `?${search}` : ""}`);
    }
    // Must land before any POST — the CSRF cookie doesn't exist until this resolves.
    bootstrapCsrf().catch(() => {
      /* login itself will fail loudly if this never lands; nothing useful to show yet */
    });
  }, []);

  // Reflects login/forgot-password in the URL as a `?mode=` param, layered on whatever the current
  // path is (App.tsx already put this on /login, keeping any deep-link target in ?redirect=).
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (mode === "forgot") params.set("mode", "forgot");
    else params.delete("mode");
    const search = params.toString();
    const url = `${window.location.pathname}${search ? `?${search}` : ""}`;
    if (`${window.location.pathname}${window.location.search}` !== url) {
      window.history.pushState({}, "", url);
    }
  }, [mode]);

  useEffect(() => {
    function onPopState() {
      setMode(readMode());
    }
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  return (
    <div className="flex min-h-screen">
      <div className="relative hidden w-1/2 flex-col justify-between overflow-hidden bg-[#1b2fc4] px-16 py-14 text-white xl:flex">
        <div>
          <BrandMark className="h-10 w-[70px] text-white" />

          <div className="mt-20 text-base font-medium text-blue-100">Get Started with CloudFuze</div>
          <h1 className="mt-4 max-w-lg text-4xl font-bold leading-tight">
            Clean up duplicate Microsoft 365 data after migration and reclaim storage — safely.
          </h1>

          <ul className="mt-12 space-y-5">
            {FEATURES.map((f) => (
              <li key={f} className="flex items-start gap-3 text-base text-blue-50">
                <span className="mt-1 text-teal-300" aria-hidden>
                  ➤
                </span>
                {f}
              </li>
            ))}
          </ul>
        </div>

        <div className="flex items-center gap-8 text-white/90">
          <OneDriveOutlineIcon className="h-8 w-8" />
          <SharePointOutlineIcon className="h-8 w-8" />
          <TeamsOutlineIcon className="h-8 w-8" />
        </div>

        <svg className="pointer-events-none absolute -bottom-10 -right-10 h-56 w-56 text-white/10" viewBox="0 0 100 100" fill="currentColor">
          <ellipse cx="55" cy="60" rx="45" ry="30" />
        </svg>
        <svg className="pointer-events-none absolute bottom-24 right-32 h-24 w-24 text-white/10" viewBox="0 0 100 100" fill="currentColor">
          <ellipse cx="50" cy="50" rx="40" ry="26" />
        </svg>
      </div>

      <div className="flex flex-1 flex-col items-center justify-center bg-white px-6 py-12">
        {mode === "login" ? (
          <LoginForm
            onLogin={onLogin}
            onForgotPassword={() => setMode("forgot")}
            initialError={oauthError}
            redirectTo={new URLSearchParams(window.location.search).get("redirect") ?? "/"}
          />
        ) : (
          <ForgotPasswordForm onBackToLogin={() => setMode("login")} />
        )}
      </div>
    </div>
  );
}

function LoginForm({
  onLogin,
  onForgotPassword,
  initialError,
  redirectTo,
}: {
  onLogin: (operator: OperatorSummary) => void;
  onForgotPassword: () => void;
  initialError: string | null;
  redirectTo: string;
}) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(initialError);
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    if (submitting) return; // guards double-click / rapid repeat clicks (LOGIN-BTN-005/006)
    setSubmitting(true);
    setError(null);
    try {
      const operator = await login(email, password);
      onLogin(operator);
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't reach the server. Check your connection and try again.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <form className="w-full max-w-md" onSubmit={handleSubmit}>
      <div className="flex flex-col items-center text-center">
        <BrandMark className="h-8 w-14 text-[#1b2fc4]" />
        <span className="mt-2 text-xl font-bold text-[#1b2fc4]">CloudFuze</span>
        <h2 className="mt-7 text-2xl font-bold text-slate-900">Login to your account</h2>
        <p className="mt-2.5 text-base text-slate-400">Welcome back!</p>
        <p className="text-base text-slate-400">Please login to gain access to CloudFuze</p>
      </div>

      {error && (
        <div role="alert" className="mt-6 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
          {error}
        </div>
      )}

      <div className="mt-6 space-y-4">
        <input
          type="email"
          required
          autoComplete="username"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="name@cloudfuze.com"
          aria-label="Email address"
          className="w-full rounded-lg border border-blue-200 bg-blue-50/60 px-4 py-3.5 text-base text-slate-800 placeholder:text-slate-400 focus:border-[#1b2fc4] focus:outline-none focus:ring-1 focus:ring-[#1b2fc4]"
        />
        <input
          type="password"
          required
          autoComplete="current-password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          placeholder="Password"
          aria-label="Password"
          className="w-full rounded-lg border border-slate-200 bg-slate-50 px-4 py-3.5 text-base text-slate-800 placeholder:text-slate-400 focus:border-[#1b2fc4] focus:outline-none focus:ring-1 focus:ring-[#1b2fc4]"
        />
      </div>

      <div className="mt-3 text-right">
        <button type="button" onClick={onForgotPassword} className="text-sm font-medium text-[#1b2fc4] underline underline-offset-2">
          Forgot password ?
        </button>
      </div>

      <button
        type="submit"
        disabled={submitting}
        className="relative mt-6 flex w-full items-center justify-center rounded-lg bg-[#131f6e] py-3.5 text-base font-semibold text-white transition-colors hover:bg-[#0e1857] disabled:cursor-not-allowed disabled:opacity-70"
      >
        {submitting ? "Logging in…" : "Login"}
        {!submitting && (
          <span className="absolute right-5" aria-hidden>
            →
          </span>
        )}
      </button>

      <div className="mt-5 flex gap-3">
        <a
          href={googleLoginHref(redirectTo)}
          className="flex flex-1 items-center overflow-hidden rounded-lg border border-slate-200 bg-[#4285F4] text-sm font-medium text-white hover:brightness-105"
        >
          <span className="flex h-full items-center bg-white px-2.5 py-2.5">
            <GoogleIcon className="h-5 w-5" />
          </span>
          <span className="flex-1 text-center">Sign in with Google</span>
        </a>
        <a
          href={office365LoginHref(redirectTo)}
          className="flex flex-1 items-center justify-between rounded-lg border border-slate-300 bg-white px-4 py-3 text-sm font-medium text-slate-700 hover:bg-slate-50"
        >
          Office 365
          <OfficeRibbonIcon className="h-5 w-5" />
        </a>
      </div>

      <div className="mt-7 border-t border-slate-100 pt-6 text-center text-sm text-slate-500">
        Create an account for free ?{" "}
        <button type="button" className="font-semibold text-[#1b2fc4]">
          Sign up
        </button>
      </div>
    </form>
  );
}

function ForgotPasswordForm({ onBackToLogin }: { onBackToLogin: () => void }) {
  const [email, setEmail] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Always the same message on success, whether or not the email is registered (LOGIN-FP-004).
  const [sent, setSent] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    if (submitting) return;
    setSubmitting(true);
    setError(null);
    try {
      await forgotPassword(email);
      setSent(true);
    } catch (err) {
      setError(err instanceof ApiClientError ? err.message : "Couldn't reach the server. Check your connection and try again.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="w-full max-w-md">
      <div className="flex flex-col items-center text-center">
        <BrandMark className="h-8 w-14 text-[#1b2fc4]" />
        <span className="mt-2 text-xl font-bold text-[#1b2fc4]">CloudFuze</span>
        <h2 className="mt-7 text-2xl font-bold text-slate-900">Reset your password</h2>
        <p className="mt-2.5 text-base text-slate-400">
          {sent ? "Check your inbox for next steps." : "Enter your email and we'll send you a reset link."}
        </p>
      </div>

      {sent ? (
        <div className="mt-6 rounded-lg border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-700">
          If an account exists for that email, a reset link has been sent.
        </div>
      ) : (
        <form onSubmit={handleSubmit}>
          {error && (
            <div role="alert" className="mt-6 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
              {error}
            </div>
          )}
          <input
            type="email"
            required
            autoComplete="username"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="name@cloudfuze.com"
            aria-label="Email address"
            className="mt-6 w-full rounded-lg border border-blue-200 bg-blue-50/60 px-4 py-3.5 text-base text-slate-800 placeholder:text-slate-400 focus:border-[#1b2fc4] focus:outline-none focus:ring-1 focus:ring-[#1b2fc4]"
          />
          <button
            type="submit"
            disabled={submitting}
            className="mt-6 flex w-full items-center justify-center rounded-lg bg-[#131f6e] py-3.5 text-base font-semibold text-white transition-colors hover:bg-[#0e1857] disabled:cursor-not-allowed disabled:opacity-70"
          >
            {submitting ? "Sending…" : "Send reset link"}
          </button>
        </form>
      )}

      <div className="mt-7 border-t border-slate-100 pt-6 text-center text-sm text-slate-500">
        <button type="button" onClick={onBackToLogin} className="font-semibold text-[#1b2fc4]">
          ← Back to login
        </button>
      </div>
    </div>
  );
}
