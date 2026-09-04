import { useEffect, useState } from "react";
import { SideNav, type Page } from "./components/layout/SideNav";
import { LoginPage } from "./pages/LoginPage";
import { CloudsPage } from "./pages/CloudsPage";
import { CleaningPage } from "./pages/CleaningPage";
import { ReportsPage } from "./pages/ReportsPage";
import { logout as apiLogout, me, type OperatorSummary } from "./api/auth";

/**
 * No router library in this app — this is a small, manual read/write of window.location covering
 * exactly the app's 3 top-level pages. Each page additionally owns syncing its OWN internal
 * view/tab/detail state into the URL (query params on top of its own path) — see LoginPage's
 * `mode`, CloudsPage's `tab`, CleaningPage's `group`/`view`, and ReportsPage's `operationId`/`page`.
 */
const PAGE_PATH: Record<Page, string> = { clouds: "/clouds", cleaning: "/cleaning", reports: "/reports" };

function pageFromPath(pathname: string): Page {
  if (pathname === "/cleaning") return "cleaning";
  if (pathname === "/reports") return "reports";
  return "clouds";
}

/** Only ever a same-origin path — never follows an absolute/protocol-relative URL some other origin could have set. */
function isSafeRedirectTarget(value: string | null): value is string {
  return !!value && value.startsWith("/") && !value.startsWith("//");
}

// Auto-logout after this long with no mouse/keyboard/touch/scroll activity anywhere in the page —
// a session left open on a shared machine shouldn't stay signed in indefinitely.
const IDLE_LOGOUT_MS = 15 * 60 * 1000;
// A 30s check cadence is more than precise enough for a 15-minute window, and — unlike resetting a
// setTimeout on every single mousemove — this only ever writes a ref, no re-renders, no timer churn.
const IDLE_CHECK_INTERVAL_MS = 30_000;
const ACTIVITY_EVENTS = ["mousedown", "mousemove", "keydown", "scroll", "touchstart", "wheel"] as const;

export default function App() {
  const [operator, setOperator] = useState<OperatorSummary | null>(null);
  const [checkingSession, setCheckingSession] = useState(true);
  const [page, setPage] = useState<Page>(() => pageFromPath(window.location.pathname));
  // Set only by the idle-timeout below, and consumed (reset to null) by the redirect-to-login
  // effect right after — lets that one effect handle every "show /login" case uniformly while
  // still letting LoginPage explain *why*, just for the idle case.
  const [logoutReason, setLogoutReason] = useState<"idle" | null>(null);

  // Restores a session across a page refresh (LOGIN-P-008 / LOGIN-SES-004) — a 401 here just
  // means "not logged in", not an error worth surfacing.
  useEffect(() => {
    me()
      .then(setOperator)
      .catch(() => setOperator(null))
      .finally(() => setCheckingSession(false));
  }, []);

  // Keeps the address bar in sync with the browser's own Back/Forward buttons — pushState alone
  // (in handleNavigate/handleCleanupStarted below) only covers forward navigation done from inside the app.
  useEffect(() => {
    function onPopState() {
      setPage(pageFromPath(window.location.pathname));
    }
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  // While logged out — right after logging out from wherever the user happened to be, an idle
  // timeout, or a deep link hit while never signed in — the address bar should say /login, not
  // whatever page was last showing. Reads the CURRENT location at the moment this fires (not some
  // location captured earlier) so ?redirect= always names wherever the user actually just was.
  useEffect(() => {
    if (checkingSession || operator || window.location.pathname === "/login") return;
    const current = `${window.location.pathname}${window.location.search}`;
    const params = new URLSearchParams();
    if (current && current !== "/") params.set("redirect", current);
    if (logoutReason) params.set("reason", logoutReason);
    const search = params.toString();
    window.history.pushState({}, "", `/login${search ? `?${search}` : ""}`);
    if (logoutReason) setLogoutReason(null); // consumed
  }, [checkingSession, operator, logoutReason]);

  // Auto-logout after IDLE_LOGOUT_MS with no activity anywhere in the window — only while actually
  // signed in, so this never runs on top of the login screen itself.
  useEffect(() => {
    if (!operator) return;
    let lastActivity = Date.now();
    const onActivity = () => {
      lastActivity = Date.now();
    };
    ACTIVITY_EVENTS.forEach((evt) => window.addEventListener(evt, onActivity, { passive: true }));
    const interval = setInterval(() => {
      if (Date.now() - lastActivity >= IDLE_LOGOUT_MS) {
        apiLogout().finally(() => {
          setOperator(null);
          setLogoutReason("idle");
        });
      }
    }, IDLE_CHECK_INTERVAL_MS);
    return () => {
      ACTIVITY_EVENTS.forEach((evt) => window.removeEventListener(evt, onActivity));
      clearInterval(interval);
    };
  }, [operator]);

  function handleLogin(op: OperatorSummary) {
    setOperator(op);
    const redirect = new URLSearchParams(window.location.search).get("redirect");
    const target = isSafeRedirectTarget(redirect) ? redirect : PAGE_PATH.clouds;
    setPage(pageFromPath(target.split("?")[0]!));
    window.history.pushState({}, "", target);
  }

  async function handleLogout() {
    try {
      await apiLogout();
    } finally {
      setOperator(null);
    }
  }

  function handleNavigate(next: Page) {
    setPage(next);
    window.history.pushState({}, "", PAGE_PATH[next]);
  }

  function handleCleanupStarted(operationId: string) {
    setPage("reports");
    // ReportsPage reads `operationId` straight back out of the URL itself on mount — no prop needed.
    window.history.pushState({}, "", `/reports?operationId=${operationId}`);
  }

  if (checkingSession) {
    return <div className="flex min-h-screen items-center justify-center text-sm text-slate-400">Loading…</div>;
  }

  if (!operator) {
    // Passed as a prop (not left for LoginPage to read back out of the URL itself) so there's no
    // race with the effect above, which pushes ?reason=idle only *after* this render has already
    // committed — LoginPage captures it once at mount via a lazy initializer, immune to this
    // prop later flipping back to false once that effect's own setLogoutReason(null) re-renders here.
    return <LoginPage onLogin={handleLogin} idleLoggedOut={logoutReason === "idle"} />;
  }

  return (
    <div className="flex min-h-screen">
      <SideNav active={page} onNavigate={handleNavigate} onLogout={handleLogout} />
      <main className="flex-1 overflow-y-auto">
        {page === "clouds" && <CloudsPage operator={operator} onLogout={handleLogout} />}
        {page === "cleaning" && <CleaningPage onCleanupStarted={handleCleanupStarted} />}
        {page === "reports" && <ReportsPage />}
      </main>
    </div>
  );
}
