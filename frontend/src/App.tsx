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

// Wherever the browser was pointed when the app booted — captured once, at module load, before
// anything below has a chance to rewrite the URL to /login. Used to send the user back to a deep
// link (e.g. /cleaning?group=...) they hit while logged out, once they actually sign in.
const bootLocation = `${window.location.pathname}${window.location.search}`;

/** Only ever a same-origin path — never follows an absolute/protocol-relative URL some other origin could have set. */
function isSafeRedirectTarget(value: string | null): value is string {
  return !!value && value.startsWith("/") && !value.startsWith("//");
}

export default function App() {
  const [operator, setOperator] = useState<OperatorSummary | null>(null);
  const [checkingSession, setCheckingSession] = useState(true);
  const [page, setPage] = useState<Page>(() => pageFromPath(window.location.pathname));

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

  // While logged out (including right after logging out from some other page), the address bar
  // should say /login, not whatever page happened to be showing — but keep the original target as
  // ?redirect= so handleLogin below can still send the user back there.
  useEffect(() => {
    if (checkingSession || operator || window.location.pathname === "/login") return;
    const url = bootLocation && bootLocation !== "/" ? `/login?redirect=${encodeURIComponent(bootLocation)}` : "/login";
    window.history.pushState({}, "", url);
  }, [checkingSession, operator]);

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
    return <LoginPage onLogin={handleLogin} />;
  }

  return (
    <div className="flex min-h-screen">
      <SideNav active={page} onNavigate={handleNavigate} />
      <main className="flex-1 overflow-y-auto">
        {page === "clouds" && <CloudsPage operator={operator} onLogout={handleLogout} />}
        {page === "cleaning" && <CleaningPage onCleanupStarted={handleCleanupStarted} />}
        {page === "reports" && <ReportsPage />}
      </main>
    </div>
  );
}
