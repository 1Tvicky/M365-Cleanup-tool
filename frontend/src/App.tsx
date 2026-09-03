import { useEffect, useState } from "react";
import { SideNav, type Page } from "./components/layout/SideNav";
import { LoginPage } from "./pages/LoginPage";
import { CloudsPage } from "./pages/CloudsPage";
import { CleanupPage } from "./pages/CleanupPage";
import { ReportsPage } from "./pages/ReportsPage";
import { logout as apiLogout, me, type OperatorSummary } from "./api/auth";

export default function App() {
  const [operator, setOperator] = useState<OperatorSummary | null>(null);
  const [checkingSession, setCheckingSession] = useState(true);
  const [page, setPage] = useState<Page>("clouds");

  // Restores a session across a page refresh (LOGIN-P-008 / LOGIN-SES-004) — a 401 here just
  // means "not logged in", not an error worth surfacing.
  useEffect(() => {
    me()
      .then(setOperator)
      .catch(() => setOperator(null))
      .finally(() => setCheckingSession(false));
  }, []);

  async function handleLogout() {
    try {
      await apiLogout();
    } finally {
      setOperator(null);
    }
  }

  if (checkingSession) {
    return <div className="flex min-h-screen items-center justify-center text-sm text-slate-400">Loading…</div>;
  }

  if (!operator) {
    return <LoginPage onLogin={setOperator} />;
  }

  return (
    <div className="flex min-h-screen">
      <SideNav active={page} onNavigate={setPage} />
      <main className="flex-1 overflow-y-auto">
        {page === "clouds" && <CloudsPage operator={operator} onLogout={handleLogout} />}
        {page === "cleanup" && <CleanupPage />}
        {page === "reports" && <ReportsPage />}
      </main>
    </div>
  );
}
