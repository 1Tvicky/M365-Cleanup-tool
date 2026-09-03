import { apiFetch } from "./client";

export interface OperatorSummary {
  id: string;
  email: string;
  displayName: string;
}

/** Bootstraps the CSRF cookie — call once before the login form (or any state-changing call) can submit. */
export function bootstrapCsrf(): Promise<void> {
  return apiFetch<void>("/auth/csrf");
}

export function login(email: string, password: string): Promise<OperatorSummary> {
  return apiFetch<OperatorSummary>("/auth/login", { method: "POST", body: JSON.stringify({ email, password }) });
}

export function logout(): Promise<void> {
  return apiFetch<void>("/auth/logout", { method: "POST" });
}

/** Restores a session across a page refresh (LOGIN-P-008 / LOGIN-SES-004) — 401 means "not logged in", not an error to surface. */
export function me(): Promise<OperatorSummary> {
  return apiFetch<OperatorSummary>("/auth/me");
}

export function forgotPassword(email: string): Promise<{ message: string }> {
  return apiFetch("/auth/forgot-password", { method: "POST", body: JSON.stringify({ email }) });
}

export function resetPassword(token: string, newPassword: string): Promise<{ message: string }> {
  return apiFetch("/auth/reset-password", { method: "POST", body: JSON.stringify({ token, newPassword }) });
}

// Full-page redirects (OAuth), not fetch calls — the browser needs to actually navigate to Google/Microsoft.
export function googleLoginHref(redirectTo = "/"): string {
  return `/api/v1/auth/google?redirect=${encodeURIComponent(redirectTo)}`;
}

export function office365LoginHref(redirectTo = "/"): string {
  return `/api/v1/auth/office365?redirect=${encodeURIComponent(redirectTo)}`;
}
