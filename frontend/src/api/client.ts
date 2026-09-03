export class ApiClientError extends Error {
  constructor(
    public status: number,
    public code: string,
    message: string,
    public details?: unknown
  ) {
    super(message);
  }
}

function getCookie(name: string): string | null {
  const match = document.cookie.match(new RegExp(`(?:^|; )${name}=([^;]*)`));
  return match ? decodeURIComponent(match[1]) : null;
}

/**
 * Thin fetch wrapper for the backend API: always sends the session cookie, echoes the CSRF
 * double-submit cookie back as a header on state-changing requests (docs/api-spec.md +
 * backend/src/middleware/csrf.ts), and normalizes error responses into ApiClientError.
 */
export async function apiFetch<T>(path: string, options: RequestInit = {}): Promise<T> {
  const method = (options.method ?? "GET").toUpperCase();
  const headers = new Headers(options.headers);

  if (method !== "GET" && method !== "HEAD") {
    const csrfToken = getCookie("cf_csrf");
    if (csrfToken) headers.set("x-csrf-token", csrfToken);
  }
  if (options.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const res = await fetch(`/api/v1${path}`, { ...options, method, headers, credentials: "include" });

  if (res.status === 204) return undefined as T;

  const data = await res.json().catch(() => null);

  if (!res.ok) {
    const error = data?.error ?? {};
    throw new ApiClientError(res.status, error.code ?? "UNKNOWN_ERROR", error.message ?? "Something went wrong. Please try again.", error.details);
  }

  return data as T;
}
