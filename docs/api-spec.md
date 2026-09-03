# Backend API Spec

Deliverable 2 of 4 — M365 Data Cleanup Utility

Base path: `/api/v1`. All endpoints require an authenticated CloudFuze operator session
(`Authorization: Bearer <session-jwt>`). Endpoints marked **[admin-role]** additionally require the
caller's role to be `cleanup_admin` for the target tenant, enforced by
[`rbac.ts`](../backend/src/middleware/rbac.ts) — `viewer` role can hit everything except execute.

Every response is JSON. Errors use `{ error: { code, message, details? } }` with a matching HTTP
status. List endpoints are cursor-paginated: `?cursor=<opaque>&limit=<n≤200>`, response includes
`nextCursor: string | null`.

## Auth & tenant connection

### `GET /auth/consent-url?tenantHint=<optional>`
Returns the Microsoft admin-consent URL to hand the customer's Global Admin (see
[azure-ad-app-registration.md](azure-ad-app-registration.md) §4).
→ `{ consentUrl: string }`

### `GET /auth/callback`
OAuth redirect target. Verifies `admin_consent=True` and `tenant` query params, records/updates the
tenant's connection row, kicks off an initial workload discovery job. Redirects the browser to
`/clouds?connected=<tenantId>`.

### `GET /tenants`
List tenants the caller's org can see (self-serve: scoped to tenants the operator connected or has
been granted access to). → `{ tenants: TenantSummary[], nextCursor }`

```ts
type TenantSummary = {
  id: string
  displayName: string
  m365TenantId: string
  status: "connected" | "consent_pending" | "token_error" | "disconnected"
  connectedAt: string        // ISO 8601
  connectedByAdminUpn: string
  lastTokenRefreshAt: string | null
  workloads: ("teams" | "onedrive" | "sharepoint")[]
}
```

### `GET /tenants/:tenantId`
Full detail for the Manage Clouds screen (adds `permissionsGranted: string[]`,
`recentJobs: JobSummary[]`).

### `POST /tenants/:tenantId/disconnect` **[admin-role]**
Stops scheduling jobs, purges cached tokens, sets `status: disconnected`. Does not delete history.
→ `204 No Content`

## Discovery / scope picker

### `GET /tenants/:tenantId/users?search=&limit=`
Typeahead user search (`User.Read.All`) for the picker. → `{ users: { id, displayName, upn }[] }`

### `GET /tenants/:tenantId/teams?search=&limit=`
→ `{ teams: { id, displayName, channelCount, memberCount }[] }`

### `GET /tenants/:tenantId/sites?search=&limit=`
→ `{ sites: { id, webUrl, displayName, storageUsedBytes }[] }`

## Preview (dry-run — no destructive calls)

### `POST /tenants/:tenantId/cleanup/preview`
Body defines the requested scope; server resolves it against live Graph data (or last-cached
discovery snapshot, max 15 min old) and returns counts/sizes only. **Never calls a delete
endpoint.**

```ts
// Request
type PreviewRequest = {
  scope: {
    users?: string[]           // user IDs, for OneDrive-by-user
    teams?: { teamId: string; channelIds?: string[] }[]   // omit channelIds = all channels
    sites?: { siteId: string; libraryIds?: string[] }[]
    removeM365Groups?: boolean // remove the Group behind a fully-cleaned Team
    fileFilter?: { olderThanCutoff?: string /* ISO date */ }
  }
}

// Response
type PreviewResult = {
  previewId: string            // pass to /confirm; expires in 30 min
  generatedAt: string
  totals: { itemCount: number; totalSizeBytes: number }
  breakdown: {
    teams: { teamId: string; displayName: string; channelsToDelete: number; groupWillBeRemoved: boolean }[]
    onedrive: { userId: string; upn: string; fileCount: number; sizeBytes: number }[]
    sharepoint: { siteId: string; displayName: string; libraryCount: number; sizeBytes: number }[]
    chats: { userId: string; upn: string; chatCount: number; approxSizeBytes: number; note: "report-only, not deletable in v1" }[]
  }
  warnings: string[]           // e.g. throttling risk, items outside cutoff, permission gaps
}
```

### `GET /tenants/:tenantId/cleanup/preview/:previewId`
Re-fetch a still-valid preview (for the confirmation screen after navigation).

## Confirm & execute

### `POST /tenants/:tenantId/cleanup/confirm` **[admin-role]**
Requires the typed confirmation phrase and, per your decision that backup export is always
mandatory, always produces an export manifest before any delete call is queued.

```ts
type ConfirmRequest = {
  previewId: string
  typedConfirmation: string      // must exactly equal "DELETE"
  exportManifestOnly: boolean    // false = also zip file contents to blob storage, not just metadata
}
// Response
type ConfirmResult = {
  jobId: string
  status: "export_in_progress"   // job always starts with export, then moves to queued → running
  auditLogId: string
}
```
Server-side, `typedConfirmation !== "DELETE"` → `400 { error: { code: "CONFIRMATION_MISMATCH" } }`.
This endpoint **enqueues** work; it does not delete synchronously.

### `GET /jobs/:jobId`
Poll for progress (frontend also may subscribe via SSE at `/jobs/:jobId/stream`).

```ts
type JobStatus = {
  jobId: string
  tenantId: string
  status: "export_in_progress" | "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled"
  startedAt: string | null
  finishedAt: string | null
  progress: { total: number; completed: number; failed: number }
  exportManifestUrl: string | null   // signed blob URL, always present once export completes
  currentThrottle: { retryAfterSeconds: number } | null
}
```

### `POST /jobs/:jobId/cancel` **[admin-role]**
Best-effort cancel — in-flight batch calls finish, no new batches start. Already-deleted items are
already deleted; this does not undo them.

## Audit / reports

### `GET /tenants/:tenantId/jobs?status=&from=&to=`
List past cleanup jobs (Reports tab history table).

### `GET /jobs/:jobId/audit`
Per-item audit trail for one job.

```ts
type AuditEntry = {
  itemType: "channel" | "file" | "site_library" | "group"
  itemId: string
  itemPath: string
  sizeBytes: number
  result: "deleted" | "failed" | "skipped"
  errorCode: string | null
  timestamp: string
}
```

### `GET /jobs/:jobId/audit/export?format=csv`
Downloads the full per-item audit log as CSV (also linked from the manifest generated at confirm
time). → `text/csv` stream.

### `GET /tenants/:tenantId/reports/summary?from=&to=`
Aggregate for the Reports dashboard: total bytes reclaimed, items deleted, jobs run, failure rate,
broken down by workload.

## Rate limiting & throttling contract

Every route that calls Graph must respect `Retry-After` on `429`/`503` and apply exponential
backoff with jitter before the next batch (implemented in
[`services/rateLimiter.ts`](../backend/src/services/rateLimiter.ts)). Preview and execute both
report `warnings`/`currentThrottle` back to the client rather than silently stalling — the progress
UI shows "Paused: tenant is being throttled by Microsoft Graph, resuming in Ns" instead of a frozen
bar.

## Roles

| Role | Preview | Confirm/Execute | Cancel | Reports/Audit | Disconnect |
|---|---|---|---|---|---|
| `viewer` | ✅ | ❌ | ❌ | ✅ | ❌ |
| `cleanup_admin` | ✅ | ✅ | ✅ | ✅ | ✅ |

Role is per-tenant (a user can be `viewer` on tenant A and `cleanup_admin` on tenant B) — matches
the "hundreds+ tenants, self-serve" model where different account teams own different customer
engagements.
