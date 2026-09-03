# Add Clouds / Manage Clouds — API Spec

Deliverable 2 — connection layer only (OAuth connect, enumeration, resync, disconnect). The
Cleanup (deletion) module is a separate, later pass and isn't touched here.

Base paths: most routes live under `/api/clouds`; the OAuth callback is a fixed top-level path
because it must exactly match the Azure AD app registration's redirect URI (see
`docs/azure-ad-app-registration.md` §4a) — it is **not** versioned under `/api/v1`.

All `/api/clouds/*` routes require an authenticated CloudFuze operator session (same session cookie
as the rest of the app — `middleware/auth.ts`). Routes marked **[internal-admin]** additionally
require `operators.is_internal_admin = true` (`middleware/internalAdmin.ts`) — this is a **global**
flag, deliberately separate from the existing per-tenant `tenant_roles.role` used by the Cleanup
module. Connecting or disconnecting a cloud is a tenant-access-granting action; running a cleanup
job inside an already-connected tenant is a different, narrower kind of trust — see
`docs/azure-ad-app-registration.md` §5 note and the migration comment on `operators.is_internal_admin`.

```ts
type CloudType = "onedrive" | "sharepoint" | "teams";
type ConnectionStatus = "connecting" | "active" | "error" | "needs_reauth" | "disconnected";
type SyncJobStatus = "queued" | "running" | "completed" | "completed_with_errors" | "failed" | "cancelled";
```

## `POST /api/clouds/:cloudType/connect/init` **[internal-admin]**

Starts the OAuth popup flow for one tile. `cloudType` is `onedrive | sharepoint | teams`.

```ts
// Response
interface ConnectInitResponse {
  authorizeUrl: string;   // open this in a popup window
  state: string;          // opaque — the frontend doesn't need to inspect it, just round-trip via the popup navigation
}
```

400 `INVALID_CLOUD_TYPE` if the path param isn't one of the three values.

## `GET /api/auth/m365/callback`

Not called directly by the frontend — this is where Microsoft redirects the popup after the admin
signs in (and consents, if this is the tenant's first connection). See
`docs/azure-ad-app-registration.md` §4a for the full flow.

Query params: `code`, `state` on success; `error`, `error_description` if the admin declined.

Responds with a minimal HTML page (not JSON) that posts a message to `window.opener` and closes
itself:

```ts
// window.opener postMessage payload
type M365ConnectMessage =
  | { type: "m365-connect-complete"; status: "success"; connectionId: string; cloudType: CloudType }
  | { type: "m365-connect-complete"; status: "error"; cloudType: CloudType | null; reason: string };
```

The frontend's popup-opener code should listen for `message` events with `data.type ===
"m365-connect-complete"`, close/cleanup its reference to the popup, and on `status: "success"`
refetch `GET /api/clouds/manage`.

## `GET /api/clouds/manage`

The Manage Clouds list — one row per `(tenant, cloud_type)` connection the operator's tenant
access covers (same tenant-scoping as the rest of the app, via `tenant_roles`).

```ts
interface ManageCloudsRow {
  id: string;              // connections.id
  cloudType: CloudType;
  iconKey: CloudType;       // same value as cloudType — named separately because the frontend's icon lookup is keyed by it explicitly
  displayName: string;      // tenant domain / friendly name
  adminEmail: string;
  totalUsers: number;       // from the connection's latest sync_jobs row
  processedUsers: number;
  percent: number;          // 0-100, derived (processedUsers / totalUsers), 0 if totalUsers is 0
  status: ConnectionStatus;
  multiUser: true;          // always true today — reserved for a future single-user connection mode
  connectedAt: string | null;   // ISO 8601
  lastSyncedAt: string | null;
  lastError: string | null;     // populated when status is 'error' or 'needs_reauth'
}

// Response
interface ManageCloudsResponse {
  connections: ManageCloudsRow[];
}
```

`status: "needs_reauth"` (requirement #5) is surfaced specifically when the enumeration job hits a
Graph auth failure that looks like expired/revoked consent (401 `InvalidAuthenticationToken`,
`consent_required`, etc.) — see the error classification in `jobs/cloudSyncWorker.ts`. The frontend
should render that distinctly from a generic `"error"` (e.g. "Needs reauthorization" with a
re-connect action) rather than looking like a stalled progress bar.

## `GET /api/clouds/:id/users?cursor=&limit=`

Backs the expand-chevron row detail. `limit` defaults to 50, capped at 200.

```ts
interface ConnectionUserRow {
  id: string;
  graphUserId: string;
  upn: string;                 // for cloud_type = 'sharepoint', this is the site's webUrl, not a person — see docs/graph-api-limitations.md
  displayName: string | null;  // site display name for SharePoint connections
  storageUsedBytes: number;
  itemCount: number;
  syncStatus: "pending" | "synced" | "failed";
  lastSyncedAt: string | null;
  errorMessage: string | null;
}

// Response
interface ConnectionUsersResponse {
  users: ConnectionUserRow[];
  nextCursor: string | null;
}
```

404 `CONNECTION_NOT_FOUND` if `:id` doesn't exist or the operator's tenant access doesn't cover it.

## `POST /api/clouds/:id/resync` **[internal-admin]**

Re-enqueues enumeration. Creates a **new** `sync_jobs` row (keeping prior sync history rather than
overwriting it) with `status: "queued"`; `connections.last_synced_at` updates when this new job
*completes*, not at request time — see the reasoning note in `routes/cloudConnections.ts` (the
original ask was to update it at resync-request time, which reads as stale/inaccurate the moment a
resync starts, so completion time is used instead).

```ts
// Response
interface ResyncResponse {
  jobId: string;
  status: "queued";
}
```

409 `RESYNC_ALREADY_RUNNING` if the connection already has a `sync_jobs` row in `queued` or
`running` status. 409 `CONNECTION_DISCONNECTED` if the connection's status is `disconnected` —
reconnect via `connect/init` instead.

## `DELETE /api/clouds/:id` **[internal-admin]**

Soft-disconnects — see `docs/azure-ad-app-registration.md` §6 for exactly what this can and can't
do (it cannot revoke the tenant's Azure AD consent grant, only our local token/job state).

204 No Content on success. 404 `CONNECTION_NOT_FOUND` if unknown/out of scope for the operator.

## Roles

| Route | Any authenticated operator (with tenant access) | `is_internal_admin` |
|---|---|---|
| `GET /api/clouds/manage` | ✅ | — |
| `GET /api/clouds/:id/users` | ✅ | — |
| `POST /api/clouds/:cloudType/connect/init` | ❌ | ✅ |
| `POST /api/clouds/:id/resync` | ❌ | ✅ |
| `DELETE /api/clouds/:id` | ❌ | ✅ |

## Error shape

Consistent with the rest of the app (`app.ts` error handler):
`{ "error": { "code": string, "message": string, "details"?: unknown } }`.
