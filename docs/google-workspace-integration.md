# Google Workspace Integration — My Drive

This is the first Google workload added to the cleanup tool, built as a vertical slice: full
Add Cloud → Manage Clouds → Sync → Cleaning → Cleanup → Reports lifecycle for **My Drive only**,
proving the pattern before Shared Drives, Google Chat, and Gmail get added in later passes. It is
purely additive — no existing Microsoft 365 (OneDrive/SharePoint/Teams/Outlook) route, worker,
table, or UI component changed behavior; every M365 code path only gained a sibling branch.

## Architecture

A Google Workspace customer gets its own `tenants` row (keyed by a new `google_customer_id`
column), never merged with an M365 tenant row for "the same" real-world company — every tenant is
exactly one directory and exactly one vendor. `connections`, `connection_users`, `sync_jobs`,
`sync_job_resources`, `cleanup_operations`, and `cleanup_operation_items` are all reused as-is —
only their `cloud_type`/`resource_type` CHECK constraints were widened. See
`backend/src/db/migrations/014_google_my_drive.sql`.

The one structurally new piece: Microsoft's app-only client-credentials model gives one shared
Graph client per tenant; Google Workspace domain-wide delegation (DWD) requires impersonating a
*specific user* per Drive/Directory call, so there is no equivalent single shared client — a fresh
impersonated client is built per user, cached by `(subject, scopes)` in
`backend/src/services/googleWorkspaceAuth.ts`.

## OAuth / connection setup — deliberately NOT a redirect flow

Google Workspace access for "read/delete data across every user in a domain" requires **domain-wide
delegation**, not a per-user OAuth consent redirect — a regular OAuth flow can only ever grant
access to the consenting user's own Drive, never "every user in the domain" the way this tool needs
for discovery. Verified against developers.google.com, not guessed.

**Setup, one-time per Google Cloud project (ours, not the customer's):**
1. Create a service account in this app's own GCP project.
2. Download its JSON key and set `GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY_PATH` (local dev) or
   `GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY_JSON` (prod, e.g. from a secrets manager) — exactly one.
   **Deliberately not** `GOOGLE_CLIENT_ID`/`GOOGLE_CLIENT_SECRET`/`GOOGLE_REDIRECT_URI` — those
   names are already used by `backend/src/config/index.ts`'s `oauth.google` for an unrelated
   feature (operator "Sign in with Google" SSO login). Reusing them would either collide with that
   login flow or force sharing one OAuth client across two very different purposes and scope sets.

**Setup, once per customer domain (the customer's Workspace super-admin does this, not us):**
1. Admin Console → Security → API controls → Domain-wide delegation → Add new.
2. Enter this service account's numeric Client ID.
3. Enter the scopes below (comma-separated) → Authorize.

This is an out-of-band admin-console action with no redirect/callback for the app to observe —
`POST /api/google-clouds/connect/verify` (`backend/src/routes/googleConnections.ts`) is how the app
confirms it actually happened: it impersonates the given admin email and makes a real Directory API
call, deriving the tenant's real `customerId`/domain from Google's own response (never trusting the
submitted domain string directly — the same "don't trust the client for identity" principle the
M365 flow applies via the OAuth token's `tid` claim, just via a different mechanism here since
there's no token to read a claim from).

## Required scopes

| Scope | Why |
|---|---|
| `https://www.googleapis.com/auth/admin.directory.user.readonly` | List/look up domain users (Directory API) — the "who has a My Drive to clean up" discovery step. |
| `https://www.googleapis.com/auth/drive` | Full Drive access. `drive.readonly`/`drive.file` were considered and rejected — neither supports deleting arbitrary existing files owned by another user, only `drive` (full) does. |

No admin.reports scope is requested — see "Storage figures" below for why.

## Connection flow

```
Add Cloud → click Google My Drive tile → GoogleConnectDialog (domain + admin email)
    → POST /api/google-clouds/connect/verify
    → verifies domain-wide delegation via a real Directory API call
    → creates/updates tenants + connections (cloud_type='google_my_drive') + tenant_roles
    → queues an initial sync_jobs row
    → Manage Clouds shows the new connection, same polling as the M365 popup flow
```

## Sync

Reuses the exact `sync_jobs`/`sync_job_resources` architecture and `cloudSyncWorker.ts`'s BullMQ
queue — a new branch in that same worker (not a second `Worker` instance on the same queue, which
would load-balance non-deterministically) dispatches to `jobs/googleDriveSync.ts`'s
`syncGoogleMyDrive`. Both full-domain and resource-level ("Sync Selected Users") sync are supported,
identically to OneDrive's resource-level sync. Resync from Manage Clouds works exactly as it does
for any other cloud type (`POST /api/clouds/:id/resync`, unchanged route, new
`cloud_type === "google_my_drive"` branch inside it for browsing/validating specific users).

**Known gap, deferred:** the Cleaning page's Dashboard has its own separate "Sync Now" convenience
button per tile, backed by a different `sync_operations` table with fixed
`onedrive_sync_job_id`/`sharepoint_sync_job_id`/`outlook_sync_job_id`/`teams_scan_id` columns. This
was not extended to add a `google_my_drive_sync_job_id` column in this pass — My Drive syncing is
fully functional via Manage Clouds' per-connection Resync (the same mechanism every other workload
actually uses under the hood), just missing this one inline dashboard-tile shortcut.

## Cleanup

Reuses `cleanup_operations`/`cleanup_operation_items`/`cleanup_operation_item_files` unchanged
(only the `resource_type` CHECK gained `'google_my_drive_account'`) and the same
`CleanupManifest` → `resolveManifestItems` → `POST /cleanup` → `cleanupExecutionWorker.ts` pipeline.
A new sibling module, `jobs/googleDriveCleanupExecution.ts`, executes Google items exactly the way
`graph/cleanupDeletion.ts` executes OneDrive items — list root-level files/folders, seed one
`cleanup_operation_item_files` row per file, delete with the operator's chosen `deletion_mode`.

Google Drive's own two delete calls map directly onto the `recycle_bin`/`permanent` choice already
built for M365:
- `recycle_bin` → `files.update({trashed: true})` — Drive's own Trash, 30-day auto-purge, restorable.
- `permanent` → `files.delete()` — a genuine, immediate permanent delete, skipping Trash entirely.
  Verified against developers.google.com's Drive API v3 reference, not assumed.

Only the user's own top-level My Drive content is ever touched — never the Workspace user account,
never the domain, never the OAuth/delegation grant itself.

## Reports

The CSV report and the progress/results UI's "Permanently Deleted" vs "Removed" labeling reuse the
exact same `deletion_mode`-gated logic already built for M365 — `google_my_drive_account` was simply
added to the `PERMANENT_DELETE_RESOURCE_TYPES` set in `CleanupProgress.tsx`/`CleanupResults.tsx`/
`ItemFilesDrilldown.tsx` and the report's resource-type label map in `routes/cleaning.ts`.

## Storage figures — a Google API limitation, not an app bug

Google exposes per-user Drive storage two ways:
- **Admin SDK Reports API** (`userUsageReport`) — explicitly documents that its data "will not
  include the last 2 days." A hard, worse-than-M365 floor.
- **Drive API `about.get`'s `storageQuota`** — the one this app uses (`getUserDriveUsage` in
  `graph/googleDriveEnumeration.ts`), closer to real-time, but not independently verified against a
  live Workspace domain yet for whether it also lags after a bulk deletion the way Microsoft's
  `drive.quota.used` does (confirmed, via real tenant data, to lag real M365 deletions).

Following the same UX philosophy already built for M365's storage-lag hint (`deletionRecalcHint` in
`routes/cleaning.ts`): show Last Synced, never claim the figure is real-time, and surface an
explanatory note rather than silently showing a possibly-stale number as if it were current.

## What this pass does NOT include

Shared Drives, Google Chat, Gmail, and any Data Dump/synthetic-test-data generator for Google — none
of these exist yet, and Data Dump doesn't exist for M365 either (confirmed by exhaustive search of
this repo), so it would be a net-new subsystem for both ecosystems if built, not a small addition.
Cross-vendor "this M365 tenant and this Google tenant are actually the same real-world customer"
linking is also out of scope — each is tracked as a fully independent `tenants` row today.
