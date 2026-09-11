# CloudFuze Cleanup Utility

A guided, audited way to find and permanently remove data left over after a migration into
**Microsoft 365** (OneDrive, SharePoint, Teams, Outlook) or **Google Workspace** (My Drive, Shared
Drives, Gmail, Chat) — without guessing what's safe to delete or losing the record of what
happened.

When a company migrates from another platform into M365 or Google Workspace, the migration tool
that got the data there usually leaves the original, now-duplicate copy behind. This app is the
tool an internal CloudFuze operator uses to go back in, review exactly what's there, and delete it
— per resource, per tenant, with every deletion backed by an export and an audit trail.

## Core concepts

Five ideas recur through the codebase, and keeping them distinct is what makes it safe to operate
on a customer's tenant:

| Concept | Meaning |
|---|---|
| **Tenant / domain** | One customer's Microsoft 365 directory or Google Workspace domain. |
| **Connection** | One workload (e.g. OneDrive, Gmail) linked to one tenant, authorized once via Microsoft admin consent or Google sign-in. The trust boundary — everything below it operates within the scope this defines. |
| **Sync** | Refreshing this app's own record of what a workload contains, for an operator-chosen set of resources — never the whole tenant by default. |
| **Cleaning** | Browsing synced resources and selecting which ones' data should be deleted. |
| **Cleanup** | The confirmed, executed deletion run — item by item, never one irreversible batch call. |
| **Reports** | The permanent, per-item audit trail of every cleanup and sync run, independent of whether the app is still connected. |

A Microsoft 365 tenant and a Google Workspace domain for the same customer are always **two
separate connections** — never merged, even when they share a company name.

## Supported workloads

| Provider | Workload | What a "cleanup" deletes |
|---|---|---|
| Microsoft 365 | OneDrive | A user's root Drive content (never the account) |
| Microsoft 365 | SharePoint | A site's content (never the site) |
| Microsoft 365 | Teams — Team | The whole Team **and every channel in it** |
| Microsoft 365 | Teams — Channel | Just that one channel (parent Team untouched) |
| Microsoft 365 | Teams — 1:1/group chat | *Not supported* — Microsoft has no unattended delete path for chat messages |
| Microsoft 365 | Outlook | Mail, calendar events, or contacts for a mailbox |
| Google Workspace | My Drive | A user's root Drive content (never the account) |
| Google Workspace | Shared Drives | A drive's content (never the drive) |
| Google Workspace | Gmail | Every message in a mailbox (never the mailbox) |
| Google Workspace | Chat | The **entire Space** — messages and memberships together |

Team, Channel, and Chat Space deletion are all immediate and permanent — Microsoft and Google
expose no recoverable/recycle-bin option for any of the three, unlike OneDrive/SharePoint/Outlook
mail, which offer a recycle-bin-vs-permanent choice at cleanup time.

## Architecture

```
frontend (React + Vite + Tailwind)
   │  REST, session-cookie auth
   ▼
backend (Express + TypeScript)
   ├─ routes/        one router per concern (auth, cloudConnections, cleaning, cleanup, reports…)
   ├─ services/      auth/token exchange, encryption, Google Workspace impersonation
   ├─ graph/         Microsoft Graph + Google API calls — enumeration and deletion, one file per workload
   ├─ jobs/          BullMQ workers: cleaning sync, cleanup execution, message-count scans
   ├─ db/            Postgres schema + incremental migrations
   └─ middleware/    session auth, per-tenant RBAC
   │
   ├─ Postgres   tenants, connections, resources, cleanup_operations, reports (source of truth)
   └─ Redis      BullMQ job queues, Graph token/response cache
```

Every request re-derives tenant/workload scope from the connection record server-side — the
frontend's own claims about tenant, cloud type, or resource IDs are never trusted. See
[`docs/`](docs/) for the full guardrail list and API spec.

## Tech stack

- **Backend:** Node.js, TypeScript, Express, PostgreSQL (`pg`), Redis + BullMQ, Zod for
  validation, `@microsoft/microsoft-graph-client` + `@azure/msal-node`, `googleapis` +
  `google-auth-library`.
- **Frontend:** React 18, Vite, TypeScript, Tailwind CSS — no additional UI framework.
- **Testing:** Vitest (backend).

## Getting started

### Prerequisites
- Node.js, PostgreSQL, Redis (a bundled dev copy of both lives under `devservices/` for local use).
- A Microsoft Entra app registration and a Google Cloud service account — see
  [`docs/azure-ad-app-registration.md`](docs/azure-ad-app-registration.md) and
  [`docs/google-workspace-integration.md`](docs/google-workspace-integration.md). Not required just
  to browse the UI with seed data, only to connect a real tenant.

### Backend
```bash
cd backend
cp .env.example .env        # fill in DB/Redis URLs at minimum; see comments for the rest
psql -U postgres -h localhost -d m365_cleanup -f src/db/schema.sql
for f in src/db/migrations/*.sql; do psql -U postgres -h localhost -d m365_cleanup -f "$f"; done
psql -U postgres -h localhost -d m365_cleanup -f scripts/seed-dev.sql   # optional demo operator
npm install
npm run dev                 # http://localhost:4000
```

The seed script creates a local-only demo login: `demo@cloudfuze.com` / `Demo@12345` (see the
script for details — never used outside local development).

### Frontend
```bash
cd frontend
npm install
npm run dev                 # http://localhost:5173
```

### Tests
```bash
cd backend && npm test
cd frontend && npx tsc --noEmit && npx vite build   # type-check + build; no frontend test suite yet
```

## Documentation

- [`docs/api-spec.md`](docs/api-spec.md) — full backend endpoint reference
- [`docs/cloud-connections-api.md`](docs/cloud-connections-api.md) — connect-flow details per workload
- [`docs/azure-ad-app-registration.md`](docs/azure-ad-app-registration.md) — Microsoft Entra app setup and permissions
- [`docs/google-workspace-integration.md`](docs/google-workspace-integration.md) — Google service account + domain-wide delegation setup
- [`docs/graph-api-limitations.md`](docs/graph-api-limitations.md) — known Microsoft Graph constraints (e.g. why chat messages are unsupported)
- [`docs/rollback-safety.md`](docs/rollback-safety.md) — the mandatory pre-delete export mechanism
