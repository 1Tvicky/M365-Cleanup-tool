# Azure AD App Registration Checklist

Deliverable 1 of 4 — M365 Data Cleanup Utility (Phase 1: Teams, OneDrive, SharePoint)

This app is multi-tenant (hundreds+ customer tenants, self-serve admin consent), so it must be
registered as a **multi-tenant** Azure AD application in CloudFuze's own Azure AD, and each
customer tenant admin consents to it independently — CloudFuze never has standing access to a
customer tenant until that tenant's Global Admin (or Privileged Role Admin) explicitly grants it.

## 1. Create the app registration

- Azure Portal → Azure Active Directory → App registrations → New registration
- Name: `CloudFuze M365 Data Cleanup Utility` (add `-staging` / `-dev` suffix for non-prod copies —
  never share one app registration across environments)
- Supported account types: **Accounts in any organizational directory (Any Azure AD directory - Multitenant)**
- Redirect URIs (Web) — both point at the same app registration; register **both**, they serve
  different flows:
  - `https://<app-host>/api/v1/auth/callback` — legacy tenant-wide consent flow (§4, superseded by
    the flow below for new connections, kept live for compatibility)
  - `https://<app-host>/api/auth/m365/callback` — the **Add Clouds** per-cloud-type connect flow
    (§4a) — must be an exact string match; if `<app-host>` changes between environments, update
    this here *and* in `OPERATOR_AZURE_REDIRECT_URI`-style env config, both have to match exactly
    what's registered
- Do **not** enable the "Public client" / mobile & desktop flow — this is a confidential client
  (server-side token exchange only)

## 2. Certificates & secrets

- Prefer a **client certificate** over a client secret for the confidential-client credential
  (longer rotation window, no plaintext secret to leak). If a secret is used instead, set expiry
  to ≤ 6 months and put rotation on a calendar reminder — an expired secret silently breaks every
  connected tenant's background jobs.
- Store the secret/certificate only in the deployment's secret manager (Azure Key Vault /
  equivalent). Never commit it, never put it in an env file that reaches source control.

## 3. API permissions (application, least privilege)

All permissions below are **Application** type (not Delegated) because cleanup jobs run as
background batch operations without an interactively signed-in user driving each call. Application
permissions require **tenant admin consent** — this is intentional; only a Global Admin (or
delegated Privileged Role Administrator) of the *customer* tenant can grant them, which is the
authorization gate for the whole tool.

| Permission | Type | Why | Used by |
|---|---|---|---|
| `Team.ReadBasic.All` | Application | Enumerate teams for scope picker | Cleanup → Teams scope |
| `Channel.ReadBasic.All` | Application | List channels per team for the breakdown | Preview |
| `Channel.Delete.All` | Application | Delete specific Teams channels | Execute |
| `Chat.Read.All` | Application | Read chat/DM metadata for the size breakdown **only** — see note below | Preview (report-only) |
| `Group.ReadWrite.All` | Application | Remove M365 Groups tied to deleted Teams | Execute |
| `TeamMember.ReadWrite.All` | Application | Enumerate + manage team membership as part of Teams connection sync | Manage Clouds (Teams) |
| `Files.ReadWrite.All` | Application | Enumerate + delete OneDrive files | Preview, Execute |
| `Sites.ReadWrite.All` | Application | Enumerate + delete SharePoint site content/libraries | Preview, Execute |
| `Sites.FullControl.All` | Application, **only if `Sites.ReadWrite.All` proves insufficient** | Some SharePoint admin-level operations (e.g. certain site-collection settings) need full control; request this only if a specific call 403s under `.ReadWrite.All` — don't request it speculatively, it's a materially heavier ask on the consent screen | Manage Clouds (SharePoint), edge cases |
| `User.Read.All` | Application | Resolve users for the tenant's user/site picker, and to enumerate users for OneDrive/Teams connection sync | Connect, Cleanup |
| `Reports.Read.All` | Application | Pull storage usage reports to size the preview without walking every drive | Preview |

**Not requested in v1:** `Chat.ReadWrite.All` / any chat-delete permission. Per your decision to
defer Teams DM deletion to v2, `Chat.Read.All` is granted for *reporting only* (chat/DM counts and
approximate size shown in the preview breakdown) — the app never calls a chat-message-delete
endpoint in v1. This also sidesteps a real Graph API limitation: application permissions cannot
bulk-delete arbitrary 1:1/group chat messages the way they can channel messages — only the
message's own author (delegated context) or specific compliance/eDiscovery flows can remove chat
content. Re-evaluate this permission set when DM deletion is scoped for v2.

Do **not** request `Directory.ReadWrite.All`, `Mail.*`, or any permission not in the table — Graph
admin-consent screens show the full requested list to the customer's Global Admin, and an
over-broad ask is the #1 reason customers stall or reject consent.

## 4. Admin consent flow — legacy, superseded by §4a

*(Kept live for compatibility with existing connected tenants and the Cleanup module's current
dependency on `tenants` rows. New connections go through §4a instead.)*

1. CloudFuze sends the customer's Global Admin the tenant-specific consent URL:
   `https://login.microsoftonline.com/{customer-tenant-id-or-common}/adminconsent?client_id={app-client-id}&redirect_uri={redirect_uri}`
2. Customer admin reviews the requested application permissions (table above) and clicks **Accept**.
3. Azure AD redirects back to `redirect_uri` with `tenant` and `admin_consent=True` in the query
   string — the backend uses this to record that tenant as consented and to mint the first
   application-token via client-credentials flow (see §5).
4. If the admin declines or closes the window, the backend must treat the tenant as
   **not connected** — do not assume consent from a redirect that lacks `admin_consent=True`.

## 4a. Add Clouds connect flow (per cloud type) — current

This is what `POST /api/clouds/:cloudType/connect/init` and `GET /api/auth/m365/callback`
implement (`backend/src/routes/cloudConnections.ts`). It's a standard authorization-code flow with
PKCE, run in a popup window, rather than the bare `/adminconsent` redirect in §4 — see
`docs/cloud-connections-api.md` for the full request/response contract.

1. Frontend calls `connect/init` for the tile the admin clicked (`onedrive` / `sharepoint` /
   `teams`). The backend generates a PKCE `code_verifier`/`code_challenge` pair, a signed `state`
   (binds the request to `cloudType` + the CloudFuze operator who initiated it, 10-minute expiry),
   stashes the verifier server-side keyed by a nonce in `state` (Redis, 10-minute TTL), and returns
   the Microsoft authorize URL for the frontend to open in a popup:

   ```
   https://login.microsoftonline.com/common/oauth2/v2.0/authorize
     ?client_id={app-client-id}
     &response_type=code
     &redirect_uri={https://<app-host>/api/auth/m365/callback}
     &scope=https://graph.microsoft.com/.default offline_access openid profile
     &prompt=admin_consent
     &state={signed state}
     &code_challenge={challenge}
     &code_challenge_method=S256
   ```

2. The customer's Global/SharePoint Admin signs into **Microsoft's own hosted page** with their
   M365 credentials — this app never sees or handles that password. `prompt=admin_consent` makes
   Microsoft show its native admin-consent screen here if the app doesn't yet have consent for its
   configured permissions in that tenant; we don't build our own consent UI.
3. Microsoft redirects the popup to `redirect_uri` with `code` and `state` (or `error` if the admin
   declined). The backend validates `state`'s signature/expiry, retrieves the matching PKCE
   verifier, and exchanges `code` via MSAL Node's `ConfidentialClientApplication.acquireTokenByCode`.
4. With the resulting **delegated** access token, the backend calls `GET /organization` (tenant ID
   + verified domain) and `GET /me` (the connecting admin's UPN) purely to identify who connected
   and which tenant — see §5 for why this delegated token isn't used for anything beyond that.
5. The backend upserts a `tenants` row (by Azure AD tenant ID) and a `connections` row for that
   `(tenant, cloud_type)` pair, enqueues the enumeration job, and responds to the popup with a small
   HTML page that `postMessage`s the result to `window.opener` and closes itself — the main window
   listens for that message and refreshes the Manage Clouds list. If the admin declined consent
   (`error` param present) or the code exchange fails, the same postMessage-and-close pattern
   reports failure instead, with no partial `connections` row left in an ambiguous state.

Document the tile-driven flow (not this URL) in customer-facing onboarding — the admin never sees
a raw URL, they see a familiar Microsoft login popup.

## 5. Token acquisition & storage

Two different token types are in play, used for two different purposes — don't conflate them:

1. **Delegated token from the §4a connect flow.** `connections.encrypted_refresh_token` stores the
   *admin's own* refresh token (envelope-encrypted — see `services/tokenEncryption.ts`, KMS-backed
   in production), obtained once at connect time. It's used **only** to re-confirm identity
   (`/me`, `/organization`) if needed later, e.g. to detect the admin's account was disabled — it
   is **not** used to enumerate other users' data. Delegated permissions are scoped to what the
   signed-in admin can personally see, which application permissions deliberately bypass.
2. **Application (client-credentials) token for actual data operations.** All enumeration and
   (later) cleanup calls use the **OAuth 2.0 client-credentials grant** — MSAL Node's
   `ConfidentialClientApplication.acquireTokenByClientCredential`, scoped to the tenant via
   `authority: https://login.microsoftonline.com/{tenant-id}`. This produces a short-lived access
   token with no refresh token (client-credentials tokens are just re-requested from the client
   secret/cert each time) — cached in memory/Redis (~60–90 min TTL), never persisted to Postgres.
   This is what `graph/client.ts`'s `graphClientForTenant()` and the cloud-sync worker both use.

**Why keep the delegated refresh token at all, then?** Admin-consent (§4a) is tenant-wide once
granted — see `docs/graph-api-limitations.md` — so nothing *requires* storing it for OneDrive vs.
SharePoint vs. Teams separately. We keep it per-connection anyway because it's the only signal we
have that a *specific* admin completed *that* connect action, which matters for the audit trail
(`connection_events`) and for showing "connected by" on each Manage Clouds row.

## 6. Connection offboarding / disconnect

Revoking a specific `appRoleAssignedTo` grant via
`DELETE https://graph.microsoft.com/v1.0/servicePrincipals/{sp-id}/appRoleAssignedTo/{assignment-id}`
is **not** something this app can do on the customer's behalf — only the customer's own Global
Admin can revoke consent, from their own Enterprise Applications page. `DELETE
/api/clouds/:id` (`routes/cloudConnections.ts`) therefore does only what's actually ours to do:

1. Cancel any in-flight `sync_jobs` for that connection (`cancel_requested_at`).
2. Discard (null out) our locally stored `encrypted_refresh_token` so we can no longer use it —
   this is a **local** revoke, not a tenant-side one.
3. Mark the `connections` row `disconnected` with `disconnected_at` (soft delete —
   `connection_users`/`sync_jobs` history stays, per the audit-trail requirement).
4. Log a `connection_events` row (`event = 'disconnected'`) with the acting operator.

It does **not** and cannot revoke the underlying Azure AD consent grant — the connection can be
technically "reconnected" by the same tenant admin at any time regardless of our local disconnect,
since the consent still exists on Microsoft's side. Show the customer admin a direct deep link to
their tenant's `Enterprise Applications → CloudFuze M365 Data Cleanup Utility → Permissions` page
with instructions to revoke there if they want to fully remove access.

## 7. Environment separation

- Separate app registrations for dev / staging / prod, each with its own client ID and secret —
  never point a staging deployment at the prod app registration's credential.
- Restrict prod app registration owners in Azure AD to the small group who can rotate credentials
  and review consent grants.
