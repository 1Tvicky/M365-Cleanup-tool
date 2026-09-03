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
- Redirect URI (Web): `https://<app-host>/api/auth/callback` — must be HTTPS, no wildcards
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
| `Files.ReadWrite.All` | Application | Enumerate + delete OneDrive files | Preview, Execute |
| `Sites.ReadWrite.All` | Application | Enumerate + delete SharePoint site content/libraries | Preview, Execute |
| `User.Read.All` | Application | Resolve users for the tenant's user/site picker | Connect, Cleanup |
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

## 4. Admin consent flow (per customer tenant)

1. CloudFuze sends the customer's Global Admin the tenant-specific consent URL:
   `https://login.microsoftonline.com/{customer-tenant-id-or-common}/adminconsent?client_id={app-client-id}&redirect_uri={redirect_uri}`
2. Customer admin reviews the requested application permissions (table above) and clicks **Accept**.
3. Azure AD redirects back to `redirect_uri` with `tenant` and `admin_consent=True` in the query
   string — the backend uses this to record that tenant as consented and to mint the first
   application-token via client-credentials flow (see §5).
4. If the admin declines or closes the window, the backend must treat the tenant as
   **not connected** — do not assume consent from a redirect that lacks `admin_consent=True`.

Document this URL and screenshot the consent screen in the customer-facing "Connect your tenant"
onboarding page — a Global Admin who doesn't recognize the permission list will not click Accept.

## 5. Token acquisition & storage

- Use the **OAuth 2.0 client-credentials grant** (via MSAL Node's `ConfidentialClientApplication`)
  to mint application-permission access tokens per customer tenant — there is no per-tenant
  refresh token in this flow; MSAL re-acquires access tokens from its own token cache using the
  app's client credential.
- What actually gets persisted per tenant is **not a user refresh token** (application permissions
  don't produce one) — it's the tenant ID plus the consent record. Short-lived access tokens are
  cached in memory/Redis (~60–90 min TTL) and re-minted from the client credential, never
  persisted to Postgres.
- If a future feature needs delegated permissions (a specific admin's identity, not app-only),
  store *that* refresh token encrypted at rest (envelope encryption via Key Vault-managed key,
  never plaintext, never logged) and scope it to the minimum delegated permissions needed.

## 6. Tenant offboarding / disconnect

- "Disconnect" in the Manage Clouds screen must call
  `DELETE https://graph.microsoft.com/v1.0/servicePrincipals/{sp-id}/appRoleAssignedTo/{assignment-id}`
  is **not** something this app can do on the customer's behalf (only the customer admin can revoke
  consent from their own Enterprise Applications page) — so "Disconnect" should:
  1. Immediately stop scheduling/running jobs for that tenant and purge cached tokens.
  2. Mark the tenant record `disconnected` in Postgres (audit trail preserved, not deleted).
  3. Show the customer admin a direct deep link to their tenant's
     `Enterprise Applications → CloudFuze M365 Data Cleanup Utility → Permissions` page with
     instructions to revoke there if they want to fully remove access.

## 7. Environment separation

- Separate app registrations for dev / staging / prod, each with its own client ID and secret —
  never point a staging deployment at the prod app registration's credential.
- Restrict prod app registration owners in Azure AD to the small group who can rotate credentials
  and review consent grants.
