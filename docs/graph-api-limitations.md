# Graph API Limitations — Add Clouds / Manage Clouds

Deliverable 5 — real constraints of Microsoft Graph that shape what the Cleanup module can and
can't promise. Read this before scoping Cleanup; nothing here is worked around silently anywhere
in this connection layer's code.

## Chat / personal message deletion is not achievable with application permissions

`Chat.Read.All` (application) lets this app **read** the metadata and content of 1:1 and group
chats tenant-wide. There is no corresponding `Chat.ReadWrite.All`-for-delete capability that lets
an application-permission caller delete **another user's** personal chat messages. Graph's
message-delete endpoint (`DELETE /chats/{chat-id}/messages/{message-id}`) only works in a
**delegated** context where the caller *is* the message's author — an app-only token can't
impersonate that. Concretely:

- We can enumerate chats and (roughly) size them for the Manage Clouds progress/storage figures.
- We cannot bulk-delete chat history as part of Cleanup, no matter what permissions are granted.
- The only real path to purging chat content at scale is Microsoft's **eDiscovery / compliance
  content search + purge** workflow, which is a fundamentally different (and much heavier) API
  surface than anything in this connection layer, and out of scope here.

**Action for the Cleanup module**: don't build a "delete chats" feature. If the product needs
this, it needs a compliance-API-based design from scratch, scoped separately.

## SharePoint site discovery via Search API can lag

Per your direction, site enumeration uses `GET /sites?search=` rather than a full site-collection
crawl (see `docs/cloud-connections-api.md`). Trade-off to keep in mind:

- Search-index-backed, so a site created moments ago may not appear in results yet. In practice
  this is a short window (typically minutes), not a correctness bug in our code — a resync will
  pick it up.
- Highly restrictive site permissions or unusual naming can occasionally cause a site to not
  surface in search results even after indexing. There's no cheap way to distinguish "doesn't
  exist" from "search missed it" without falling back to a full crawl.
- A tenant that needs guaranteed-complete enumeration (compliance-driven cleanup, for example)
  should not rely on this connection layer's default sync — that's a "full crawl" mode we
  explicitly chose not to build in this pass.

## Graph throttling (429/503) is real and per-tenant

Every tenant has its own Graph API rate budget. A large tenant's enumeration job can and will hit
`429 Too Many Requests` mid-run under normal operation — this isn't an error condition, it's
expected steady-state behavior at scale. The enumeration worker (`jobs/cloudSyncWorker.ts`) honors
`Retry-After` and backs off exponentially (reusing `services/rateLimiter.ts`), but this means large
tenants will have **slower** enumeration, not failed enumeration. Don't interpret a long-running
sync job as stuck without checking `sync_jobs.status` first.

## Application-permission admin consent is tenant-wide, not per-workload

Azure AD admin consent is granted for the **app registration's entire configured permission set**
in one action — there is no way to grant "just OneDrive permissions" separately from "just Teams
permissions" at the Azure AD level. This connection layer still tracks one `connections` row per
`cloud_type` (matching the three separate "Add Cloud" tiles), but once an admin consents once for
any tile, the underlying application-permission grant technically covers all three. See the "why"
comment in `routes/cloudConnections.ts` — this is intentional, not a bug, and it's why connecting
OneDrive today and SharePoint next week doesn't require a second admin-consent prompt (though the
UI still shows it as a separate connect action, since the *enumeration and data* per cloud type are
tracked independently).

## We cannot force-revoke a tenant's consent on disconnect

There is no Graph API call this app can make to unilaterally revoke the admin consent grant a
customer tenant gave us — that's the customer's own tenant configuration, not ours to touch.
"Disconnect" in this app:

1. Discards our locally stored (encrypted) refresh token so we can no longer use it.
2. Cancels any in-flight enumeration job.
3. Marks the connection `disconnected` (soft delete, audit trail preserved).

It does **not** and cannot revoke the underlying Azure AD consent grant. See
`docs/azure-ad-app-registration.md` §6 and the code comment in the `DELETE /api/clouds/:id`
handler — the customer's own Global Admin has to do that from their Enterprise Applications page
if they want to fully remove access.

## SharePoint has no "users" — `connection_users` rows mean something different per cloud type

The `connection_users` table (shared schema across all three cloud types, per the original spec)
is genuinely user-shaped for OneDrive and Teams, but SharePoint doesn't enumerate *users* — it
enumerates *sites*. For `cloud_type = 'sharepoint'`, each `connection_users` row represents one
**site**, not one person: `graph_user_id` holds the site ID, `upn` holds the site's `webUrl`, and
`display_name` holds the site's display name. This is a deliberate schema reuse (documented inline
in `graph/cloudEnumeration.ts`), not an accident — flagging it here too so it isn't mistaken for a
data-quality bug later.
