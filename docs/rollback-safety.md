# Rollback & Safety Doc

Deliverable 4 of 4 — what's recoverable vs. not, and what the tool does to compensate.

## Why this matters

Graph API deletes used by this tool are **application-permission, programmatic deletes** — not a
user dragging a file to the Recycle Bin in the SharePoint UI. Some land in a recoverable state by
default; others are effectively permanent within normal operational timeframes. This doc is meant
to be read by the admin before they type `DELETE`, and linked from the confirmation modal.

## Recoverability by workload

| Workload | What we delete | Native recovery window | Recoverable how | Notes |
|---|---|---|---|---|
| OneDrive files | `DELETE /drives/{id}/items/{id}` | ~93 days (OneDrive/SharePoint Recycle Bin, first + second stage) | Site collection admin restores from Recycle Bin | Recycle bin has a storage quota; a very large batch delete can push older recycled items into permanent purge sooner than 93 days |
| SharePoint document libraries / site content | `DELETE` on drive items, or full library removal | ~93 days for item-level deletes; **library/list deletion itself may not be recoverable via UI** past the site's recycle bin retention | Site collection admin restores from site Recycle Bin (Site Settings → Recycle Bin) | Deleting an entire document library (not just its files) is a heavier operation — verify the tenant's SharePoint recycle bin retention hasn't been shortened by tenant policy before relying on this |
| Teams channels (`Channel.Delete.All`) | Standard channels: soft-deleted, restorable for **21 days** via Teams admin center or Graph `POST /teams/{id}/channels/{id}/restore` | 21 days | Teams/Global admin restores via admin center or Graph restore endpoint | **Private channels cannot be restored once deleted** — Graph offers no restore endpoint for private channels. Shared channels: treat as unrestorable until Microsoft documents otherwise |
| M365 Groups (when a Team's group is removed) | `DELETE /groups/{id}` | **30 days** (Azure AD soft-delete) | Restore via `POST /directory/deletedItems/{id}/restore` within the 30-day window; after that, permanent | Restoring the group does not automatically restore an already-deleted Team's channels/tabs/files if those were separately purged |
| Teams chat / DM messages | **Not deleted in v1** (report-only, per your Phase 1 scope decision) | N/A | N/A | See [azure-ad-app-registration.md](azure-ad-app-registration.md) — application permissions can't bulk-delete chat content the way channel deletes work; deferred to v2 |

**Bottom line ranked by risk:** private Teams channels and anything past its recycle-bin/soft-delete
window are **permanent**. Everything else has a real but time-boxed recovery window that depends on
someone noticing the mistake and acting inside it — which is why this tool does not treat "there's a
recycle bin" as a substitute for its own export step.

## What this tool does to reduce risk (independent of native recovery windows)

1. **Dry-run by default, always.** `POST /cleanup/preview` never calls a delete endpoint — see
   [api-spec.md](api-spec.md#preview-dry-run--no-destructive-calls). An admin can preview a scope
   as many times as needed with zero risk.
2. **Mandatory pre-delete export, no skip option.** Per your decision, every confirmed job first
   generates a manifest CSV (item-level: path, ID, size, owner) and, unless the admin explicitly
   chose metadata-only, a zip of file contents to blob storage — *before* any delete call is queued.
   This is CloudFuze's own recovery path, independent of Microsoft's recycle-bin/soft-delete
   windows, and it's the one thing that survives a purged recycle bin or an expired 21/30/93-day
   window.
3. **Typed confirmation.** The literal string `DELETE` must be typed, matched server-side
   (`api-spec.md` → `CONFIRMATION_MISMATCH`), not just a checkbox — deliberately adds friction
   proportional to the irreversibility above.
4. **Per-item audit log, always persisted.** Every attempted delete (success, failure, or skip) is
   written to the audit table before the job reports complete, exportable as CSV. If something was
   deleted that shouldn't have been, the audit log plus the manifest is what tells you exactly what,
   when, by whom, and — for anything still inside its recovery window — what to go restore and
   where.
5. **RBAC gate on execute, not just UI hiding.** `viewer` role is enforced server-side on every
   mutating route, not just hidden in the frontend, so a viewer can't discover an execute endpoint
   and use it directly.
6. **Best-effort cancel, not undo.** `POST /jobs/:jobId/cancel` stops new batches from starting; it
   does not and cannot reverse items already deleted in-flight. The progress UI is intentionally
   granular (X of Y) so an admin who spots a mistake mid-run can act before the whole scope
   processes.

## What this tool deliberately does *not* claim

- It does not promise deletes are reversible through the tool itself — recovery for anything within
  its native window happens in the M365 admin center / Graph restore endpoints, by a tenant admin,
  not by clicking something in this app. Phase 1 does not build an in-app "undo" flow.
- It does not extend or override a customer tenant's own retention/recycle-bin policy — if a
  tenant's SharePoint recycle bin retention has been shortened by their own admin, this tool has no
  visibility into that until it's reflected in Graph responses, and the preview step should surface
  a warning when retention metadata suggests a shorter-than-default window.
- Blob-storage exports are themselves data at rest containing customer content — they need their
  own retention/deletion policy so CloudFuze doesn't end up as an indefinite second copy of data the
  customer asked to have removed. (Recommend: export retention = 90 days, configurable per
  engagement, documented separately in the data-handling policy — out of scope for this doc.)
