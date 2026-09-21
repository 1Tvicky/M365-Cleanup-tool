# Data Dump API

Deliverable — Data Dump module. Generates real Microsoft 365 objects (OneDrive/SharePoint files and
folders, Teams teams/channels, Outlook mail/calendar/contacts) for demos, migration testing,
regression testing, and performance testing. The inverse business function of Cleaning (which
deletes real data) — see `routes/dataDump.ts`'s own docstring for why the two are kept in entirely
separate tables, routes, and workers rather than sharing business logic.

## Mount point and auth

Mounted at `/api/data-dump`, its own namespace (same precedent as `/api/clouds` and `/api/cleaning`).
Reuses this app's existing session auth and `tenant_roles` RBAC unchanged: `viewer` can preview/read,
`cleanup_admin` is required to create, pause, resume, or cancel an operation — the same split
Cleanup's own execute/cancel actions already use. No new roles, no new auth mechanism.

## Endpoints

| Method & path | Purpose |
|---|---|
| `GET /tenants` | Tenant + workload/connectionId picker for the Data Dump UI |
| `GET /tenants/:tenantId/users` | Tenant-wide real-member user listing (powers the Teams member picker) |
| `GET /profile-defaults?profile=&workloads=` | Default config for a profile, scoped to selected workloads |
| `POST /preview` | Pure computation of requested-object counts/estimated size — never creates anything |
| `POST /` | Creates an operation and queues generation; returns `{ operationId, status: "queued" }` |
| `GET /history` | Paginated/searchable operation list, for Data Dump Reports |
| `GET /summary` | Aggregate counts for the Reports/Dashboard summary cards |
| `GET /:operationId` | One operation's current status, counts, `m365TenantId`, and per-workload task detail (including `subcounts`) |
| `GET /:operationId/resources` | Structural containers + a page of batch summaries — "View Generated Resources" |
| `POST /:operationId/pause` / `/resume` / `/cancel` | Operation control |
| `GET /:operationId/report` | CSV download, `data-dump-report-<id>.csv`, with an explicit `Report Type: Data Dump` line and the M365 Tenant ID |

`GET /:operationId`, `GET /history`, and the CSV report all include the tenant's Azure AD directory
(tenant) ID (`tenants.m365_tenant_id`) — added specifically so an operator can correlate a run
against Entra ID / Graph audit logs when troubleshooting a failed or partially-failed operation,
without having to separately go look up which tenant a given operation ran against.

Resource *discovery* (which users/sites/teams exist to select from) deliberately reuses Add Clouds'
existing `GET /api/clouds/:connectionId/available-resources` endpoint unchanged — this is shared,
read-only infrastructure, not something Data Dump needed its own version of.

## Explicit resource selection

Each workload's config accepts an explicit selection alongside its count-based quick-start fields:

- `onedrive.selectedUserIds` / `outlook.selectedUserIds` — exact users to generate into.
- `sharepoint.selectedSiteIds` — existing sites; `sharepoint.newSite` — create one new site first.
- `teams.selectedTeamIds` — existing teams; `teams.newTeam` — create one new team first;
  `teams.channels` — explicit per-channel `{name, memberUserIds, messages, replies}`.

When a selection is present, the worker resolves it via direct Graph lookups
(`getUserById`/`getSiteById`/`getTeamById`) rather than re-listing and slicing the tenant. Without
any selection, each workload falls back to its original count-based quick-start behavior
(`userCount`, `teamCount`, etc.) so profile-driven generation with no manual picking still works.

## Folder/file structure

OneDrive (`onedrive.*`) and SharePoint (`sharepoint.*`) both generate the same simple, uniform
folder shape, controlled by three numbers per workload:

- `rootFolders` / `foldersPerLibrary` — total root folders created directly under the user's
  OneDrive (or under each document library).
- `subFoldersPerFolder` — how many children every folder gets, uniformly, at every level below the
  root (e.g. 20 root folders + 3 here means every one of the 20 gets exactly 3 children).
- `maxFolderDepth` — how many levels deep the tree goes; `1` means root folders only, no
  sub-folders at all regardless of `subFoldersPerFolder`.

`filesPerFolder` is generated in **every** folder in the tree — root, sub-folder, and nested folder
alike, not leaf folders only. See `services/dataDump/businessContent.ts`'s `buildFolderTree` /
`countFolderNodes` (the same tree-count math backs `POST /preview`'s estimate, so the number shown
there always matches what a run will actually attempt) and each runner's `processFolderNode`, which
walks the whole tree with one recursive function — every node is created, filled with files, then
recursed into identically, no root/leaf distinction.

Every folder gets a real, professional, **guaranteed-unique** name across the whole tree (a
seeded-random pick from a large business-word pool, numbered only once the pool is exhausted) —
never `Folder1`/`Folder2` unless `namingStyle: "synthetic"` is explicitly selected.

## Content generation

Generated files are complete, realistic business documents, not placeholder text — each file's name
and its content always come from the same generator, so a file never opens to content that
contradicts its own name (`services/dataDump/businessContent.ts` + `fileContent.ts`):

- **docx** — a Project Charter (with a real Name/Role/Responsibility stakeholder table), a Business
  Requirements Document, or a Standard Operating Procedure, picked per file.
- **xlsx** — a quarterly Budget Summary or Headcount Plan workbook with live `SUM` formulas, styled
  header/total rows, and a notes section.
- **pdf** — a Master Services Agreement or Mutual NDA: numbered clauses, named parties with real
  addresses, and a two-party signature block, word-wrapped across as many pages as needed.
- **pptx** — a title slide plus one slide per topic, each with real bullet content.
- **txt** — a structured status memo (From/To/Subject header, discussion points, action items with
  owners and due dates).
- **csv** — a tabular data export.

An operator can restrict generation to specific extensions via each workload's `fileTypeDistribution`
(exposed in the UI as a "File Types" checkbox picker) instead of the profile's default mixed
distribution. `namingStyle: "synthetic"` bypasses all of the above in favor of flat, obviously-fake
`File1.docx`-style names and minimal content, for callers that specifically want non-business-looking
test data.

## Scalability and resumability

- BullMQ payload is `{ operationId }` only (spec: never put bulk config/ids in the queue) — the
  worker reloads everything from Postgres.
- Leaf objects (files, mail items) are tracked in **batches** (`data_dump_batches`, one row per up to
  `DATA_DUMP_BATCH_SIZE` objects — see `services/dataDump/batching.ts`), never one row per object.
  Structural containers (folders, teams, channels, libraries, sites) get their own row in
  `data_dump_containers` — bounded count even at huge scale.
- Resumability is at container + batch granularity: a restart never recreates an already-recorded
  container (`services/dataDump/containers.ts`'s get-or-create), and resumes file/mail generation
  from `nextBatchIndexFor` — at most one partial batch is redone, never the whole workload.
- Failures are isolated at every level of the tree, not just per-user/per-site: one user's missing
  OneDrive (confirmed live) does not abort the other users in the same run, and — one level deeper —
  one folder's Graph failure deep in a large tree (confirmed live: a 630-folder/1,890-file run where
  an early failure used to bulk-fail everything after it) now aborts only that folder's own subtree,
  never its siblings or any other root. See each runner's outer per-user/per-site catch AND
  `processFolderNode`'s own per-node catch.
- Whichever share of a failed node's fair-share folders/files never got a chance to run is credited
  as failed explicitly (mirroring the same tree-count math `POST /preview` uses), so
  `Requested = Created + Failed + Skipped` always reconciles exactly on a finished operation — never
  a permanent, unexplained "Remaining" gap.
- Every resource kind (users, folders, files, sites, libraries, teams, channels, channel members,
  messages, replies, emails, calendar events, contacts) tracks its own `requested`/`created`/
  `failed`/`skipped` in `data_dump_workload_tasks.subcounts`, reconciling independently.

## Known limitations — Microsoft Graph platform, not app-imposed

1. **New SharePoint site creation requires the beta API + `Sites.Create.All`.** Implemented for
   real (`graph/dataDump/sharePointCreation.ts`'s `createSharePointSite`), but this app's Azure AD
   registration does not request `Sites.Create.All` by default — see
   `docs/azure-ad-app-registration.md` §6a. Without it, this one feature 403s with a clear error;
   every other Data Dump capability, including generating into *existing* SharePoint sites, is
   unaffected.
2. **Teams channel messages/replies are never posted.** `ChannelMessage.Send` is delegated-only;
   application-permission channel-message posting is restricted by Microsoft to a "migration mode"
   team-import state this app does not implement. Configured message/reply counts are shown in
   Preview/Review for planning purposes and accounted as "skipped," never faked as created.
3. **Historical Outlook mail timestamps are not settable.** Custom `receivedDateTime`/`sentDateTime`
   on a created message is only available on the unsupported Graph beta endpoint. Generated mail is
   timestamped at creation time. Calendar events do NOT have this limitation — a genuinely historical
   `start`/`end` on an event is a normal, v1.0-supported, client-set property.

## Safety

- Generated Outlook mail is created directly inside a mailbox folder
  (`POST /users/{id}/mailFolders/{folder}/messages`) — never via `/sendMail` or any endpoint capable
  of dispatching real delivery. No code path in `graph/dataDump/outlookCreation.ts` can send mail.
- Every "member"/"owner"/"attendee" is an existing tenant user, resolved by id — Data Dump never
  creates an external or guest identity.
- `services/dataDump/targetUsers.ts`'s `filterRealMemberUsers` excludes guest (`#EXT#`) and
  soft-deleted/orphaned directory objects from the profile-driven quick-start user pool — confirmed
  live against a real tenant to matter (the first several users returned by `listAllUsers` were guest/
  deleted placeholders with no provisionable OneDrive).
