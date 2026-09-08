-- Changes Sync from tenant-wide to resource-level. Mirrors cleanup_operation_items'
-- relationship to cleanup_operations (005_cleanup_execution.sql) — one row per resource selected
-- for a given sync_jobs run, rather than adding a job-scoping column to connection_users (which is
-- a single mutable "latest known state" row per resource, read unfiltered by GET /manage's
-- COUNT(*) ... GROUP BY connection_id and by routes/cleaning.ts's discovery tables — allowing
-- multiple historical rows there per (connection_id, graph_user_id) would silently break both).
--
-- An empty sync_job_resources set for a given sync_jobs row means "no selection was made" —
-- jobs/cloudSyncWorker.ts treats that as today's full-tenant enumeration, unchanged. This is what
-- keeps routes/cleaning.ts's "Sync Now" and the initial connect flow (both still insert directly
-- into sync_jobs with no resource selection) working with zero changes to those call sites, and
-- keeps pre-existing sync_jobs history displaying via the same aggregate-only path as before.
CREATE TABLE sync_job_resources (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sync_job_id       UUID NOT NULL REFERENCES sync_jobs(id) ON DELETE CASCADE,
  -- Denormalized from sync_jobs.connection_id — lets the resource-aware concurrency-lock query
  -- (routes/cloudConnections.ts) filter directly on sync_job_resources without a join back to
  -- sync_jobs for every candidate row.
  connection_id     UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  graph_resource_id TEXT NOT NULL,
  display_name      TEXT NOT NULL,
  -- UPN (onedrive/outlook/teams) or webUrl (sharepoint) — snapshotted alongside display_name at
  -- selection time so jobs/cloudSyncWorker.ts can construct the same BasicUser/SiteSummary/
  -- TeamSummary shape its sync* functions already expect directly from this table, with zero Graph
  -- calls needed just to know *which* resources to sync — the worker only calls Graph to actually
  -- sync each one. This is what keeps a resource-scoped run from ever re-enumerating the tenant.
  secondary         TEXT,
  status            TEXT NOT NULL DEFAULT 'pending'
                       CHECK (status IN ('pending', 'processing', 'completed', 'failed', 'cancelled')),
  error_message     TEXT,
  started_at        TIMESTAMPTZ,
  completed_at      TIMESTAMPTZ,
  created_at        TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at        TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (sync_job_id, graph_resource_id)
);

CREATE INDEX sync_job_resources_job_idx ON sync_job_resources (sync_job_id, status);
CREATE INDEX sync_job_resources_connection_resource_idx ON sync_job_resources (connection_id, graph_resource_id);

-- Teams sync moves from "one connection_users row per user (item_count = teams that user joined)"
-- to "one row per actual Team (item_count = channel count)" — the only way Teams can have a
-- selectable-resource shape consistent with OneDrive(users)/SharePoint(sites)/Outlook(mailboxes).
-- Existing user-keyed rows for cloud_type='teams' would otherwise linger under a now-relabeled
-- "Teams" column until the next sync happens to overwrite them by coincidence of matching ids
-- (which it never will, since team ids and user ids don't collide) — delete them now so the next
-- sync populates clean, team-keyed rows instead of leaving stale user data mislabeled.
DELETE FROM connection_users WHERE connection_id IN (SELECT id FROM connections WHERE cloud_type = 'teams');
