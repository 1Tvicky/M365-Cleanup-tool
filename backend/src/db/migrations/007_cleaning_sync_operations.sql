-- "Sync Now" for the Cleaning page. This does NOT introduce a new sync engine — it's a thin
-- tenant-level wrapper around the existing sync_jobs (OneDrive/SharePoint, cloudSyncWorker.ts) and
-- cleaning_scans (Teams, cleaningScanWorker.ts) mechanisms, so the Cleaning page can trigger and
-- poll all of a tenant's connections with one id instead of up to three.

CREATE TABLE cleaning_sync_operations (
  id                     UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id              UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  requested_by           UUID NOT NULL REFERENCES operators(id),
  -- Nullable: a tenant may not have all three connections. No status/completed_at column here —
  -- deliberately computed live from these referenced rows' own current state on every read,
  -- rather than a second, driftable copy of it.
  onedrive_sync_job_id   UUID REFERENCES sync_jobs(id),
  sharepoint_sync_job_id UUID REFERENCES sync_jobs(id),
  teams_scan_id          UUID REFERENCES cleaning_scans(id),
  started_at             TIMESTAMPTZ NOT NULL DEFAULT now(),
  created_at             TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX cleaning_sync_operations_tenant_idx ON cleaning_sync_operations (tenant_id, created_at DESC);

-- cleaning_channels/cleaning_chats are owned exclusively by the Cleaning module (unlike
-- connection_users, which Manage Clouds also reads unfiltered) so soft-deleting here is zero-risk:
-- a channel/chat no longer enumerated by Microsoft 365 is marked inactive instead of removed,
-- preserving history, and reactivated (is_active = true) if it reappears on a later sync.
ALTER TABLE cleaning_channels ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT true;
ALTER TABLE cleaning_chats ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT true;

ALTER TABLE connection_events DROP CONSTRAINT connection_events_event_check;
ALTER TABLE connection_events ADD CONSTRAINT connection_events_event_check CHECK (event IN (
  'token_exchange', 'connected', 'connect_failed', 'job_started', 'job_finished',
  'reauth_required', 'resync_requested', 'disconnected',
  'cleaning_scan_started', 'cleaning_scan_finished',
  'cleanup_requested', 'cleanup_started', 'cleanup_completed', 'cleanup_completed_with_errors',
  'cleanup_cancelled', 'cleanup_failed',
  'cleaning_sync_requested'
));
