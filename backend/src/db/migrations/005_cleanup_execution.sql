-- Cleanup/deletion execution phase. A cleanup operation can span up to 3 connections of one
-- tenant (OneDrive + SharePoint + Teams), matching the existing Review-your-selection screen,
-- so this is keyed by tenant_id, not connection_id (see cleanup_operation_items for the per-
-- connection breakdown). See docs/ (cleanup-execution plan) for the full design rationale.

CREATE TABLE cleanup_operations (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id             UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  requested_by          UUID NOT NULL REFERENCES operators(id),
  status                TEXT NOT NULL DEFAULT 'queued'
                          CHECK (status IN ('queued', 'running', 'completed', 'completed_with_errors', 'failed', 'cancelled')),
  total_items           INTEGER NOT NULL DEFAULT 0,
  processed_items       INTEGER NOT NULL DEFAULT 0,
  successful_items      INTEGER NOT NULL DEFAULT 0,
  failed_items          INTEGER NOT NULL DEFAULT 0,
  -- Includes 'unsupported' items (e.g. Teams channels/chats — see cleanupDeletion.ts) in the aggregate.
  skipped_items         INTEGER NOT NULL DEFAULT 0,
  retry_of_operation_id UUID REFERENCES cleanup_operations(id),
  started_at            TIMESTAMPTZ,
  completed_at          TIMESTAMPTZ,
  cancel_requested_at   TIMESTAMPTZ,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  error_message         TEXT
);

CREATE INDEX cleanup_operations_tenant_idx ON cleanup_operations (tenant_id, created_at DESC);

CREATE TABLE cleanup_operation_items (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  cleanup_operation_id  UUID NOT NULL REFERENCES cleanup_operations(id) ON DELETE CASCADE,
  -- Every other FK to connections() in this repo is ON DELETE CASCADE; deliberately RESTRICT here —
  -- a destructive-deletion audit trail must outlive the connection it was performed through.
  connection_id         UUID NOT NULL REFERENCES connections(id) ON DELETE RESTRICT,
  resource_type         TEXT NOT NULL CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'channel', 'chat')),
  -- connection_users.id / cleaning_channels.id / cleaning_chats.id at manifest time — for
  -- display/back-reference only, never re-joined against those tables at execution time.
  resource_id           UUID NOT NULL,
  -- Snapshotted at manifest time so later discovery changes never affect what this record shows.
  display_name          TEXT NOT NULL,
  -- Snapshotted Graph-facing id(s) needed to execute the delete:
  -- {"userId":...} | {"siteId":...} | {"teamId":...,"channelId":...} | {"chatId":...}
  graph_ref             JSONB NOT NULL,
  status                TEXT NOT NULL DEFAULT 'pending'
                          CHECK (status IN ('pending', 'processing', 'completed', 'failed', 'skipped', 'unsupported')),
  attempts              INTEGER NOT NULL DEFAULT 0,
  started_at            TIMESTAMPTZ,
  completed_at          TIMESTAMPTZ,
  error_code            TEXT,
  error_message         TEXT,
  last_response         JSONB,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX cleanup_operation_items_operation_idx ON cleanup_operation_items (cleanup_operation_id, status);
CREATE INDEX cleanup_operation_items_connection_idx ON cleanup_operation_items (connection_id);

ALTER TABLE connection_events DROP CONSTRAINT connection_events_event_check;
ALTER TABLE connection_events ADD CONSTRAINT connection_events_event_check CHECK (event IN (
  'token_exchange', 'connected', 'connect_failed', 'job_started', 'job_finished',
  'reauth_required', 'resync_requested', 'disconnected',
  'cleaning_scan_started', 'cleaning_scan_finished',
  'cleanup_requested', 'cleanup_started', 'cleanup_completed', 'cleanup_completed_with_errors',
  'cleanup_cancelled', 'cleanup_failed'
));
