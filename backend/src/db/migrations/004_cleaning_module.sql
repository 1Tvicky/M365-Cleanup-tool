-- Cleaning module (discovery phase). Read-only: browse the already-connected OneDrive/SharePoint/
-- Teams data via `connections`, no deletion. OneDrive/SharePoint discovery reads the existing
-- `connection_users` table directly (already populated by cloudSyncWorker.ts) — only Teams needs
-- new tables, since nothing today lists actual teams/channels/chats or message counts.

-- Generalized job-status table for both Teams discovery phases (mirrors sync_jobs's shape).
CREATE TABLE cleaning_scans (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id       UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  scan_type           TEXT NOT NULL CHECK (scan_type IN ('teams_structure', 'message_counts')),
  status              TEXT NOT NULL DEFAULT 'queued'
                         CHECK (status IN ('queued', 'running', 'completed', 'completed_with_errors', 'failed', 'cancelled')),
  total_items         INTEGER NOT NULL DEFAULT 0,
  processed_items     INTEGER NOT NULL DEFAULT 0,
  started_at          TIMESTAMPTZ,
  finished_at         TIMESTAMPTZ,
  error_log           JSONB NOT NULL DEFAULT '[]',
  cancel_requested_at TIMESTAMPTZ,
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX cleaning_scans_connection_idx ON cleaning_scans (connection_id, created_at DESC);

-- One row per team+channel. message_count is NULL until a 'message_counts' scan computes it —
-- Graph has no count endpoint, so this is filled in by paginating messages, never faked.
CREATE TABLE cleaning_channels (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id   UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  team_id         TEXT NOT NULL,
  team_name       TEXT NOT NULL,
  channel_id      TEXT NOT NULL,
  channel_name    TEXT NOT NULL,
  message_count   INTEGER,
  count_status    TEXT NOT NULL DEFAULT 'pending' CHECK (count_status IN ('pending', 'calculating', 'completed', 'failed')),
  error_message   TEXT,
  last_counted_at TIMESTAMPTZ,
  UNIQUE (connection_id, channel_id)
);

CREATE INDEX cleaning_channels_connection_idx ON cleaning_channels (connection_id);

-- One row per 1:1/group chat, deduplicated across participants — Graph has no tenant-wide
-- "list all chats" call under application permissions, so these are discovered by enumerating
-- every user's own chats and merging by chat_id. See docs/graph-api-limitations.md.
CREATE TABLE cleaning_chats (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id   UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  chat_id         TEXT NOT NULL,
  chat_type       TEXT NOT NULL CHECK (chat_type IN ('oneOnOne', 'group', 'meeting', 'unknownFutureValue')),
  participants    JSONB NOT NULL DEFAULT '[]',
  message_count   INTEGER,
  count_status    TEXT NOT NULL DEFAULT 'pending' CHECK (count_status IN ('pending', 'calculating', 'completed', 'failed')),
  error_message   TEXT,
  last_message_at TIMESTAMPTZ,
  last_counted_at TIMESTAMPTZ,
  UNIQUE (connection_id, chat_id)
);

CREATE INDEX cleaning_chats_connection_idx ON cleaning_chats (connection_id);

-- Extend the existing connection_events audit trail rather than adding a parallel table.
ALTER TABLE connection_events DROP CONSTRAINT connection_events_event_check;
ALTER TABLE connection_events ADD CONSTRAINT connection_events_event_check CHECK (event IN (
  'token_exchange', 'connected', 'connect_failed', 'job_started', 'job_finished',
  'reauth_required', 'resync_requested', 'disconnected',
  'cleaning_scan_started', 'cleaning_scan_finished'
));
