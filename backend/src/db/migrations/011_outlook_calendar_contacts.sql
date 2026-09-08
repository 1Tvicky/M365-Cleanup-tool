-- Adds Calendar and Contacts as cleanable resources under the existing Outlook connection
-- (cloud_type stays 'outlook' — these are resource kinds within it, not new cloud types).
--
-- Deliberately two new tables, not a resource_type column added to connection_users: mirrors how
-- Teams' channels/chats already get their own dedicated tables (cleaning_channels/cleaning_chats)
-- rather than overloading a shared table with a discriminator. Keeps the already-working
-- OneDrive/SharePoint/Teams/Outlook-Mail code paths (which all assume exactly one
-- (storage_used_bytes, item_count) pair per connection_users row) completely untouched.
--
-- Both are mailbox-level *summary* rows — one row per user, item_count is an aggregate (total
-- calendar events / total contacts across all of that user's calendars/contact folders) — the same
-- granularity connection_users already uses for Mail. There is no per-event or per-contact row
-- anywhere in this schema; cleanup at execution time still walks Graph directly per selected user
-- (see graph/cleanupDeletion.ts), this table only drives the discovery/selection UI list.

CREATE TABLE connection_outlook_calendars (
  id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id      UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  graph_user_id      TEXT NOT NULL,
  upn                TEXT NOT NULL,
  display_name       TEXT,
  storage_used_bytes BIGINT NOT NULL DEFAULT 0, -- not meaningful, same as Mail — kept for CleaningResourceRow shape parity
  item_count         INTEGER NOT NULL DEFAULT 0, -- total calendar EVENT count across all the user's calendars
  sync_status        TEXT NOT NULL DEFAULT 'pending' CHECK (sync_status IN ('pending', 'synced', 'failed')),
  last_synced_at     TIMESTAMPTZ,
  error_message      TEXT,
  UNIQUE (connection_id, graph_user_id)
);

CREATE INDEX connection_outlook_calendars_connection_idx ON connection_outlook_calendars (connection_id);

CREATE TABLE connection_outlook_contacts (
  id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id      UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  graph_user_id      TEXT NOT NULL,
  upn                TEXT NOT NULL,
  display_name       TEXT,
  storage_used_bytes BIGINT NOT NULL DEFAULT 0,
  item_count         INTEGER NOT NULL DEFAULT 0, -- total contact count across all the user's contact folders
  sync_status        TEXT NOT NULL DEFAULT 'pending' CHECK (sync_status IN ('pending', 'synced', 'failed')),
  last_synced_at     TIMESTAMPTZ,
  error_message      TEXT,
  UNIQUE (connection_id, graph_user_id)
);

CREATE INDEX connection_outlook_contacts_connection_idx ON connection_outlook_contacts (connection_id);

ALTER TABLE cleanup_operation_items DROP CONSTRAINT cleanup_operation_items_resource_type_check;
ALTER TABLE cleanup_operation_items ADD CONSTRAINT cleanup_operation_items_resource_type_check
  CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'outlook_mailbox', 'outlook_calendar', 'outlook_contacts', 'channel', 'chat'));
