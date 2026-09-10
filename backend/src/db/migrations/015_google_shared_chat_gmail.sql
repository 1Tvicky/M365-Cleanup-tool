-- Google Shared Drives, Google Chat, and Gmail — the remaining three Google Workspace workloads
-- alongside My Drive (migration 014). See docs/google-workspace-integration.md.

ALTER TABLE connections DROP CONSTRAINT connections_cloud_type_check;
ALTER TABLE connections ADD CONSTRAINT connections_cloud_type_check
  CHECK (cloud_type IN ('onedrive', 'sharepoint', 'teams', 'outlook', 'google_my_drive', 'shared_drive', 'google_chat', 'gmail'));

ALTER TABLE cleanup_operation_items DROP CONSTRAINT cleanup_operation_items_resource_type_check;
ALTER TABLE cleanup_operation_items ADD CONSTRAINT cleanup_operation_items_resource_type_check
  CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'outlook_mailbox',
    'outlook_calendar', 'outlook_contacts', 'channel', 'chat', 'google_my_drive_account',
    'shared_drive', 'google_chat_space', 'gmail_mailbox'));

-- Shared Drives: reused connection_users as-is, same pattern already documented for
-- sharepoint/google_my_drive — graph_user_id = the Shared Drive's id, upn = its display name (a
-- Shared Drive has no email/URL the way a SharePoint site has a webUrl, so display_name doubles
-- for both name and secondary identifier here). No DDL change beyond the CHECK widening above.
--
-- Gmail: also reuses connection_users unchanged — graph_user_id = Directory API user id, upn =
-- primary email, same shape as google_my_drive/outlook_mailbox rows.
--
-- Google Chat is different: a Space isn't a per-user resource the way a Drive/mailbox/site is, so
-- it gets its own dedicated table rather than being force-fit into connection_users — same
-- reasoning that already gives Teams channels their own cleaning_channels table instead of reusing
-- connection_users for them.
CREATE TABLE connection_google_spaces (
  id             UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  connection_id  UUID NOT NULL REFERENCES connections(id) ON DELETE CASCADE,
  space_id       TEXT NOT NULL,
  display_name   TEXT,
  member_count   INTEGER NOT NULL DEFAULT 0,
  message_count  INTEGER NOT NULL DEFAULT 0,
  sync_status    TEXT NOT NULL DEFAULT 'pending' CHECK (sync_status IN ('pending', 'synced', 'failed')),
  last_synced_at TIMESTAMPTZ,
  error_message  TEXT,
  UNIQUE (connection_id, space_id)
);

CREATE INDEX connection_google_spaces_connection_idx ON connection_google_spaces (connection_id);
