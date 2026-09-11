-- Adds 'team' as a cleanup_operation_items.resource_type — whole-Team deletion (DELETE
-- /groups/{id}, which removes the M365 Group backing the Team along with every channel under it
-- per Graph/Entra semantics), distinct from the existing 'channel' type (deletes one channel,
-- parent Team untouched). No new table: a Team has no dedicated row anywhere in this schema (see
-- migrations/004_cleaning_module.sql's comment — team_id/team_name are denormalized columns on
-- cleaning_channels, a Team is only implicit as the distinct team_id values across channel rows),
-- so a "team" cleanup_operation_items row's resource_id is the team_id (Graph group id) itself,
-- not a synced row's primary key — the one deliberate exception to this table's usual convention,
-- documented on CleanupManifest in types/cleaning.ts.
ALTER TABLE cleanup_operation_items DROP CONSTRAINT cleanup_operation_items_resource_type_check;
ALTER TABLE cleanup_operation_items ADD CONSTRAINT cleanup_operation_items_resource_type_check
  CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'outlook_mailbox',
    'outlook_calendar', 'outlook_contacts', 'channel', 'chat', 'google_my_drive_account',
    'shared_drive', 'google_chat_space', 'gmail_mailbox', 'team'));
