-- Adds "outlook" as a fourth cloud type, alongside onedrive/sharepoint/teams. Postgres CHECK
-- constraints can't be altered in place, so each is dropped and recreated with the wider list —
-- using Postgres's default <table>_<column>_check naming, since neither constraint was given an
-- explicit name when the table was created (002_cloud_connections.sql, 005_cleanup_execution.sql).

ALTER TABLE connections DROP CONSTRAINT connections_cloud_type_check;
ALTER TABLE connections ADD CONSTRAINT connections_cloud_type_check
  CHECK (cloud_type IN ('onedrive', 'sharepoint', 'teams', 'outlook'));

ALTER TABLE cleanup_operation_items DROP CONSTRAINT cleanup_operation_items_resource_type_check;
ALTER TABLE cleanup_operation_items ADD CONSTRAINT cleanup_operation_items_resource_type_check
  CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'outlook_mailbox', 'channel', 'chat'));
