-- Google Workspace "My Drive" — first Google workload, added additively alongside the existing
-- M365 cloud types. See docs/google-workspace-integration.md for the full architecture writeup.
--
-- tenants.m365_tenant_id becomes nullable because a Google Workspace customer has no Azure AD
-- directory id — the new CHECK enforces that every tenant still has exactly one kind of external
-- directory identity, never neither. This widens the column's TypeScript type across every file
-- that reads a `tenants` row, but only Google-cloud-type connections will ever actually see a NULL
-- here — every M365 code path is reached exclusively through an onedrive/sharepoint/teams/outlook
-- connection, which always has a non-null m365_tenant_id, so this is a type-widening, not a
-- behavior change, for the existing M365 pipeline.
ALTER TABLE tenants ALTER COLUMN m365_tenant_id DROP NOT NULL;
ALTER TABLE tenants ADD COLUMN google_customer_id TEXT UNIQUE;
ALTER TABLE tenants ADD CONSTRAINT tenants_directory_id_check
  CHECK (m365_tenant_id IS NOT NULL OR google_customer_id IS NOT NULL);

-- Postgres CHECK constraints can't be altered in place — dropped and recreated with the wider list,
-- same pattern as 009_outlook_cloud_type.sql.
ALTER TABLE connections DROP CONSTRAINT connections_cloud_type_check;
ALTER TABLE connections ADD CONSTRAINT connections_cloud_type_check
  CHECK (cloud_type IN ('onedrive', 'sharepoint', 'teams', 'outlook', 'google_my_drive'));

ALTER TABLE cleanup_operation_items DROP CONSTRAINT cleanup_operation_items_resource_type_check;
ALTER TABLE cleanup_operation_items ADD CONSTRAINT cleanup_operation_items_resource_type_check
  CHECK (resource_type IN ('onedrive_account', 'sharepoint_site', 'outlook_mailbox',
    'outlook_calendar', 'outlook_contacts', 'channel', 'chat', 'google_my_drive_account'));

-- connection_users is reused as-is (no DDL change) for google_my_drive connections:
--   graph_user_id = the Google Directory API user id (numeric string).
--   upn           = the Workspace user's primary email — also the domain-wide-delegation
--                   impersonation subject used at sync/cleanup time.
-- Same reuse pattern already documented for onedrive/teams/sharepoint in 002_cloud_connections.sql.
COMMENT ON COLUMN connection_users.graph_user_id IS
  'External user/site id. For google_my_drive: the Google Directory API user id.';
COMMENT ON COLUMN connection_users.upn IS
  'External user principal / email / site URL. For google_my_drive: the Workspace user''s primary email (also the DWD impersonation subject).';
