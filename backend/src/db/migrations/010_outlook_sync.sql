-- "Sync Now" (cleaning_sync_operations, 007_cleaning_sync_operations.sql) hardcodes one column per
-- cloud sub-resource rather than a generic list — Outlook needs its own slot, same shape as
-- onedrive_sync_job_id (both use the sync_jobs table, unlike Teams' cleaning_scans).

ALTER TABLE cleaning_sync_operations ADD COLUMN outlook_sync_job_id UUID REFERENCES sync_jobs(id);
