-- Tracks each removed file's size so the completed-cleanup report (and progress screen) can state
-- a real "data cleared" total for the client, not just item/file counts. Captured at listing time
-- (before delete — the Graph delete response carries no size), so it's known even for a file whose
-- deletion later fails. Additive only; existing rows default to 0 (unknown, pre-dates this column).

ALTER TABLE cleanup_operation_item_files ADD COLUMN file_size_bytes BIGINT NOT NULL DEFAULT 0;
