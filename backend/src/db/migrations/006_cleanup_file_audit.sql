-- Per-file audit trail for cleanup execution, so the progress screen can show individual
-- file/folder names as they're removed and the completed-cleanup report can list every file,
-- not just "account X was cleaned". Additive only — existing cleanup_operations/items rows are
-- untouched; the new columns default to 0 for any operation created before this migration.

ALTER TABLE cleanup_operation_items ADD COLUMN files_total INTEGER NOT NULL DEFAULT 0;
ALTER TABLE cleanup_operation_items ADD COLUMN files_completed INTEGER NOT NULL DEFAULT 0;

CREATE TABLE cleanup_operation_item_files (
  id                         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  -- Always created/removed together with its parent item — CASCADE here is fine (unlike
  -- cleanup_operation_items.connection_id, this FK doesn't need to survive its parent).
  cleanup_operation_item_id  UUID NOT NULL REFERENCES cleanup_operation_items(id) ON DELETE CASCADE,
  file_name                  TEXT NOT NULL,
  graph_item_id              TEXT NOT NULL,
  status                     TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'deleted', 'already_gone', 'failed')),
  error_message              TEXT,
  completed_at               TIMESTAMPTZ,
  created_at                 TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX cleanup_operation_item_files_item_idx ON cleanup_operation_item_files (cleanup_operation_item_id, status);
-- Powers the progress screen's "recently removed" live feed (most recent completions across the whole operation).
CREATE INDEX cleanup_operation_item_files_completed_idx ON cleanup_operation_item_files (completed_at DESC) WHERE completed_at IS NOT NULL;
