-- Records, per cleanup_operations row, whether the deletions it performed were true permanent
-- deletion (Graph's permanentDelete, skipping the recycle bin / Deleted Items entirely) or the
-- older recycle-bin-recoverable soft delete. Needed because cleanup_operation_items.status =
-- 'deleted' otherwise means something categorically more dangerous the moment permanentDelete
-- ships, with nothing in the schema distinguishing an old (recoverable) run from a new
-- (unrecoverable) one — a report/label layer that just relabels "deleted" as "Permanently Deleted"
-- unconditionally would misrepresent historical soft-delete operations the moment someone re-pulls
-- an old report. One column, operation-level (this is a single global behavior switch applied to
-- every new operation, not a per-tenant/per-item choice — see graph/cleanupDeletion.ts).
ALTER TABLE cleanup_operations ADD COLUMN deletion_mode TEXT NOT NULL DEFAULT 'recycle_bin'
  CHECK (deletion_mode IN ('recycle_bin', 'permanent'));
