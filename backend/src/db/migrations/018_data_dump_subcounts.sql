-- Per-resource-kind accounting for Data Dump (spec: "requested = created + failed + skipped" must
-- reconcile for EVERY resource kind a workload touches — a OneDrive task's folders and files each
-- need their own requested/created/failed/skipped, not just one blended total; same for a Teams
-- task's teams/channels/channel_members/messages/replies, etc. See services/dataDump/progress.ts's
-- bumpSubcount and types/dataDump.ts's DataDumpSubcounts.
ALTER TABLE data_dump_workload_tasks ADD COLUMN subcounts JSONB NOT NULL DEFAULT '{}';
