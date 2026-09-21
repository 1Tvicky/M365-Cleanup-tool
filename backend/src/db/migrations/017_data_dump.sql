-- Data Dump module: generates real Microsoft 365 objects (OneDrive/SharePoint files+folders, Teams
-- teams+channels, Outlook mailbox items) for demo/migration-testing/regression/performance use —
-- the inverse of Cleaning's delete flow, kept in its own tables on purpose (see routes/dataDump.ts's
-- docstring): a data_dump_operations row must never be mistaken for a cleanup_operations row, at the
-- database level or anywhere above it.
--
-- Row-count design deliberately mirrors cleanup's own scale lesson (cleanup_operation_items is one
-- row per resource, which is fine there because a cleanup selection is bounded by what's already
-- been synced) but Data Dump has no such bound — a Performance Test profile can request millions of
-- files, and this schema must not grow one row per file. So leaf objects (files, mail items) are
-- tracked in BATCHES (data_dump_batches, one row per up-to-N created objects), while STRUCTURAL
-- containers (a folder, a Team, a channel, a library) get their own row in data_dump_containers,
-- since there are only ever dozens-to-low-thousands of those even at "Enterprise" scale — this is
-- also what a future "Delete Generated Data" feature would walk to find real Graph objects to
-- remove, without needing a per-file inventory.

CREATE TABLE data_dump_operations (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id           UUID NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  requested_by        UUID NOT NULL REFERENCES operators(id),
  label               TEXT NOT NULL, -- naming-prefix-derived display label, e.g. "CF-Demo-Migration-001"
  profile             TEXT NOT NULL CHECK (profile IN ('basic_demo', 'migration_demo', 'enterprise_demo', 'performance_test', 'custom')),
  workloads           TEXT[] NOT NULL, -- subset of onedrive/sharepoint/teams/outlook, validated at the API layer against types/dataDump.ts
  -- Full DataDumpConfig (per-workload generation settings, naming prefix, seed, historical date
  -- range) — the worker reloads this from here by operationId, never from the BullMQ payload
  -- itself (see jobs/queue.ts's dataDumpQueue comment).
  config              JSONB NOT NULL,
  status              TEXT NOT NULL DEFAULT 'queued'
                        CHECK (status IN ('queued', 'running', 'paused', 'completed', 'completed_with_errors', 'failed', 'cancelled')),
  requested_items     BIGINT NOT NULL DEFAULT 0,
  created_items       BIGINT NOT NULL DEFAULT 0,
  failed_items        BIGINT NOT NULL DEFAULT 0,
  skipped_items       BIGINT NOT NULL DEFAULT 0,
  total_size_bytes    BIGINT NOT NULL DEFAULT 0,
  -- Delta generation: a later "Add More Data" operation against the same tenant/workloads points
  -- back at the operation it's adding to, so Reports/History can show the lineage.
  parent_operation_id UUID REFERENCES data_dump_operations(id),
  pause_requested_at  TIMESTAMPTZ,
  cancel_requested_at TIMESTAMPTZ,
  started_at          TIMESTAMPTZ,
  completed_at        TIMESTAMPTZ,
  error_message       TEXT,
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX data_dump_operations_tenant_idx ON data_dump_operations (tenant_id, created_at DESC);
CREATE INDEX data_dump_operations_parent_idx ON data_dump_operations (parent_operation_id) WHERE parent_operation_id IS NOT NULL;

-- One row per workload selected within an operation — owns that workload's own status, counters,
-- and resumability checkpoint independently (multi-workload requirement: Teams failing must not
-- roll back a OneDrive workload that already succeeded in the same operation).
CREATE TABLE data_dump_workload_tasks (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  operation_id      UUID NOT NULL REFERENCES data_dump_operations(id) ON DELETE CASCADE,
  -- Deliberately RESTRICT, mirroring cleanup_operation_items.connection_id (migrations/005): this
  -- generated-data provenance record must outlive the connection it was created through.
  connection_id     UUID NOT NULL REFERENCES connections(id) ON DELETE RESTRICT,
  workload          TEXT NOT NULL CHECK (workload IN ('onedrive', 'sharepoint', 'teams', 'outlook')),
  status            TEXT NOT NULL DEFAULT 'pending'
                       CHECK (status IN ('pending', 'running', 'paused', 'completed', 'completed_with_errors', 'failed', 'cancelled', 'skipped')),
  requested_items   BIGINT NOT NULL DEFAULT 0,
  created_items     BIGINT NOT NULL DEFAULT 0,
  failed_items      BIGINT NOT NULL DEFAULT 0,
  skipped_items     BIGINT NOT NULL DEFAULT 0,
  total_size_bytes  BIGINT NOT NULL DEFAULT 0,
  -- Resumability cursor, e.g. {"phase":"files","userIndex":4,"folderIndex":12,"fileIndex":340,
  -- "nextBatchIndex":7} — the worker reads this on (re)start and resumes from here instead of
  -- restarting the whole workload. Shape is worker-owned, not enforced by a CHECK.
  checkpoint        JSONB NOT NULL DEFAULT '{}',
  error_message     TEXT,
  started_at        TIMESTAMPTZ,
  completed_at      TIMESTAMPTZ,
  UNIQUE (operation_id, workload)
);

CREATE INDEX data_dump_workload_tasks_operation_idx ON data_dump_workload_tasks (operation_id);
CREATE INDEX data_dump_workload_tasks_connection_idx ON data_dump_workload_tasks (connection_id);

-- Structural containers this operation created in the real tenant: a OneDrive root folder, a
-- SharePoint document library, a Team, a channel, a mailbox folder. Bounded count (dozens to low
-- thousands) even at huge scale, unlike the leaf objects inside them.
CREATE TABLE data_dump_containers (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  workload_task_id    UUID NOT NULL REFERENCES data_dump_workload_tasks(id) ON DELETE CASCADE,
  kind                TEXT NOT NULL CHECK (kind IN ('root_folder', 'folder', 'team', 'channel', 'site_library', 'mailbox_folder')),
  graph_id            TEXT NOT NULL,
  display_name        TEXT NOT NULL,
  parent_container_id UUID REFERENCES data_dump_containers(id),
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX data_dump_containers_task_idx ON data_dump_containers (workload_task_id);
CREATE INDEX data_dump_containers_parent_idx ON data_dump_containers (parent_container_id) WHERE parent_container_id IS NOT NULL;

-- Batched leaf-object manifests (files, mail items) — one row per up-to-N created objects (N =
-- services/dataDump/batching.ts's DATA_DUMP_BATCH_SIZE), never one row per object. A 1,000,000-file
-- workload at batch size 500 is 2,000 rows here, not 1,000,000.
CREATE TABLE data_dump_batches (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  workload_task_id UUID NOT NULL REFERENCES data_dump_workload_tasks(id) ON DELETE CASCADE,
  container_id     UUID REFERENCES data_dump_containers(id),
  batch_index      INTEGER NOT NULL,
  created_count    INTEGER NOT NULL DEFAULT 0,
  failed_count     INTEGER NOT NULL DEFAULT 0,
  size_bytes       BIGINT NOT NULL DEFAULT 0,
  -- Compact per-item summary, capped detail (name/id/size/status) — the counts above are the
  -- source of truth for totals; this is for "View Generated Resources" drill-down, not audit-grade.
  items            JSONB NOT NULL DEFAULT '[]',
  created_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (workload_task_id, batch_index)
);

CREATE INDEX data_dump_batches_task_idx ON data_dump_batches (workload_task_id);
