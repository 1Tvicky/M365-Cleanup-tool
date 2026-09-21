import { Router } from "express";
import { z } from "zod";
import { query } from "../db/pool.js";
import { asyncHandler } from "../middleware/asyncHandler.js";
import { requireSession } from "../middleware/auth.js";
import { ApiError } from "../types/index.js";
import type { OperatorRole } from "../types/index.js";
import {
  DATA_DUMP_PROFILES,
  DATA_DUMP_WORKLOADS,
  FILE_TYPE_KEYS,
  type DataDumpConfig,
  type DataDumpOperationStatus,
  type DataDumpProfile,
  type DataDumpWorkload,
} from "../types/dataDump.js";
import { buildDefaultConfig } from "../services/dataDump/profiles.js";
import { computeDataDumpPreview } from "../services/dataDump/preview.js";
import { enqueueDataDumpJob } from "../jobs/queue.js";
import { graphClientForTenant } from "../graph/client.js";
import { listAllUsers } from "../graph/cloudEnumeration.js";
import { filterRealMemberUsers } from "../services/dataDump/targetUsers.js";

/**
 * Data Dump module — generates real Microsoft 365 objects (the inverse of Cleaning's delete flow).
 * Mounted at its own top-level /api/data-dump prefix (app.ts), same precedent as /api/clouds and
 * /api/cleaning: its own contract, never nested under either of those. Reuses this app's existing
 * session auth + tenant_roles RBAC (viewer/cleanup_admin) unchanged — no new auth mechanism, no new
 * roles, per the standing "reuse existing tenant isolation/RBAC" requirement. `cleanup_admin` gates
 * every mutating route here exactly the same way it gates Cleanup's execute/cancel/disconnect
 * actions; Data Dump is a different business function from Cleanup, but not a different trust model.
 */
export const dataDumpRouter = Router();
dataDumpRouter.use(requireSession);

async function requireTenantAccess(tenantId: string, operatorId: string, minRole: OperatorRole): Promise<void> {
  const result = await query<{ role: OperatorRole }>(`SELECT role FROM tenant_roles WHERE tenant_id = $1 AND operator_id = $2`, [tenantId, operatorId]);
  const role = result.rows[0]?.role;
  if (!role || (minRole === "cleanup_admin" && role !== "cleanup_admin")) {
    throw new ApiError(403, "FORBIDDEN", `Requires ${minRole} role on this tenant`);
  }
}

/** Every requested workload must already have an active connection for this tenant — Data Dump never creates a connection, it only ever generates into one that Add Clouds already established. */
async function requireActiveWorkloadConnections(tenantId: string, workloads: DataDumpWorkload[]): Promise<void> {
  const result = await query<{ cloud_type: string }>(`SELECT cloud_type FROM connections WHERE tenant_id = $1 AND status <> 'disconnected'`, [tenantId]);
  const connected = new Set(result.rows.map((r) => r.cloud_type));
  const missing = workloads.filter((w) => !connected.has(w));
  if (missing.length > 0) {
    throw new ApiError(400, "WORKLOAD_NOT_CONNECTED", `Connect ${missing.join(", ")} via Add Clouds before generating data for it`);
  }
}

const fileTypeDistributionSchema = z.record(z.enum(FILE_TYPE_KEYS as unknown as [string, ...string[]]), z.number().nonnegative()).optional();

const dateRangeSchema = z.object({
  mode: z.enum(["last_30_days", "last_6_months", "last_1_year", "last_3_years", "custom"]),
  customStartDate: z.string().optional(),
  customEndDate: z.string().optional(),
});

const positiveInt = z.number().int().positive();
const nonNegativeInt = z.number().int().nonnegative();

const idList = z.array(z.string().min(1));

const oneDriveConfigSchema = z
  .object({
    selectedUserIds: idList,
    userCount: positiveInt,
    rootFolders: positiveInt,
    subFoldersPerFolder: nonNegativeInt,
    maxFolderDepth: positiveInt,
    filesPerFolder: nonNegativeInt,
    targetTotalSizeBytes: nonNegativeInt,
    minFileSizeBytes: positiveInt,
    maxFileSizeBytes: positiveInt,
    fileTypeDistribution: fileTypeDistributionSchema,
    namingStyle: z.enum(["professional", "synthetic"]),
  })
  .partial()
  .optional();

const newSiteSchema = z.object({
  displayName: z.string().min(1).max(255),
  urlSlug: z
    .string()
    .min(1)
    .max(64)
    .regex(/^[a-zA-Z0-9-]+$/, "URL may only contain letters, numbers, and hyphens"),
  template: z.enum(["teamSiteWithoutMicrosoft365Group", "communicationSite"]),
  description: z.string().max(1000).optional(),
  ownerUserIds: idList.optional(),
});

const sharePointConfigSchema = z
  .object({
    selectedSiteIds: idList,
    newSite: newSiteSchema,
    siteCount: positiveInt,
    librariesPerSite: positiveInt,
    foldersPerLibrary: positiveInt,
    subFoldersPerFolder: nonNegativeInt,
    maxFolderDepth: positiveInt,
    filesPerFolder: nonNegativeInt,
    targetTotalSizeBytes: nonNegativeInt,
    minFileSizeBytes: positiveInt,
    maxFileSizeBytes: positiveInt,
    fileTypeDistribution: fileTypeDistributionSchema,
    generatePermissions: z.boolean(),
    namingStyle: z.enum(["professional", "synthetic"]),
  })
  .partial()
  .optional();

const channelConfigSchema = z.object({
  name: z.string().min(1).max(50), // Graph's own documented displayName limit for a channel
  memberUserIds: idList.optional(),
  messages: nonNegativeInt,
  replies: nonNegativeInt,
});

const newTeamSchema = z.object({
  displayName: z.string().min(1).max(255),
  description: z.string().max(1000).optional(),
  visibility: z.enum(["private", "public"]),
  memberUserIds: idList,
});

const teamsConfigSchema = z
  .object({
    selectedTeamIds: idList,
    newTeam: newTeamSchema,
    channels: z.array(channelConfigSchema),
    teamCount: positiveInt,
    membersPerTeam: positiveInt,
    channelsPerTeam: positiveInt,
    messagesPerChannel: nonNegativeInt,
    repliesPerMessage: nonNegativeInt,
    namingStyle: z.enum(["professional", "synthetic"]),
  })
  .partial()
  .optional();

const outlookConfigSchema = z
  .object({
    selectedUserIds: idList,
    userCount: positiveInt,
    emailsPerUser: nonNegativeInt,
    attachmentsPerEmail: z.number().min(0).max(1),
    averageAttachmentSizeBytes: nonNegativeInt,
    calendarEventCount: nonNegativeInt,
    includeRecurringEvents: z.boolean(),
    includeAttendees: z.boolean(),
    contactCount: nonNegativeInt,
    namingStyle: z.enum(["professional", "synthetic"]),
  })
  .partial()
  .optional();

const createSchema = z.object({
  tenantId: z.string().uuid(),
  profile: z.enum(DATA_DUMP_PROFILES as unknown as [DataDumpProfile, ...DataDumpProfile[]]),
  workloads: z.array(z.enum(DATA_DUMP_WORKLOADS as unknown as [DataDumpWorkload, ...DataDumpWorkload[]])).min(1),
  namingPrefix: z.string().min(1).max(60).optional(),
  seed: z.number().int().optional(),
  dateRange: dateRangeSchema.optional(),
  onedrive: oneDriveConfigSchema,
  sharepoint: sharePointConfigSchema,
  teams: teamsConfigSchema,
  outlook: outlookConfigSchema,
  parentOperationId: z.string().uuid().optional(), // delta generation (spec §18): "Add More Data" against a prior operation
});

function mergeConfig(body: z.infer<typeof createSchema>): DataDumpConfig {
  return buildDefaultConfig(body.profile, body.workloads, {
    namingPrefix: body.namingPrefix,
    seed: body.seed,
    dateRange: body.dateRange,
    onedrive: body.onedrive as DataDumpConfig["onedrive"],
    sharepoint: body.sharepoint as DataDumpConfig["sharepoint"],
    teams: body.teams as DataDumpConfig["teams"],
    outlook: body.outlook as DataDumpConfig["outlook"],
  });
}

/**
 * GET /api/data-dump/tenants — purpose-built for the workload/tenant picker on the Data Dump page.
 * Deliberately its own small query rather than reusing Cleaning's connection-listing shape (which
 * that page's own grouping conventions are tailored to) — keeps Data Dump's frontend decoupled from
 * Cleaning's internal display logic, per the standing "keep Data Dump business logic separate"
 * requirement. Only ever returns M365 tenants (onedrive/sharepoint/teams/outlook are all Graph-based
 * — Data Dump does not support Google Workspace tenants in this pass, see docs/data-dump-api.md).
 */
dataDumpRouter.get(
  "/tenants",
  asyncHandler(async (req, res) => {
    // connectionsByWorkload is workload -> connectionId (jsonb) — the Select Resources step needs
    // the actual connectionId per workload to call the shared GET /api/clouds/:id/available-resources
    // endpoint (Add Clouds' own live-Graph resource browser, reused as-is here — see
    // docs/data-dump-api.md's "shared infrastructure" note), not just which workloads exist.
    // adminUpn/adminDisplayName/lastUpdatedAt are picked from whichever of the tenant's connections
    // was touched most recently (DISTINCT ON ... ORDER BY connected_at DESC) — purely for the
    // dashboard card, same "one representative admin per tenant" shape Cleaning's own Landing card
    // already shows (components: CleaningPage.tsx's Landing/TenantGroup).
    const result = await query<{
      tenant_id: string;
      display_name: string;
      workloads: string[];
      connections_by_workload: Record<string, string>;
      admin_upn: string;
      admin_display_name: string | null;
      last_updated_at: string | null;
    }>(
      `SELECT t.id AS tenant_id, t.display_name,
              array_agg(DISTINCT c.cloud_type) FILTER (WHERE c.status <> 'disconnected') AS workloads,
              jsonb_object_agg(c.cloud_type, c.id) FILTER (WHERE c.status <> 'disconnected') AS connections_by_workload,
              (array_agg(c.admin_upn ORDER BY c.connected_at DESC) FILTER (WHERE c.status <> 'disconnected'))[1] AS admin_upn,
              (array_agg(c.admin_display_name ORDER BY c.connected_at DESC) FILTER (WHERE c.status <> 'disconnected'))[1] AS admin_display_name,
              MAX(COALESCE(c.last_synced_at, c.connected_at)) FILTER (WHERE c.status <> 'disconnected') AS last_updated_at
       FROM tenants t
       JOIN tenant_roles tr ON tr.tenant_id = t.id AND tr.operator_id = $1
       LEFT JOIN connections c ON c.tenant_id = t.id AND c.cloud_type = ANY($2::text[])
       WHERE t.m365_tenant_id IS NOT NULL
       GROUP BY t.id, t.display_name
       HAVING array_agg(DISTINCT c.cloud_type) FILTER (WHERE c.status <> 'disconnected') IS NOT NULL
       ORDER BY t.display_name`,
      [req.session!.operatorId, DATA_DUMP_WORKLOADS]
    );
    res.json({
      tenants: result.rows.map((r) => ({
        tenantId: r.tenant_id,
        displayName: r.display_name,
        workloads: r.workloads,
        connectionsByWorkload: r.connections_by_workload,
        adminUpn: r.admin_upn,
        adminDisplayName: r.admin_display_name,
        lastUpdatedAt: r.last_updated_at,
      })),
    });
  })
);

/**
 * GET /api/data-dump/tenants/:tenantId/users — powers the Teams member picker (spec §15) and any
 * other cross-workload "pick existing tenant users" need that isn't already covered by a specific
 * workload's own GET /api/clouds/:id/available-resources (e.g. Teams membership isn't tied to any
 * one connectionId the way OneDrive/Outlook users are). A live Graph listing, server-side paginated/
 * searched the same way listAvailableResources already is, so it plugs into the same DiscoveryTable-
 * style UI. Filtered to real internal member accounts (services/dataDump/targetUsers.ts) — the same
 * guest/deleted-account exclusion the generation workers themselves apply, confirmed live to matter
 * (see jobs/dataDumpWorker.ts).
 */
dataDumpRouter.get(
  "/tenants/:tenantId/users",
  asyncHandler(async (req, res) => {
    await requireTenantAccess(req.params.tenantId!, req.session!.operatorId, "viewer");
    const tenantRow = await query<{ m365_tenant_id: string | null }>(`SELECT m365_tenant_id FROM tenants WHERE id = $1`, [req.params.tenantId]);
    const m365TenantId = tenantRow.rows[0]?.m365_tenant_id;
    if (!m365TenantId) throw new ApiError(400, "NOT_M365_TENANT", "Data Dump only supports Microsoft 365 tenants");

    const client = await graphClientForTenant(m365TenantId);
    const all = filterRealMemberUsers(await listAllUsers(client));
    const search = typeof req.query.search === "string" ? req.query.search.trim().toLowerCase() : "";
    const filtered = search ? all.filter((u) => u.displayName?.toLowerCase().includes(search) || u.upn.toLowerCase().includes(search)) : all;
    const pageSize = Math.min(Number(req.query.pageSize) || 20, 200);
    const page = Math.max(Number(req.query.page) || 1, 1);
    const users = filtered.slice((page - 1) * pageSize, page * pageSize).map((u) => ({ id: u.id, displayName: u.displayName ?? u.upn, secondary: u.upn }));

    res.json({ resources: users, total: filtered.length, page, pageSize });
  })
);

/** POST /api/data-dump/preview — spec §21: never creates real objects, pure computation over the same config the real run would use. */
dataDumpRouter.post(
  "/preview",
  asyncHandler(async (req, res) => {
    const body = createSchema.parse(req.body);
    await requireTenantAccess(body.tenantId, req.session!.operatorId, "viewer");
    const config = mergeConfig(body);
    const preview = computeDataDumpPreview(config.seed ?? 1, body.workloads, config);
    res.json({ preview, config });
  })
);

/** GET /api/data-dump/profile-defaults?profile=&workloads=a,b,c — prefill helper for the frontend's Advanced Options panel. */
dataDumpRouter.get(
  "/profile-defaults",
  asyncHandler(async (req, res) => {
    const profile = z.enum(DATA_DUMP_PROFILES as unknown as [DataDumpProfile, ...DataDumpProfile[]]).parse(req.query.profile);
    const workloads = String(req.query.workloads ?? "")
      .split(",")
      .filter(Boolean)
      .map((w) => z.enum(DATA_DUMP_WORKLOADS as unknown as [DataDumpWorkload, ...DataDumpWorkload[]]).parse(w));
    res.json({ config: buildDefaultConfig(profile, workloads) });
  })
);

dataDumpRouter.post(
  "/",
  asyncHandler(async (req, res) => {
    const body = createSchema.parse(req.body);
    await requireTenantAccess(body.tenantId, req.session!.operatorId, "cleanup_admin");
    await requireActiveWorkloadConnections(body.tenantId, body.workloads);

    const config = mergeConfig(body);
    const preview = computeDataDumpPreview(config.seed ?? 1, body.workloads, config);

    const existingCount = await query<{ count: string }>(`SELECT COUNT(*) FROM data_dump_operations WHERE tenant_id = $1 AND profile = $2`, [body.tenantId, body.profile]);
    const ordinal = Number(existingCount.rows[0]!.count) + 1;
    const label = `${config.namingPrefix}-${body.profile.replace(/_/g, "-")}-${String(ordinal).padStart(3, "0")}`;

    const inserted = await query<{ id: string }>(
      `INSERT INTO data_dump_operations (tenant_id, requested_by, label, profile, workloads, config, requested_items, total_size_bytes, parent_operation_id)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9) RETURNING id`,
      [
        body.tenantId,
        req.session!.operatorId,
        label,
        body.profile,
        body.workloads,
        JSON.stringify(config),
        preview.totalRequestedObjects,
        0,
        body.parentOperationId ?? null,
      ]
    );
    const operationId = inserted.rows[0]!.id;
    await enqueueDataDumpJob({ operationId });

    res.status(202).json({ operationId, status: "queued", label });
  })
);

interface OperationDetailRow {
  id: string;
  tenant_id: string;
  /** Azure AD directory (tenant) ID — surfaced on the report/detail view so an operator can correlate a run against Entra ID/Graph audit logs; Data Dump only ever targets M365 tenants (spec), so this is never null in practice. */
  m365_tenant_id: string | null;
  label: string;
  profile: string;
  workloads: string[];
  status: DataDumpOperationStatus;
  requested_items: string;
  created_items: string;
  failed_items: string;
  skipped_items: string;
  total_size_bytes: string;
  parent_operation_id: string | null;
  started_at: string | null;
  completed_at: string | null;
  error_message: string | null;
  created_at: string;
}

async function loadOperationForOperator(operationId: string, operatorId: string): Promise<OperationDetailRow> {
  const result = await query<OperationDetailRow>(
    `SELECT o.id, o.tenant_id, t.m365_tenant_id, o.label, o.profile, o.workloads, o.status, o.requested_items, o.created_items,
            o.failed_items, o.skipped_items, o.total_size_bytes, o.parent_operation_id, o.started_at, o.completed_at,
            o.error_message, o.created_at
     FROM data_dump_operations o
     JOIN tenant_roles tr ON tr.tenant_id = o.tenant_id AND tr.operator_id = $2
     JOIN tenants t ON t.id = o.tenant_id
     WHERE o.id = $1`,
    [operationId, operatorId]
  );
  const row = result.rows[0];
  if (!row) throw new ApiError(404, "OPERATION_NOT_FOUND", "No such Data Dump operation");
  return row;
}

function toOperationJson(
  row: OperationDetailRow,
  tasks: {
    workload: string;
    status: string;
    requestedItems: number;
    createdItems: number;
    failedItems: number;
    skippedItems: number;
    totalSizeBytes: number;
    subcounts?: Record<string, { requested: number; created: number; failed: number; skipped: number }>;
    errorMessage: string | null;
  }[]
) {
  return {
    id: row.id,
    m365TenantId: row.m365_tenant_id,
    label: row.label,
    profile: row.profile,
    workloads: row.workloads,
    status: row.status,
    requestedItems: Number(row.requested_items),
    createdItems: Number(row.created_items),
    failedItems: Number(row.failed_items),
    skippedItems: Number(row.skipped_items),
    totalSizeBytes: Number(row.total_size_bytes),
    parentOperationId: row.parent_operation_id,
    startedAt: row.started_at,
    completedAt: row.completed_at,
    errorMessage: row.error_message,
    createdAt: row.created_at,
    tasks,
  };
}

dataDumpRouter.get(
  "/history",
  asyncHandler(async (req, res) => {
    const page = Math.max(1, Number(req.query.page) || 1);
    const pageSize = Math.min(100, Math.max(1, Number(req.query.pageSize) || 20));
    const search = typeof req.query.search === "string" ? req.query.search.trim() : "";

    const params: unknown[] = [req.session!.operatorId];
    let searchClause = "";
    if (search) {
      params.push(`%${search}%`);
      searchClause = `AND o.label ILIKE $${params.length}`;
    }
    params.push(pageSize, (page - 1) * pageSize);

    const rows = await query<OperationDetailRow>(
      `SELECT o.id, o.tenant_id, t.m365_tenant_id, o.label, o.profile, o.workloads, o.status, o.requested_items, o.created_items,
              o.failed_items, o.skipped_items, o.total_size_bytes, o.parent_operation_id, o.started_at, o.completed_at,
              o.error_message, o.created_at
       FROM data_dump_operations o
       JOIN tenant_roles tr ON tr.tenant_id = o.tenant_id AND tr.operator_id = $1
       JOIN tenants t ON t.id = o.tenant_id
       WHERE 1=1 ${searchClause}
       ORDER BY o.created_at DESC
       LIMIT $${params.length - 1} OFFSET $${params.length}`,
      params
    );
    const countParams = search ? [req.session!.operatorId, `%${search}%`] : [req.session!.operatorId];
    const total = await query<{ count: string }>(
      `SELECT COUNT(*) FROM data_dump_operations o
       JOIN tenant_roles tr ON tr.tenant_id = o.tenant_id AND tr.operator_id = $1
       WHERE 1=1 ${search ? "AND o.label ILIKE $2" : ""}`,
      countParams
    );

    res.json({
      operations: rows.rows.map((r) => toOperationJson(r, [])),
      total: Number(total.rows[0]!.count),
    });
  })
);

dataDumpRouter.get(
  "/summary",
  asyncHandler(async (req, res) => {
    const result = await query<{ total: string; running: string; completed: string; failed: string }>(
      `SELECT
         COUNT(*) AS total,
         COUNT(*) FILTER (WHERE o.status IN ('queued','running','paused')) AS running,
         COUNT(*) FILTER (WHERE o.status IN ('completed','completed_with_errors')) AS completed,
         COUNT(*) FILTER (WHERE o.status IN ('failed','cancelled')) AS failed
       FROM data_dump_operations o
       JOIN tenant_roles tr ON tr.tenant_id = o.tenant_id AND tr.operator_id = $1`,
      [req.session!.operatorId]
    );
    const row = result.rows[0]!;
    res.json({ total: Number(row.total), running: Number(row.running), completed: Number(row.completed), failed: Number(row.failed) });
  })
);

dataDumpRouter.get(
  "/:operationId",
  asyncHandler(async (req, res) => {
    const operation = await loadOperationForOperator(req.params.operationId!, req.session!.operatorId);
    const tasks = await query<{
      workload: string;
      status: string;
      requested_items: string;
      created_items: string;
      failed_items: string;
      skipped_items: string;
      total_size_bytes: string;
      subcounts: Record<string, { requested: number; created: number; failed: number; skipped: number }>;
      error_message: string | null;
    }>(
      `SELECT workload, status, requested_items, created_items, failed_items, skipped_items, total_size_bytes, subcounts, error_message
       FROM data_dump_workload_tasks WHERE operation_id = $1 ORDER BY workload`,
      [operation.id]
    );
    res.json(
      toOperationJson(
        operation,
        tasks.rows.map((t) => ({
          workload: t.workload,
          status: t.status,
          requestedItems: Number(t.requested_items),
          createdItems: Number(t.created_items),
          failedItems: Number(t.failed_items),
          skippedItems: Number(t.skipped_items),
          totalSizeBytes: Number(t.total_size_bytes),
          subcounts: t.subcounts,
          errorMessage: t.error_message,
        }))
      )
    );
  })
);

/**
 * GET /api/data-dump/:operationId/resources — spec §37's "View Generated Resources": the actual
 * per-resource detail (which user's OneDrive, which SharePoint site/library, which Team/channel,
 * which Outlook mailbox) an operator needs to verify exactly what was created, not just aggregate
 * counts. Returns every container for this operation with its leaf-object counts AGGREGATED
 * DB-side (SUM over data_dump_batches, grouped by container) — never ships individual batch/file
 * rows to the client, so this scales the same way at 10 objects or 10,000,000 (see
 * db/migrations/017_data_dump.sql's batched-manifest design). The frontend builds the container
 * tree itself from each row's parentContainerId.
 */
dataDumpRouter.get(
  "/:operationId/resources",
  asyncHandler(async (req, res) => {
    const operation = await loadOperationForOperator(req.params.operationId!, req.session!.operatorId);
    const containers = await query<{
      id: string;
      kind: string;
      graph_id: string;
      display_name: string;
      parent_container_id: string | null;
      workload: string;
      created_count: string;
      failed_count: string;
      size_bytes: string;
    }>(
      `SELECT c.id, c.kind, c.graph_id, c.display_name, c.parent_container_id, t.workload,
              COALESCE(SUM(b.created_count), 0) AS created_count,
              COALESCE(SUM(b.failed_count), 0) AS failed_count,
              COALESCE(SUM(b.size_bytes), 0) AS size_bytes
       FROM data_dump_containers c
       JOIN data_dump_workload_tasks t ON t.id = c.workload_task_id
       LEFT JOIN data_dump_batches b ON b.container_id = c.id
       WHERE t.operation_id = $1
       GROUP BY c.id, t.workload
       ORDER BY t.workload, c.created_at`,
      [operation.id]
    );
    res.json({
      containers: containers.rows.map((c) => ({
        id: c.id,
        kind: c.kind,
        graphId: c.graph_id,
        displayName: c.display_name,
        parentContainerId: c.parent_container_id,
        workload: c.workload,
        createdCount: Number(c.created_count),
        failedCount: Number(c.failed_count),
        sizeBytes: Number(c.size_bytes),
      })),
    });
  })
);

async function requireOperationMutationAccess(operationId: string, operatorId: string): Promise<OperationDetailRow> {
  const operation = await loadOperationForOperator(operationId, operatorId);
  await requireTenantAccess(operation.tenant_id, operatorId, "cleanup_admin");
  return operation;
}

dataDumpRouter.post(
  "/:operationId/pause",
  asyncHandler(async (req, res) => {
    const operation = await requireOperationMutationAccess(req.params.operationId!, req.session!.operatorId);
    if (!["queued", "running"].includes(operation.status)) {
      throw new ApiError(409, "INVALID_STATE", "Only a queued or running operation can be paused");
    }
    await query(`UPDATE data_dump_operations SET pause_requested_at = now() WHERE id = $1`, [operation.id]);
    res.json({ status: "pause_requested" });
  })
);

dataDumpRouter.post(
  "/:operationId/resume",
  asyncHandler(async (req, res) => {
    const operation = await requireOperationMutationAccess(req.params.operationId!, req.session!.operatorId);
    if (operation.status !== "paused") throw new ApiError(409, "INVALID_STATE", "Only a paused operation can be resumed");
    await query(`UPDATE data_dump_operations SET status = 'queued', pause_requested_at = NULL WHERE id = $1`, [operation.id]);
    await enqueueDataDumpJob({ operationId: operation.id });
    res.json({ status: "queued" });
  })
);

dataDumpRouter.post(
  "/:operationId/cancel",
  asyncHandler(async (req, res) => {
    const operation = await requireOperationMutationAccess(req.params.operationId!, req.session!.operatorId);
    if (["completed", "completed_with_errors", "failed", "cancelled"].includes(operation.status)) {
      throw new ApiError(409, "INVALID_STATE", "This operation has already finished");
    }
    await query(`UPDATE data_dump_operations SET cancel_requested_at = now() WHERE id = $1`, [operation.id]);
    // A paused (or not-yet-picked-up queued) operation has no worker actively polling the cancel
    // flag — flip it to a terminal state directly rather than waiting for a run that isn't happening.
    if (operation.status === "paused" || operation.status === "queued") {
      await query(`UPDATE data_dump_operations SET status = 'cancelled', completed_at = now() WHERE id = $1`, [operation.id]);
    }
    res.json({ status: "cancel_requested" });
  })
);

function csvLine(fields: (string | number)[]): string {
  return fields.map((f) => `"${String(f).replace(/"/g, '""')}"`).join(",") + "\r\n";
}

/** GET /api/data-dump/:operationId/report — spec §38: "data-dump-report-<id>.csv", body carries an explicit "Report Type: Data Dump" line so it is never confusable with a Cleanup report's CSV. */
dataDumpRouter.get(
  "/:operationId/report",
  asyncHandler(async (req, res) => {
    const operation = await loadOperationForOperator(req.params.operationId!, req.session!.operatorId);
    const tasks = await query<{ workload: string; status: string; requested_items: string; created_items: string; failed_items: string; skipped_items: string; total_size_bytes: string }>(
      `SELECT workload, status, requested_items, created_items, failed_items, skipped_items, total_size_bytes FROM data_dump_workload_tasks WHERE operation_id = $1 ORDER BY workload`,
      [operation.id]
    );

    const lines: string[] = [];
    lines.push(csvLine(["Report Type", "Data Dump"]));
    lines.push(csvLine(["Operation ID", operation.id]));
    lines.push(csvLine(["M365 Tenant ID", operation.m365_tenant_id ?? "(unknown)"]));
    lines.push(csvLine(["Label", operation.label]));
    lines.push(csvLine(["Profile", operation.profile]));
    lines.push(csvLine(["Status", operation.status]));
    lines.push(csvLine(["Requested Objects", Number(operation.requested_items)]));
    lines.push(csvLine(["Created Objects", Number(operation.created_items)]));
    lines.push(csvLine(["Failed Objects", Number(operation.failed_items)]));
    lines.push(csvLine(["Skipped Objects", Number(operation.skipped_items)]));
    lines.push(csvLine(["Total Data Size (bytes)", Number(operation.total_size_bytes)]));
    lines.push(csvLine([]));
    lines.push(csvLine(["Workload", "Status", "Requested", "Created", "Failed", "Skipped", "Size (bytes)"]));
    for (const t of tasks.rows) {
      lines.push(csvLine([t.workload, t.status, Number(t.requested_items), Number(t.created_items), Number(t.failed_items), Number(t.skipped_items), Number(t.total_size_bytes)]));
    }

    res.setHeader("Content-Type", "text/csv");
    res.setHeader("Content-Disposition", `attachment; filename="data-dump-report-${operation.id}.csv"`);
    res.send(lines.join(""));
  })
);
