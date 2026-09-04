import { Queue } from "bullmq";
import { Redis } from "ioredis";
import { config } from "../config/index.js";

export const connection = new Redis(config.redis.url, { maxRetriesPerRequest: null });

/** One job per confirmed cleanup run. The worker (cleanupWorker.ts) does export-then-delete in order. */
export const cleanupQueue = new Queue("cleanup-jobs", { connection });

export interface CleanupJobPayload {
  jobId: string; // cleanup_jobs.id — worker looks up scope/preview from Postgres by this id
}

export async function enqueueCleanupJob(payload: CleanupJobPayload): Promise<void> {
  await cleanupQueue.add("run-cleanup", payload, {
    attempts: 1, // retries are handled item-by-item inside the worker via runThrottled, not at the job level
    removeOnComplete: { age: 60 * 60 * 24 * 30 },
    removeOnFail: false,
  });
}

/** One job per connection enumeration run (initial connect, or resync). See jobs/cloudSyncWorker.ts. */
export const cloudSyncQueue = new Queue("cloud-sync-jobs", { connection });

export interface CloudSyncJobPayload {
  syncJobId: string; // sync_jobs.id — worker looks up the connection + cloud_type from Postgres by this id
}

export async function enqueueCloudSyncJob(payload: CloudSyncJobPayload): Promise<void> {
  await cloudSyncQueue.add("run-cloud-sync", payload, {
    attempts: 1, // per-item retries happen inside the worker via runThrottled, not at the job level
    removeOnComplete: { age: 60 * 60 * 24 * 30 },
    removeOnFail: false,
  });
}

/** One job per Cleaning discovery scan (teams/channels/chats structure, or message-count calculation). See jobs/cleaningScanWorker.ts. */
export const cleaningScanQueue = new Queue("cleaning-scan-jobs", { connection });

export interface CleaningScanJobPayload {
  scanId: string; // cleaning_scans.id — worker looks up connection + scan_type from Postgres by this id
}

export async function enqueueCleaningScanJob(payload: CleaningScanJobPayload): Promise<void> {
  await cleaningScanQueue.add("run-cleaning-scan", payload, {
    attempts: 1,
    removeOnComplete: { age: 60 * 60 * 24 * 30 },
    removeOnFail: false,
  });
}

/** One job per confirmed Cleaning cleanup/deletion run. Deliberately named differently from the legacy "cleanup-jobs" queue above — see jobs/cleanupExecutionWorker.ts. */
export const cleanupExecutionQueue = new Queue("cleanup-execution-jobs", { connection });

export interface CleanupExecutionJobPayload {
  operationId: string; // cleanup_operations.id — worker looks up tenant/items from Postgres by this id
}

export async function enqueueCleanupExecutionJob(payload: CleanupExecutionJobPayload): Promise<void> {
  await cleanupExecutionQueue.add("run-cleanup-execution", payload, {
    attempts: 1,
    removeOnComplete: { age: 60 * 60 * 24 * 30 },
    removeOnFail: false,
  });
}
