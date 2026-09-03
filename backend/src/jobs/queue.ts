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
