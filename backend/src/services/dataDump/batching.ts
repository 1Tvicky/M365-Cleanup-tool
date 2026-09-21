import { query } from "../../db/pool.js";

/**
 * Row-count design (see db/migrations/017_data_dump.sql): one data_dump_batches row per up-to-N
 * created leaf objects (files, mail items), never one row per object — a 1,000,000-file workload at
 * this batch size writes 2,000 rows, not 1,000,000. Also bounds how much this worker holds in memory
 * at once: `items` never exceeds this length before being flushed.
 */
export const DATA_DUMP_BATCH_SIZE = 500;

export interface BatchItemSummary {
  name: string;
  graphId: string | null;
  sizeBytes: number;
  status: "created" | "failed";
}

/** Accumulates created/failed leaf-object summaries and flushes them to data_dump_batches in fixed-size batches — the worker's write-side counterpart to the batched-manifest schema design. */
export class BatchAccumulator {
  private items: BatchItemSummary[] = [];
  private nextBatchIndex: number;
  private readonly workloadTaskId: string;
  private readonly containerId: string | null;

  constructor(workloadTaskId: string, containerId: string | null, startingBatchIndex: number) {
    this.workloadTaskId = workloadTaskId;
    this.containerId = containerId;
    this.nextBatchIndex = startingBatchIndex;
  }

  async add(item: BatchItemSummary): Promise<void> {
    this.items.push(item);
    if (this.items.length >= DATA_DUMP_BATCH_SIZE) await this.flush();
  }

  async flush(): Promise<void> {
    if (this.items.length === 0) return;
    const createdCount = this.items.filter((i) => i.status === "created").length;
    const failedCount = this.items.filter((i) => i.status === "failed").length;
    const sizeBytes = this.items.reduce((sum, i) => sum + i.sizeBytes, 0);
    await query(
      `INSERT INTO data_dump_batches (workload_task_id, container_id, batch_index, created_count, failed_count, size_bytes, items)
       VALUES ($1, $2, $3, $4, $5, $6, $7)
       ON CONFLICT (workload_task_id, batch_index) DO NOTHING`,
      [this.workloadTaskId, this.containerId, this.nextBatchIndex, createdCount, failedCount, sizeBytes, JSON.stringify(this.items)]
    );
    this.nextBatchIndex++;
    this.items = [];
  }

  get currentBatchIndex(): number {
    return this.nextBatchIndex;
  }
}

/** Resumability: the next batch index to continue from after a restart — one past the highest batch_index already written for this workload task. */
export async function nextBatchIndexFor(workloadTaskId: string): Promise<number> {
  const result = await query<{ max: number | null }>(`SELECT MAX(batch_index) AS max FROM data_dump_batches WHERE workload_task_id = $1`, [workloadTaskId]);
  const max = result.rows[0]?.max;
  return max == null ? 0 : max + 1;
}
