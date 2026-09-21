import { describe, expect, it, vi } from "vitest";

const queryMock = vi.fn(async (..._args: unknown[]): Promise<{ rows: any[] }> => ({ rows: [] }));
vi.mock("../../db/pool.js", () => ({ query: (...args: unknown[]) => queryMock(...args) }));

const { BatchAccumulator, DATA_DUMP_BATCH_SIZE, nextBatchIndexFor } = await import("./batching.js");

describe("BatchAccumulator", () => {
  it("does not write to the database until it reaches DATA_DUMP_BATCH_SIZE items (bounded-memory batching, not one row per object)", async () => {
    queryMock.mockClear();
    const acc = new BatchAccumulator("task-1", null, 0);
    for (let i = 0; i < DATA_DUMP_BATCH_SIZE - 1; i++) {
      await acc.add({ name: `f${i}`, graphId: `g${i}`, sizeBytes: 10, status: "created" });
    }
    expect(queryMock).not.toHaveBeenCalled();
  });

  it("flushes automatically exactly once it reaches the batch size", async () => {
    queryMock.mockClear();
    const acc = new BatchAccumulator("task-1", null, 0);
    for (let i = 0; i < DATA_DUMP_BATCH_SIZE; i++) {
      await acc.add({ name: `f${i}`, graphId: `g${i}`, sizeBytes: 10, status: "created" });
    }
    expect(queryMock).toHaveBeenCalledTimes(1);
    const [, params] = queryMock.mock.calls[0]!;
    expect((params as unknown[])[3]).toBe(DATA_DUMP_BATCH_SIZE); // created_count
  });

  it("counts created vs failed items separately in one flushed row", async () => {
    queryMock.mockClear();
    const acc = new BatchAccumulator("task-1", null, 5);
    await acc.add({ name: "a", graphId: "1", sizeBytes: 100, status: "created" });
    await acc.add({ name: "b", graphId: null, sizeBytes: 0, status: "failed" });
    await acc.flush();
    const [, params] = queryMock.mock.calls[0]!;
    const [, , batchIndex, createdCount, failedCount, sizeBytes] = params as unknown[];
    expect(batchIndex).toBe(5);
    expect(createdCount).toBe(1);
    expect(failedCount).toBe(1);
    expect(sizeBytes).toBe(100);
  });

  it("flush() on an empty accumulator is a no-op (never writes an empty batch row)", async () => {
    queryMock.mockClear();
    const acc = new BatchAccumulator("task-1", null, 0);
    await acc.flush();
    expect(queryMock).not.toHaveBeenCalled();
  });

  it("increments currentBatchIndex only after an actual flush", async () => {
    queryMock.mockClear();
    const acc = new BatchAccumulator("task-1", null, 3);
    expect(acc.currentBatchIndex).toBe(3);
    await acc.add({ name: "a", graphId: "1", sizeBytes: 1, status: "created" });
    await acc.flush();
    expect(acc.currentBatchIndex).toBe(4);
  });
});

describe("nextBatchIndexFor", () => {
  it("resumes at 0 when no batches exist yet", async () => {
    queryMock.mockClear();
    queryMock.mockResolvedValueOnce({ rows: [{ max: null }] });
    expect(await nextBatchIndexFor("task-1")).toBe(0);
  });

  it("resumes one past the highest recorded batch_index (resumability)", async () => {
    queryMock.mockClear();
    queryMock.mockResolvedValueOnce({ rows: [{ max: 41 }] });
    expect(await nextBatchIndexFor("task-1")).toBe(42);
  });
});
