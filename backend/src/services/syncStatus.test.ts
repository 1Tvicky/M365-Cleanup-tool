import { describe, expect, it } from "vitest";
import { computeSyncStatus } from "./syncStatus.js";

describe("computeSyncStatus", () => {
  it("returns 'completed' when there are no sub-resources at all", () => {
    expect(computeSyncStatus([])).toBe("completed");
  });

  it("returns 'queued' only when every sub-resource is still queued", () => {
    expect(computeSyncStatus(["queued"])).toBe("queued");
    expect(computeSyncStatus(["queued", "queued"])).toBe("queued");
  });

  it("returns 'running' once at least one sub-resource has started", () => {
    expect(computeSyncStatus(["queued", "running"])).toBe("running");
    expect(computeSyncStatus(["running"])).toBe("running");
    expect(computeSyncStatus(["running", "completed"])).toBe("running");
  });

  it("returns 'completed' when every sub-resource finished successfully", () => {
    expect(computeSyncStatus(["completed"])).toBe("completed");
    expect(computeSyncStatus(["completed", "completed", "completed"])).toBe("completed");
  });

  it("returns 'completed_with_errors' when some, but not all, sub-resources failed", () => {
    expect(computeSyncStatus(["completed", "failed"])).toBe("completed_with_errors");
    expect(computeSyncStatus(["completed", "cancelled", "completed"])).toBe("completed_with_errors");
  });

  it("returns 'completed_with_errors' when a sub-resource itself finished 'completed_with_errors' — e.g. OneDrive accounts with no provisioned drive — not a clean 'completed'", () => {
    expect(computeSyncStatus(["completed_with_errors"])).toBe("completed_with_errors");
    expect(computeSyncStatus(["completed", "completed_with_errors"])).toBe("completed_with_errors");
  });

  it("returns 'failed' when every sub-resource ended unsuccessfully", () => {
    expect(computeSyncStatus(["failed"])).toBe("failed");
    expect(computeSyncStatus(["failed", "cancelled"])).toBe("failed");
  });

  it("does not return 'failed' just because one sub-resource had partial errors alongside a fully failed one", () => {
    expect(computeSyncStatus(["failed", "completed_with_errors"])).toBe("completed_with_errors");
  });
});
