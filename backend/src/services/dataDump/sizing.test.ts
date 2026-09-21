import { describe, expect, it } from "vitest";
import { createRng } from "./seededRandom.js";
import { estimateFileCountForTargetSize, pickFileSizeBytes } from "./sizing.js";

describe("pickFileSizeBytes", () => {
  it("never returns a value outside [min, max] across many samples", () => {
    const rng = createRng(11);
    for (let i = 0; i < 2000; i++) {
      const v = pickFileSizeBytes(rng, 1000, 5_000_000, 100_000);
      expect(v).toBeGreaterThanOrEqual(1000);
      expect(v).toBeLessThanOrEqual(5_000_000);
    }
  });

  it("collapses to a single value when min === max", () => {
    const rng = createRng(1);
    expect(pickFileSizeBytes(rng, 2048, 2048, 2048)).toBe(2048);
  });

  it("clamps a swapped min/max without throwing", () => {
    const rng = createRng(1);
    const v = pickFileSizeBytes(rng, 5000, 1000, 3000);
    expect(v).toBeGreaterThanOrEqual(1000);
    expect(v).toBeLessThanOrEqual(5000);
  });
});

describe("estimateFileCountForTargetSize", () => {
  it("returns 0 when there is no target size", () => {
    expect(estimateFileCountForTargetSize(0, 1000)).toBe(0);
  });

  it("estimates a sensible file count for a target size", () => {
    expect(estimateFileCountForTargetSize(1_000_000, 100_000)).toBe(10);
  });

  it("never returns 0 for a positive target (always at least 1 file)", () => {
    expect(estimateFileCountForTargetSize(1, 1_000_000_000)).toBe(1);
  });
});
