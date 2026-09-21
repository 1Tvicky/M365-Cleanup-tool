import { describe, expect, it } from "vitest";
import { createRng, forkSeed, pick, pickWeighted, randomInt, seedFromString } from "./seededRandom.js";

describe("createRng", () => {
  it("is deterministic: the same seed always produces the same sequence", () => {
    const a = createRng(42);
    const b = createRng(42);
    const seqA = Array.from({ length: 10 }, () => a());
    const seqB = Array.from({ length: 10 }, () => b());
    expect(seqA).toEqual(seqB);
  });

  it("different seeds produce different sequences", () => {
    const a = createRng(1);
    const b = createRng(2);
    expect(a()).not.toBe(b());
  });

  it("always returns a value in [0, 1)", () => {
    const rng = createRng(7);
    for (let i = 0; i < 1000; i++) {
      const v = rng();
      expect(v).toBeGreaterThanOrEqual(0);
      expect(v).toBeLessThan(1);
    }
  });
});

describe("seedFromString / forkSeed", () => {
  it("is a pure function of its input (same string -> same seed)", () => {
    expect(seedFromString("hello")).toBe(seedFromString("hello"));
  });

  it("different strings produce different seeds (no trivial collisions for these inputs)", () => {
    expect(seedFromString("onedrive")).not.toBe(seedFromString("sharepoint"));
  });

  it("forking the same base seed with different suffixes yields independent-looking sequences", () => {
    const seedA = forkSeed(123, "workload:onedrive");
    const seedB = forkSeed(123, "workload:sharepoint");
    expect(seedA).not.toBe(seedB);
  });

  it("forking is itself deterministic", () => {
    expect(forkSeed(123, "x")).toBe(forkSeed(123, "x"));
  });
});

describe("randomInt", () => {
  it("never returns a value outside [min, max]", () => {
    const rng = createRng(99);
    for (let i = 0; i < 500; i++) {
      const v = randomInt(rng, 5, 10);
      expect(v).toBeGreaterThanOrEqual(5);
      expect(v).toBeLessThanOrEqual(10);
    }
  });

  it("collapses to min when max <= min", () => {
    const rng = createRng(1);
    expect(randomInt(rng, 7, 7)).toBe(7);
    expect(randomInt(rng, 7, 3)).toBe(7);
  });
});

describe("pick / pickWeighted", () => {
  it("pick always returns an element from the array", () => {
    const rng = createRng(5);
    const items = ["a", "b", "c"];
    for (let i = 0; i < 50; i++) expect(items).toContain(pick(rng, items));
  });

  it("pickWeighted never selects a key with zero/absent weight", () => {
    const rng = createRng(5);
    for (let i = 0; i < 200; i++) {
      const result = pickWeighted(rng, { a: 10, b: 0 }, ["a", "b", "c"]);
      expect(result).toBe("a");
    }
  });

  it("pickWeighted falls back to a uniform pick over `fallback` when every weight is zero/absent", () => {
    const rng = createRng(5);
    const result = pickWeighted(rng, {}, ["x", "y"]);
    expect(["x", "y"]).toContain(result);
  });
});
