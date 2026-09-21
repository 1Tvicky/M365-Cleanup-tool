/**
 * Deterministic generation (spec §19): given the same DataDumpConfig and seed, naming/content
 * selection follows the same sequence every run. Not cryptographic — this only needs to be a stable,
 * fast, pure function of its seed, never security-sensitive.
 */

/** mulberry32 — small, fast, good-enough statistical quality for content variety, fully deterministic. */
export function createRng(seed: number): () => number {
  let a = seed >>> 0;
  return function rng() {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Derives a stable 32-bit seed from an arbitrary string — used to fork a per-workload/per-batch RNG off one operation-level seed without them all producing identical sequences. */
export function seedFromString(input: string): number {
  let h = 2166136261 >>> 0; // FNV-1a
  for (let i = 0; i < input.length; i++) {
    h ^= input.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

/** Combines an operation seed with a string suffix (workload name, batch index, ...) to fork an independent-looking but fully reproducible sub-sequence. */
export function forkSeed(baseSeed: number, suffix: string): number {
  return seedFromString(`${baseSeed}:${suffix}`);
}

export function randomInt(rng: () => number, minInclusive: number, maxInclusive: number): number {
  if (maxInclusive <= minInclusive) return minInclusive;
  return minInclusive + Math.floor(rng() * (maxInclusive - minInclusive + 1));
}

export function pick<T>(rng: () => number, items: readonly T[]): T {
  return items[Math.floor(rng() * items.length)]!;
}

/** Weighted pick over a Partial<Record<K,number>> distribution — missing/zero weights are never chosen; falls back to a uniform pick over `fallback` if every weight is zero/absent. */
export function pickWeighted<K extends string>(rng: () => number, distribution: Partial<Record<K, number>>, fallback: readonly K[]): K {
  const entries = Object.entries(distribution).filter(([, w]) => typeof w === "number" && w! > 0) as [K, number][];
  if (entries.length === 0) return pick(rng, fallback);
  const total = entries.reduce((sum, [, w]) => sum + w, 0);
  let roll = rng() * total;
  for (const [key, weight] of entries) {
    roll -= weight;
    if (roll <= 0) return key;
  }
  return entries[entries.length - 1]![0];
}

/** A per-instance RNG that also exposes a plain seed number, for content generators that want to derive further sub-seeds without exposing the closure. */
export interface SeededSource {
  rng: () => number;
  seed: number;
}

export function createSeededSource(seed: number): SeededSource {
  return { rng: createRng(seed), seed };
}
