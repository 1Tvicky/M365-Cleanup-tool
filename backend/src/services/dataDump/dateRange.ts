import type { DateRangeConfig } from "../../types/dataDump.js";

/** Resolves a DateRangeConfig (spec §17) into a concrete [startMs, endMs] window, ending "now" for every preset mode. */
export function resolveDateRange(config: DateRangeConfig): { startMs: number; endMs: number } {
  const now = Date.now();
  const DAY = 24 * 60 * 60 * 1000;
  switch (config.mode) {
    case "last_30_days":
      return { startMs: now - 30 * DAY, endMs: now };
    case "last_6_months":
      return { startMs: now - 182 * DAY, endMs: now };
    case "last_1_year":
      return { startMs: now - 365 * DAY, endMs: now };
    case "last_3_years":
      return { startMs: now - 3 * 365 * DAY, endMs: now };
    case "custom": {
      const start = config.customStartDate ? Date.parse(config.customStartDate) : now - 30 * DAY;
      const end = config.customEndDate ? Date.parse(config.customEndDate) : now;
      return { startMs: Number.isFinite(start) ? start : now - 30 * DAY, endMs: Number.isFinite(end) ? end : now };
    }
    default:
      return { startMs: now - 30 * DAY, endMs: now };
  }
}

export function randomTimestampInRange(rng: () => number, range: { startMs: number; endMs: number }): string {
  const lo = Math.min(range.startMs, range.endMs);
  const hi = Math.max(range.startMs, range.endMs);
  const ts = lo + rng() * (hi - lo);
  return new Date(ts).toISOString();
}
