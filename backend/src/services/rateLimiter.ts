import { config } from "../config/index.js";

export interface ThrottleState {
  retryAfterSeconds: number;
}

/**
 * Runs `items` through `execute` in batches, honoring Graph throttling (429/503 + Retry-After)
 * with exponential backoff + jitter. `onThrottle` lets callers surface pause state to the job's
 * progress record instead of the batch silently stalling (docs/api-spec.md rate-limiting contract).
 */
export async function runThrottled<T, R>(
  items: T[],
  execute: (item: T) => Promise<R>,
  opts: {
    onItemSettled?: (item: T, result: { ok: true; value: R } | { ok: false; error: unknown }) => void;
    onThrottle?: (state: ThrottleState) => void;
    isCancelled?: () => boolean | Promise<boolean>;
  } = {}
): Promise<void> {
  const { batchSize, maxRetries, baseBackoffMs, maxBackoffMs } = config.graph;

  for (let i = 0; i < items.length; i += batchSize) {
    if (await opts.isCancelled?.()) return;

    const batch = items.slice(i, i + batchSize);
    await Promise.all(
      batch.map(async (item) => {
        let attempt = 0;
        while (true) {
          try {
            const value = await execute(item);
            opts.onItemSettled?.(item, { ok: true, value });
            return;
          } catch (err) {
            const retryAfter = extractRetryAfterSeconds(err);
            if (retryAfter !== null && attempt < maxRetries) {
              opts.onThrottle?.({ retryAfterSeconds: retryAfter });
              await sleep(retryAfter * 1000);
              attempt++;
              continue;
            }
            if (isTransient(err) && attempt < maxRetries) {
              const backoff = Math.min(baseBackoffMs * 2 ** attempt, maxBackoffMs);
              const jitter = Math.random() * backoff * 0.25;
              await sleep(backoff + jitter);
              attempt++;
              continue;
            }
            opts.onItemSettled?.(item, { ok: false, error: err });
            return;
          }
        }
      })
    );
  }
}

function extractRetryAfterSeconds(err: unknown): number | null {
  const status = (err as { statusCode?: number })?.statusCode;
  if (status !== 429 && status !== 503) return null;
  const header = (err as { headers?: Record<string, string> })?.headers?.["retry-after"];
  const parsed = header ? Number(header) : NaN;
  return Number.isFinite(parsed) ? parsed : 5;
}

function isTransient(err: unknown): boolean {
  const status = (err as { statusCode?: number })?.statusCode;
  return status === 502 || status === 503 || status === 504;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
