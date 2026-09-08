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
    /**
     * Override the default concurrency. Some Graph resources throttle far more aggressively than
     * the general default (observed: Teams chat listing enforces roughly 10 requests per 10
     * seconds tenant-wide) — running the usual batch of 20 against one of those guarantees most of
     * the batch gets throttled immediately, wasting retries instead of just... not sending as many
     * requests at once.
     */
    batchSize?: number;
    /** Tags the throttle/backoff console warnings below with which workload made the call (e.g. "OneDrive", "Outlook-Mail") — see the isolation assessment's logging requirement. Defaults to the generic "graph" for call sites that don't pass one. */
    label?: string;
    /**
     * Overrides config.graph.callTimeoutMs for this call site; pass `null` to skip the per-item
     * timeout entirely. The default 30s deadline assumes `execute(item)` is (or tightly wraps) a
     * single atomic Graph call — it's the wrong tool when `execute(item)` is itself a whole
     * multi-step pipeline (list-everything, seed DB rows, delete-everything) with no fixed upper
     * bound on how long that legitimately takes, e.g. cleanupExecutionWorker.ts's per-item loop:
     * a mailbox with a few thousand messages finishes in seconds, but one with millions of
     * messages cannot possibly finish inside 30s no matter how healthy Graph is — every atomic
     * Graph call inside that pipeline already gets its own correctly-scoped timeout via its own
     * nested runThrottled call, so this outer layer doesn't need (and must not impose) one too.
     */
    callTimeoutMs?: number | null;
  } = {}
): Promise<void> {
  const { maxRetries, baseBackoffMs, maxBackoffMs } = config.graph;
  const callTimeoutMs = opts.callTimeoutMs === undefined ? config.graph.callTimeoutMs : opts.callTimeoutMs;
  const batchSize = opts.batchSize ?? config.graph.batchSize;
  const label = opts.label ?? "graph";

  for (let i = 0; i < items.length; i += batchSize) {
    if (await opts.isCancelled?.()) return;

    const batch = items.slice(i, i + batchSize);
    await Promise.all(
      batch.map(async (item) => {
        let attempt = 0;
        while (true) {
          try {
            const value = await withTimeout(execute(item), callTimeoutMs);
            opts.onItemSettled?.(item, { ok: true, value });
            return;
          } catch (err) {
            const retryAfter = extractRetryAfterSeconds(err);
            if (retryAfter !== null && attempt < maxRetries) {
              // Deliberately just status/attempt/wait — never the request body or auth header,
              // which is where a token could leak.
              console.warn(`[${label}] throttled (429/503), attempt ${attempt + 1}/${maxRetries}, waiting ${retryAfter}s`);
              opts.onThrottle?.({ retryAfterSeconds: retryAfter });
              await sleep(retryAfter * 1000);
              attempt++;
              continue;
            }
            if ((isTransient(err) || err instanceof GraphCallTimeoutError) && attempt < maxRetries) {
              const backoff = Math.min(baseBackoffMs * 2 ** attempt, maxBackoffMs);
              const jitter = Math.random() * backoff * 0.25;
              const reason = err instanceof GraphCallTimeoutError ? "timeout" : "transient error";
              console.warn(`[${label}] ${reason}, attempt ${attempt + 1}/${maxRetries}, backing off ${Math.round(backoff + jitter)}ms`);
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

class GraphCallTimeoutError extends Error {}

/**
 * Neither the Graph SDK client nor Node's fetch applies a default timeout — a connection that
 * never receives a response (observed against a large real tenant enumerating chats) would
 * otherwise leave `execute(item)` pending forever, stalling its whole batch — and therefore the
 * entire job — permanently. Racing a timeout here guarantees every item eventually settles one way
 * or another, feeding back into the same retry/backoff path as any other transient failure.
 */
function withTimeout<T>(promise: Promise<T>, ms: number | null): Promise<T> {
  if (ms === null) return promise;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new GraphCallTimeoutError(`Graph call exceeded ${ms}ms`)), ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      }
    );
  });
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
