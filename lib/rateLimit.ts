type Bucket = { count: number; resetAt: number };

const WINDOW_MS = 15 * 60 * 1000;
const PAIR_LIMIT = 5;
const CLIENT_LIMIT = 100;
const MAX_PAIR_BUCKETS = 2000;
const MAX_CLIENT_BUCKETS = 500;
const pairBuckets = new Map<string, Bucket>();
const clientBuckets = new Map<string, Bucket>();

function sweepExpired(store: Map<string, Bucket>, now: number) {
  for (const [key, bucket] of store) {
    if (bucket.resetAt <= now) store.delete(key);
  }
}

function capStore(store: Map<string, Bucket>, maxSize: number) {
  while (store.size >= maxSize) {
    const oldestKey = store.keys().next().value as string | undefined;
    if (!oldestKey) break;
    store.delete(oldestKey);
  }
}

function consume(store: Map<string, Bucket>, key: string, limit: number, maxSize: number, now: number) {
  const existing = store.get(key);
  if (!existing || existing.resetAt <= now) {
    if (existing) store.delete(key);
    capStore(store, maxSize);
    store.set(key, { count: 1, resetAt: now + WINDOW_MS });
    return { allowed: true, retryAfter: 0 };
  }

  existing.count += 1;
  store.set(key, existing);
  return {
    allowed: existing.count <= limit,
    retryAfter: Math.max(1, Math.ceil((existing.resetAt - now) / 1000)),
  };
}

export function checkScheduleLoginRateLimit(clientKey: string, email: string, now = Date.now()) {
  sweepExpired(clientBuckets, now);
  sweepExpired(pairBuckets, now);

  const client = clientKey || "unknown";
  const normalizedEmail = email.trim().toLowerCase() || "invalid";
  const clientResult = consume(clientBuckets, client, CLIENT_LIMIT, MAX_CLIENT_BUCKETS, now);
  if (!clientResult.allowed) {
    return { allowed: false, retryAfter: clientResult.retryAfter };
  }

  const pairResult = consume(pairBuckets, `${client}|${normalizedEmail}`, PAIR_LIMIT, MAX_PAIR_BUCKETS, now);
  return {
    allowed: pairResult.allowed,
    retryAfter: pairResult.allowed ? 0 : pairResult.retryAfter,
  };
}
