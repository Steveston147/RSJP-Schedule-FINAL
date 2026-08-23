type Bucket = { count: number; resetAt: number };

const WINDOW_MS = 15 * 60 * 1000;
const PAIR_LIMIT = 5;
const CLIENT_LIMIT = 100;
const pairBuckets = new Map<string, Bucket>();
const clientBuckets = new Map<string, Bucket>();

function consume(store: Map<string, Bucket>, key: string, limit: number, now: number) {
  const existing = store.get(key);
  if (!existing || existing.resetAt <= now) {
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
  const client = clientKey || "unknown";
  const normalizedEmail = email.trim().toLowerCase() || "invalid";
  const clientResult = consume(clientBuckets, client, CLIENT_LIMIT, now);
  const pairResult = consume(pairBuckets, `${client}|${normalizedEmail}`, PAIR_LIMIT, now);
  const allowed = clientResult.allowed && pairResult.allowed;
  return {
    allowed,
    retryAfter: allowed ? 0 : Math.max(clientResult.retryAfter, pairResult.retryAfter),
  };
}
