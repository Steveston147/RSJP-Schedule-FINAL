import { createHmac, scryptSync, timingSafeEqual } from "node:crypto";

export type ScheduleAuthMode = "demo" | "restricted";
export type ScheduleAuthUser = { email: string; mode: ScheduleAuthMode };
export type ScheduleAuthSession = ScheduleAuthUser & { expiresAt: number };

type SessionPayload = ScheduleAuthUser & { iat: number; exp: number };

export const SCHEDULE_SESSION_COOKIE_NAME = "rsjp_schedule_session";
export const SCHEDULE_SESSION_TTL_SECONDS = 8 * 60 * 60;

const DEMO_PASSWORD = "12345";
const DEMO_SESSION_SECRET = "rsjp-schedule-preview-demo-session-secret-2026";
const SCRYPT_KEY_LENGTH = 64;

function env(name: string): string {
  return String(process.env[name] ?? "").trim();
}

export function normalizeScheduleEmail(email: string): string {
  return email.trim().toLowerCase();
}

function parseCsv(value: string): Set<string> {
  return new Set(value.split(",").map(normalizeScheduleEmail).filter(Boolean));
}

export function isValidScheduleEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeScheduleEmail(email));
}

function isVercelProduction(): boolean {
  return env("VERCEL_ENV").toLowerCase() === "production";
}

export function getEffectiveScheduleAuthMode(): ScheduleAuthMode {
  if (isVercelProduction()) return "restricted";
  return env("SCHEDULE_AUTH_MODE").toLowerCase() === "restricted" ? "restricted" : "demo";
}

function getSessionSecret(): string {
  const configured = env("SCHEDULE_AUTH_SESSION_SECRET");
  if (configured.length >= 32) return configured;
  if (getEffectiveScheduleAuthMode() === "demo" && !isVercelProduction()) return DEMO_SESSION_SECRET;
  return "";
}

function isScryptHash(value: string): boolean {
  const [scheme, salt, hash, extra] = value.split("$");
  return scheme === "scrypt"
    && !extra
    && /^[0-9a-f]{32}$/i.test(salt ?? "")
    && /^[0-9a-f]{128}$/i.test(hash ?? "");
}

function verifyPasswordHash(password: string, encoded: string): boolean {
  try {
    if (!isScryptHash(encoded)) return false;
    const [, salt, expectedHex] = encoded.split("$");
    const actual = Uint8Array.from(scryptSync(password, salt, SCRYPT_KEY_LENGTH));
    const expected = Uint8Array.from(Buffer.from(expectedHex, "hex"));
    return actual.length === expected.length && timingSafeEqual(actual, expected);
  } catch {
    return false;
  }
}

export function isRestrictedScheduleEmailAllowed(email: string): boolean {
  return parseCsv(env("SCHEDULE_AUTH_ALLOWED_EMAILS")).has(normalizeScheduleEmail(email));
}

export function getPublicScheduleAuthConfig(): { mode: ScheduleAuthMode; configured: boolean; demoHint: string | null } {
  const mode = getEffectiveScheduleAuthMode();
  if (mode === "demo") {
    return { mode, configured: true, demoHint: "Preview / ローカル限定：任意のメール形式 + 12345" };
  }

  const secretReady = getSessionSecret().length >= 32;
  const allowedEmails = parseCsv(env("SCHEDULE_AUTH_ALLOWED_EMAILS"));
  const sharedHash = env("SCHEDULE_AUTH_SHARED_PASSWORD_HASH");
  return {
    mode,
    configured: secretReady && allowedEmails.size > 0 && isScryptHash(sharedHash),
    demoHint: null,
  };
}

export function verifyScheduleCredentials(email: string, password: string): ScheduleAuthUser | null {
  const normalized = normalizeScheduleEmail(email);
  if (!isValidScheduleEmail(normalized) || !password) return null;

  const mode = getEffectiveScheduleAuthMode();
  if (mode === "demo") return password === DEMO_PASSWORD ? { email: normalized, mode } : null;

  const allowedEmails = parseCsv(env("SCHEDULE_AUTH_ALLOWED_EMAILS"));
  const sharedHash = env("SCHEDULE_AUTH_SHARED_PASSWORD_HASH");
  if (allowedEmails.has(normalized) && verifyPasswordHash(password, sharedHash)) {
    return { email: normalized, mode };
  }
  return null;
}

function sign(encodedPayload: string): string {
  const secret = getSessionSecret();
  if (secret.length < 32) throw new Error("RSJP Schedule auth session secret is not configured.");
  return createHmac("sha256", secret).update(encodedPayload).digest("base64url");
}

export function createScheduleSessionToken(user: ScheduleAuthUser, nowSeconds = Math.floor(Date.now() / 1000)): string {
  const payload: SessionPayload = {
    email: normalizeScheduleEmail(user.email),
    mode: user.mode,
    iat: nowSeconds,
    exp: nowSeconds + SCHEDULE_SESSION_TTL_SECONDS,
  };
  const encodedPayload = Buffer.from(JSON.stringify(payload)).toString("base64url");
  return `${encodedPayload}.${sign(encodedPayload)}`;
}

export function verifyScheduleSessionToken(token: string, nowSeconds = Math.floor(Date.now() / 1000)): ScheduleAuthSession | null {
  try {
    const [encodedPayload, signature, extra] = token.split(".");
    if (!encodedPayload || !signature || extra) return null;
    const expected = Uint8Array.from(Buffer.from(sign(encodedPayload)));
    const actual = Uint8Array.from(Buffer.from(signature));
    if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) return null;

    const parsed = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8")) as Partial<SessionPayload>;
    if (typeof parsed.email !== "string"
      || (parsed.mode !== "demo" && parsed.mode !== "restricted")
      || typeof parsed.exp !== "number"
      || parsed.exp <= nowSeconds) return null;

    const email = normalizeScheduleEmail(parsed.email);
    if (!isValidScheduleEmail(email) || parsed.mode !== getEffectiveScheduleAuthMode()) return null;
    if (parsed.mode === "restricted" && !isRestrictedScheduleEmailAllowed(email)) return null;
    return { email, mode: parsed.mode, expiresAt: parsed.exp };
  } catch {
    return null;
  }
}
