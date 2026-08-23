import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";

const COOKIE_NAME = "rsjp_schedule_session";
const DEMO_SECRET = "rsjp-schedule-preview-demo-session-secret-2026";

type Mode = "demo" | "restricted";
type Payload = { email?: unknown; mode?: unknown; exp?: unknown };

function env(name: string) {
  return String(process.env[name] ?? "").trim();
}

function normalize(email: string) {
  return email.trim().toLowerCase();
}

function effectiveMode(): Mode {
  if (env("VERCEL_ENV").toLowerCase() === "production") return "restricted";
  return env("SCHEDULE_AUTH_MODE").toLowerCase() === "restricted" ? "restricted" : "demo";
}

function sessionSecret() {
  const configured = env("SCHEDULE_AUTH_SESSION_SECRET");
  if (configured.length >= 32) return configured;
  return effectiveMode() === "demo" ? DEMO_SECRET : "";
}

function allowedEmail(email: string) {
  const allowed = new Set(env("SCHEDULE_AUTH_ALLOWED_EMAILS").split(",").map(normalize).filter(Boolean));
  return allowed.has(normalize(email));
}

function decodeBase64Url(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

async function verifySessionToken(token: string) {
  try {
    const [encodedPayload, signature, extra] = token.split(".");
    if (!encodedPayload || !signature || extra) return false;
    const secret = sessionSecret();
    if (secret.length < 32) return false;

    const key = await crypto.subtle.importKey(
      "raw",
      new TextEncoder().encode(secret),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["verify"],
    );
    const valid = await crypto.subtle.verify(
      "HMAC",
      key,
      decodeBase64Url(signature),
      new TextEncoder().encode(encodedPayload),
    );
    if (!valid) return false;

    const payload = JSON.parse(new TextDecoder().decode(decodeBase64Url(encodedPayload))) as Payload;
    if (typeof payload.email !== "string"
      || (payload.mode !== "demo" && payload.mode !== "restricted")
      || typeof payload.exp !== "number") return false;
    if (payload.exp <= Math.floor(Date.now() / 1000) || payload.mode !== effectiveMode()) return false;
    if (payload.mode === "restricted" && !allowedEmail(payload.email)) return false;
    return true;
  } catch {
    return false;
  }
}

export async function proxy(request: NextRequest) {
  const pathname = request.nextUrl.pathname;
  const token = request.cookies.get(COOKIE_NAME)?.value ?? "";
  const valid = token ? await verifySessionToken(token) : false;

  if (pathname === "/login") {
    return valid ? NextResponse.redirect(new URL("/", request.url)) : NextResponse.next();
  }
  if (valid) return NextResponse.next();
  return NextResponse.redirect(new URL("/login", request.url));
}

export const config = {
  matcher: [
    "/((?!api/auth|_next/static|_next/image|favicon.ico|.*\\.(?:png|svg|jpg|jpeg|gif|webp|ico|woff2?|ttf)$).*)",
  ],
};
