import { readFileSync } from "node:fs";

const auth = readFileSync("lib/auth.ts", "utf8");
const middleware = readFileSync("proxy.ts", "utf8");
const scheduleApp = readFileSync("components/ScheduleApp.tsx", "utf8");
const nextConfig = readFileSync("next.config.js", "utf8");
const loginRoute = readFileSync("app/api/auth/login/route.ts", "utf8");

const checks = [
  [auth.includes('if (isVercelProduction()) return "restricted";'), "Production must force restricted auth mode"],
  [auth.includes('const DEMO_PASSWORD = "12345";'), "Preview demo password contract must remain explicit"],
  [auth.includes("SCHEDULE_AUTH_SESSION_SECRET"), "Session secret must come from server environment"],
  [auth.includes("SCHEDULE_AUTH_ALLOWED_EMAILS"), "Restricted email allowlist must be server-side"],
  [auth.includes("SCHEDULE_AUTH_SHARED_PASSWORD_HASH"), "Production password must be stored as a hash"],
  [middleware.includes('pathname === "/login"'), "Proxy must protect application routes"],
  [middleware.includes("api/auth"), "Auth API routes must remain reachable before login"],
  [loginRoute.includes("checkScheduleLoginRateLimit"), "Login must apply rate limiting before credential verification"],
  [scheduleApp.includes('const STORAGE_KEY = "rsjp_schedule_mvp_state_v2";'), "Known localStorage boundary must remain explicit"],
  [scheduleApp.includes("localStorage.getItem(STORAGE_KEY)"), "Schedule data remains browser-local"],
  [scheduleApp.includes("localStorage.setItem(STORAGE_KEY"), "Schedule persistence remains browser-local"],
  [!scheduleApp.includes("indexedDB"), "No IndexedDB persistence is expected"],
  [!scheduleApp.includes("fetch("), "Core schedule data must not be sent to a server API"],
  [nextConfig.includes("X-Content-Type-Options"), "Security headers must include nosniff"],
  [nextConfig.includes("X-Frame-Options"), "Security headers must prevent framing"],
  [nextConfig.includes("Permissions-Policy"), "Permissions policy must remain explicit"],
];

const failures = checks.filter(([ok]) => !ok).map(([, message]) => message);
if (failures.length) {
  console.error("Security boundary validation failed:");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log("Security boundary validation passed.");
