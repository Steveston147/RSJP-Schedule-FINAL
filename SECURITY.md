# RSJP Schedule Security Boundary

Updated: 2026-08-23

## Current scope

RSJP Schedule is a controlled internal PoC / operational support application. It is not a system of record and is not approved for passport, visa, health, payment, or other high-sensitivity personal data.

## Authentication

Production:

- always forces `restricted` mode;
- requires `SCHEDULE_AUTH_SESSION_SECRET` of at least 32 characters;
- requires at least one address in `SCHEDULE_AUTH_ALLOWED_EMAILS`;
- requires `SCHEDULE_AUTH_SHARED_PASSWORD_HASH` in the supported scrypt format;
- fails closed when required configuration is missing;
- uses an 8-hour signed HttpOnly, Secure-in-production, SameSite=Strict session cookie;
- re-checks the current email allowlist when a restricted session is verified.

Preview/local:

- defaults to demo mode;
- accepts a syntactically valid email plus `12345`;
- never changes Production into demo mode.

The PoC password must be application-specific. Company Microsoft 365, Google, or other existing account passwords must not be collected or reused.

## Login abuse control

The login route applies a 15-minute in-memory rate limit before scrypt verification:

- 5 attempts per client + email pair;
- 100 attempts per client overall.

This is a PoC guard, not a durable distributed rate-limit service. Formal deployment should rely on company SSO and platform/WAF controls.

## Browser-local data boundary

The schedule working dataset is stored under `localStorage` key `rsjp_schedule_mvp_state_v2`.

Verified application boundary:

- no central schedule database;
- no IndexedDB use in the core ScheduleApp;
- no `fetch()` call in the core ScheduleApp that transmits the working schedule dataset;
- normal logout keeps browser-local schedule data;
- explicit delete-and-logout removes the known schedule localStorage key;
- JSON export remains the backup/transfer mechanism.

Authentication cookies do not contain schedule data.

## Security headers

All routes receive:

- `X-Content-Type-Options: nosniff`;
- `Referrer-Policy: no-referrer`;
- `X-Frame-Options: DENY`;
- `Permissions-Policy: camera=(), microphone=(), geolocation=()`;
- `X-Robots-Tag: noindex, nofollow, noarchive`.

Metadata also requests `noindex,nofollow`.

## CI security gate

Before browser Preview/UAT, GitHub CI must pass:

1. `npm ci`;
2. `npm audit --audit-level=high`;
3. `npm run validate:security`;
4. `npm run lint`;
5. `npm run typecheck`;
6. `npm run build`.

Dependency versions are not upgraded merely for uniformity with other applications. Upgrades must be driven by a concrete audit/build finding and then revalidated.

On 2026-08-23 the existing Next.js 15.5.21 lockfile produced 13 high-severity and 2 moderate audit findings. The audit output identified Next.js transitive PostCSS/sharp findings that required the Next.js 16.3.2 path. A GitHub-runner candidate using Next.js 16.3.2, React 19.2.7, React DOM 19.2.7, ESLint 9.39.1, and eslint-config-next 16.3.2 then completed with `npm audit` reporting 0 vulnerabilities and passed the security validator, typecheck, and production build. The final candidate also passed lint after limiting two legacy Hook compatibility exceptions to `components/ScheduleApp.tsx` and removing one stale lint-suppression comment.

## Future target

For formal company deployment, replace PoC credentials with company-managed SSO such as Microsoft Entra ID, inherit corporate MFA/Conditional Access, and review the browser-local data model under the company's information-governance process.
