# RSJP Schedule

RSJP Schedule is a browser-based operational tool for preparing RSJP and custom short-term programme schedules.

It supports programme setup, Japanese-class schedule generation, date-specific overrides, manual event management, calendar preview, and JSON / CSV / ICS / HTML export.

## Status

The application is intended for controlled internal PoC / operational use. Production authentication is server-side fail-closed and the application remains a browser-local scheduling tool rather than a central institutional system.

## What this application is for

RSJP Schedule helps staff:

- create RSJP or custom programme schedules;
- define programme dates and participant counts;
- generate recurring Japanese-class schedule entries;
- apply day-specific Japanese-class overrides;
- add cultural activities, company visits, buddy events, transport details, rooms, and other operational events;
- review a calendar-style preview;
- export Japanese or English schedule files;
- move or restore working data with JSON backup files.

## Authentication

Production uses a simple PoC authentication boundary:

- allowed email address;
- RSJP Schedule-specific PoC password stored only as a server-side scrypt hash;
- signed HttpOnly session cookie with an 8-hour lifetime;
- Production always forces restricted mode and fails closed if required environment variables are missing.

Preview / local development defaults to demo mode with any syntactically valid email address plus `12345`. Production never falls back to demo mode.

Do not use Microsoft 365, Google, or other existing corporate passwords as the RSJP Schedule PoC password. The intended future production target is company SSO such as Microsoft Entra ID.

Required Production environment variables:

- `SCHEDULE_AUTH_SESSION_SECRET`
- `SCHEDULE_AUTH_ALLOWED_EMAILS`
- `SCHEDULE_AUTH_SHARED_PASSWORD_HASH`

Optional non-Production override: `SCHEDULE_AUTH_MODE=restricted`.

## Important data-storage rule

The working dataset is stored only in the browser's `localStorage` key `rsjp_schedule_mvp_state_v2`.

This means:

- there is no central shared database;
- another computer or browser does not automatically have the same data;
- clearing browser/site data can remove the working dataset;
- Vercel deployment history is not a schedule-data backup;
- normal logout keeps the browser-local working dataset;
- `端末データを削除してログアウト` removes the known RSJP Schedule localStorage dataset before ending the session.

**Export JSON after material changes and before destructive operations or device/browser changes.**

## Human-control rule

Automatically generated schedule entries are templates. The application does not confirm room bookings, transport reservations, staffing, attendance, contracts, or other institutional transactions. Staff must check the schedule against the authoritative source for each operational arrangement before circulation.

## Information-security rule

Do not enter unnecessary personal or sensitive student information. This tool is intended for programme-level scheduling information, not passports, visa documents, health information, payment data, personal addresses, or credentials.

Authentication controls access to the application; it does not change the browser-local storage model. Shared/public computers must not be used for operational schedule data.

## Documentation

- [`docs/SPECIFICATION.md`](docs/SPECIFICATION.md) — functional scope, behaviour, data model, limitations
- [`docs/OPERATIONS.md`](docs/OPERATIONS.md) — operating rules, authentication, backup/restore, export, incident handling
- [`docs/UAT.md`](docs/UAT.md) — production-readiness and browser UAT checklist
- [`SECURITY.md`](SECURITY.md) — security boundary and PoC authentication policy

## Local development

```bash
npm ci
npm run validate:security
npm run lint
npm run typecheck
npm run build
```

Then run:

```bash
npm run dev
```

and open `http://localhost:3000`.

## CI / dependency gate

Pull requests and main-branch changes must pass:

- deterministic `npm ci`;
- `npm audit --audit-level=high`;
- security-boundary validation;
- lint;
- TypeScript typecheck;
- production build.

Documentation-only, Markdown-only, and `.github/**` changes are configured to skip Vercel builds.

## Technology

- Next.js
- React
- TypeScript
- Tailwind CSS
- browser localStorage

## Architectural limitations

The current release has no central database, real-time multi-user collaboration, server-side audit log, or institutional backup service. Authentication is a PoC access boundary, not a replacement for future company SSO or centralized data governance.
