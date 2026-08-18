# RSJP Schedule

RSJP Schedule is a browser-based operational tool for preparing RSJP and custom short-term programme schedules.

It supports programme setup, Japanese-class schedule generation, date-specific overrides, manual event management, calendar preview, and JSON / CSV / ICS / HTML export.

## Status

Production-readiness documentation was refreshed on 2026-08-18. The application should be treated as **controlled operational use** until the UAT checklist in `docs/UAT.md` has been completed against the deployed production revision.

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

## Important data-storage rule

The working dataset is stored in the browser's `localStorage`.

This means:

- there is no central shared database;
- another computer or browser does not automatically have the same data;
- clearing browser/site data can remove the working dataset;
- Vercel deployment history is not a schedule-data backup.

**Export JSON after material changes and before destructive operations or device/browser changes.**

## Human-control rule

Automatically generated schedule entries are templates. The application does not confirm room bookings, transport reservations, staffing, attendance, contracts, or other institutional transactions. Staff must check the schedule against the authoritative source for each operational arrangement before circulation.

## Information-security rule

Do not enter unnecessary personal or sensitive student information. This tool is intended for programme-level scheduling information, not passports, visa documents, health information, payment data, personal addresses, or credentials.

## Documentation

- [`docs/SPECIFICATION.md`](docs/SPECIFICATION.md) — functional scope, behaviour, data model, limitations
- [`docs/OPERATIONS.md`](docs/OPERATIONS.md) — operating rules, backup/restore, export, incident handling
- [`docs/UAT.md`](docs/UAT.md) — production-readiness and browser UAT checklist

## Local development

```bash
npm ci
npm run dev
```

Then open `http://localhost:3000`.

Before merging a production change:

```bash
npm run lint
npm run build
```

## Technology

- Next.js
- React
- TypeScript
- Tailwind CSS
- browser localStorage

## Current architectural limitations

The current release has no built-in authentication, central database, real-time multi-user collaboration, server-side audit log, or institutional backup service. These are deliberate boundaries that must be understood when deciding where and how to deploy the application.
