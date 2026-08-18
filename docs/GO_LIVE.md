# RSJP Schedule — Production Readiness / GO Decision

Decision date: 2026-08-18

## Decision

**GO — controlled operational use**

RSJP Schedule may be used for operational schedule preparation under the controls documented in this repository.

This is not an unrestricted enterprise-system approval. The application remains a browser-local scheduling tool and must be operated within the documented data and human-review boundaries.

## Evidence reviewed

- GitHub Actions clean install (`npm ci`): PASS
- GitHub Actions lint: PASS
- GitHub Actions production build: PASS
- Security baseline verification: PASS
- Next.js upgraded to 15.5.21
- React / React DOM upgraded to 19.1.9
- Vercel Preview production build: READY
- Completion Ceremony automatic-generation regression: PASS; the ceremony is generated on the programme end date
- Operational README, functional specification, operations manual, and UAT checklist are present

See `CI_RESULT.md` and `SECURITY_UPGRADE.md` for machine-generated verification records.

## Conditions of use

1. Working data is browser `localStorage`, not a shared institutional database.
2. Export JSON after material changes and before destructive operations, imports, browser resets, or device changes.
3. Do not store sensitive personal information, credentials, payment information, passport/visa contents, or medical information.
4. Automatically generated schedule items are draft operational templates and must be checked by staff.
5. Room bookings, transport orders, staffing, attendance, payments, and other institutional transactions remain authoritative in their respective source systems.
6. Review English output before external distribution; fallback translation is limited.
7. CSV, ICS, and HTML are distribution/output formats; JSON is the application backup/transfer format.

## Known non-blocking items

- The production build reports React Hooks dependency warnings in `ScheduleApp.tsx`. They do not fail lint/build and were not changed during this readiness pass because changing dependency behaviour without a dedicated regression test could alter form-reset or preview-refresh behaviour. They should be addressed in a separate maintenance PR with behavioural tests.
- Browserslist data may report that its browser database is old. This does not block the current static application build and can be refreshed in normal dependency maintenance.
- A full human browser walkthrough remains recommended at the start of the first live programme, especially for JSON restore and external CSV/ICS/HTML output review.

## Operational status

The application is approved for **controlled internal operational use** with human final judgement and routine JSON backup.

PR #2 should be merged only while the final CI and Vercel checks remain green/ready and no new blocking review comments are introduced.
