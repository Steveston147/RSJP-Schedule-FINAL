# RSJP Schedule — Production Readiness / UAT

Updated: 2026-08-23

## Purpose

Use this checklist before declaring a production revision ready for controlled operational use.

## A. Repository, dependency, and build gate

- [ ] `npm ci` completes successfully.
- [ ] `npm audit --audit-level=high` completes successfully.
- [ ] `npm run validate:security` completes successfully.
- [ ] `npm run lint` completes successfully.
- [ ] `npm run typecheck` completes successfully.
- [ ] `npm run build` completes successfully.
- [ ] No secrets or environment credentials are committed.
- [ ] The production deployment is generated from the intended main-branch revision.

## B. Authentication

- [ ] Unauthenticated access to `/` redirects to `/login`.
- [ ] `/api/auth/session` returns 401 without a valid session.
- [ ] Preview accepts a syntactically valid email plus `12345`.
- [ ] Production rejects an email that is not in `SCHEDULE_AUTH_ALLOWED_EMAILS`.
- [ ] Production fails closed when required authentication environment variables are missing.
- [ ] Production never uses demo mode.
- [ ] Successful Production login uses the allowed email plus the RSJP Schedule-specific PoC password.
- [ ] Normal logout returns to `/login`.
- [ ] Restricted session is rejected after the email is removed from the allowlist.

## C. Browser-local data boundary

- [ ] Existing schedule data survives normal logout and a later successful login in the same browser profile.
- [ ] `端末データを削除してログアウト` asks for confirmation.
- [ ] After confirmed delete-and-logout, `rsjp_schedule_mvp_state_v2` is absent from localStorage.
- [ ] A JSON backup is exported before destructive data-removal UAT.
- [ ] No IndexedDB storage is introduced.
- [ ] Core ScheduleApp does not send the working schedule dataset to a server API.
- [ ] No sensitive personal data is used in the production test dataset.

## D. Security headers

- [ ] `X-Content-Type-Options: nosniff` is present.
- [ ] `Referrer-Policy: no-referrer` is present.
- [ ] `X-Frame-Options: DENY` is present.
- [ ] `Permissions-Policy` disables camera, microphone, and geolocation.
- [ ] `X-Robots-Tag` prevents indexing/archiving.

## E. Basic programme workflow

- [ ] Application loads without a browser console error that blocks operation.
- [ ] A new programme can be created.
- [ ] Programme name can be changed.
- [ ] RSJP / Custom type can be selected.
- [ ] Participant count can be changed.
- [ ] Start and end dates can be changed.
- [ ] Programme selection works after multiple programmes exist.
- [ ] Programme deletion requires confirmation.

## F. Japanese-class generation

- [ ] Default Japanese-class schedule can be configured.
- [ ] Automatic generation produces weekday Japanese classes.
- [ ] Weekend dates are not populated by the default Japanese-class generator.
- [ ] A date-specific override can change start time.
- [ ] A date-specific override can disable Japanese classes for the date.
- [ ] Classroom values appear in generated Japanese-class entries when provided.
- [ ] Regeneration replaces prior auto-generated entries rather than multiplying them.
- [ ] Manually entered events remain after regeneration.

## G. Manual events and persistence

- [ ] Manual event can be created.
- [ ] Date and time are retained after reload.
- [ ] Category, title, location, counts, transport, and notes are retained.
- [ ] Manual event can be edited.
- [ ] Manual event can be deleted.
- [ ] Data survives a normal browser reload.

## H. Backup and restore

- [ ] JSON export downloads a readable JSON file.
- [ ] JSON import restores programme and item data in a clean browser profile.
- [ ] Invalid JSON or orphan references are rejected.
- [ ] Older-backup warning is displayed when applicable.
- [ ] Current data is exported before testing destructive restore scenarios.

## I. Outputs

- [ ] Japanese CSV exports successfully.
- [ ] English CSV exports successfully.
- [ ] Japanese ICS imports into a calendar with representative times correct in Japan time.
- [ ] English ICS imports with representative titles correct.
- [ ] Japanese HTML calendar opens and prints legibly.
- [ ] English HTML calendar opens and prints legibly.
- [ ] English output has been manually reviewed when intended for external circulation.

## J. Final operational sign-off

Record at least:

- tested commit SHA:
- production URL:
- test date:
- browser / OS:
- tester:
- result: GO / NO-GO
- exceptions / known limitations:

A production revision is GO only when blocking defects are absent and the operational limitations in `README.md`, `SECURITY.md`, `SPECIFICATION.md`, and `OPERATIONS.md` are accepted.
