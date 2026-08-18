# RSJP Schedule — Production Readiness / UAT

Updated: 2026-08-18

## Purpose

Use this checklist before declaring a production revision ready for operational use.

## A. Repository and build

- [ ] `npm ci` completes successfully.
- [ ] `npm run lint` completes successfully.
- [ ] `npm run build` completes successfully.
- [ ] No secrets or environment credentials are committed.
- [ ] The production deployment is generated from the intended main-branch revision.

## B. Basic programme workflow

- [ ] Application loads without a browser console error that blocks operation.
- [ ] A new programme can be created.
- [ ] Programme name can be changed.
- [ ] RSJP / Custom type can be selected.
- [ ] Participant count can be changed.
- [ ] Start and end dates can be changed.
- [ ] Programme selection works after multiple programmes exist.
- [ ] Programme deletion requires confirmation.

## C. Japanese-class generation

- [ ] Default Japanese-class schedule can be configured.
- [ ] Automatic generation produces weekday Japanese classes.
- [ ] Weekend dates are not populated by the default Japanese-class generator.
- [ ] A date-specific override can change start time.
- [ ] A date-specific override can disable Japanese classes for the date.
- [ ] Classroom values appear in generated Japanese-class entries when provided.
- [ ] Regeneration replaces prior auto-generated entries rather than multiplying them.
- [ ] Manually entered events remain after regeneration.

## D. Manual events

- [ ] Manual event can be created.
- [ ] Date and time are retained after reload.
- [ ] Category, title, location, counts, transport, and notes are retained.
- [ ] Manual event can be edited.
- [ ] Manual event can be deleted.

## E. Backup and restore

- [ ] JSON export downloads a readable JSON file.
- [ ] JSON import restores programme and item data in a clean browser profile.
- [ ] Older-backup warning is displayed when applicable.
- [ ] Current data is exported before testing destructive restore scenarios.

## F. Outputs

- [ ] Japanese CSV exports successfully.
- [ ] English CSV exports successfully.
- [ ] Japanese ICS imports into a calendar with representative times correct in Japan time.
- [ ] English ICS imports with representative titles correct.
- [ ] Japanese HTML calendar opens and prints legibly.
- [ ] English HTML calendar opens and prints legibly.
- [ ] English output has been manually reviewed when intended for external circulation.

## G. Data persistence and safety

- [ ] Data survives a normal browser reload.
- [ ] Staff understand that data is browser-local and not centrally synchronized.
- [ ] A current JSON backup has been saved outside browser localStorage.
- [ ] No sensitive personal data is used in the production test dataset.

## H. Final operational sign-off

Record at least:

- tested commit SHA:
- production URL:
- test date:
- browser / OS:
- tester:
- result: GO / NO-GO
- exceptions / known limitations:

A production revision is GO only when blocking defects are absent and the operational limitations in `SPECIFICATION.md` and `OPERATIONS.md` are accepted.
