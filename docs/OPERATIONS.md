# RSJP Schedule — Operations Manual

Updated: 2026-08-23

## 1. Operating position

RSJP Schedule is an operational support application for preparing and distributing programme schedules. It is not the institutional source of truth for reservations, student records, contracts, payments, or attendance.

Production access is protected by a PoC email/password boundary. The working schedule data itself remains browser-local and is not stored in a central database.

## 2. Login and logout

Production login requires an allowed email address plus the RSJP Schedule-specific PoC password. Do not enter Microsoft 365, Google, or other existing corporate passwords.

Normal `ログアウト` ends the authenticated session but keeps the working dataset in this browser so work can continue later.

`端末データを削除してログアウト` removes the known RSJP Schedule localStorage dataset and then ends the session. Use it when a device is being handed over, retired, or should no longer retain the schedule dataset. Export JSON first if the data may be needed again.

Do not use shared/public computers for operational schedule data.

## 3. Daily operating rules

- Use one clearly named programme record for each operational programme.
- Confirm the programme period and participant count before generating schedule items.
- Treat automatically generated entries as draft templates until checked by a staff member.
- Enter English-specific title/location/notes when an English schedule will be circulated externally.
- Do not store unnecessary personal or sensitive information in notes.
- Export a JSON backup after material changes and before importing another JSON file.
- Keep operationally circulated CSV/ICS/HTML files under normal departmental document-control rules.

## 4. Recommended backup procedure

At minimum, export JSON:

1. after the programme framework is first completed;
2. after major schedule revisions;
3. before importing any JSON file;
4. before clearing browser data or changing devices;
5. before using `端末データを削除してログアウト`;
6. before a programme enters final operational use.

Recommended filename retention pattern:

`ScheduleData_YYYY-MM-DD.json`

If multiple backups are created on the same date, add a departmental version suffix after download.

## 5. Restore procedure

1. Open the application in the intended authenticated browser/profile.
2. Export the current JSON first if it may still be needed.
3. Select JSON import.
4. Choose the intended backup file.
5. If the application warns that the imported file appears older, verify the timestamps before proceeding.
6. After import, check programme name, period, participant count, and several representative schedule entries.
7. Export a fresh JSON after confirming the restored state.

Imported JSON is validated for the expected state structure, duplicate Program IDs, and orphan schedule references before it replaces the current working state.

## 6. Automatic generation procedure

Before pressing automatic generation, confirm:

- programme start date
- programme end date
- Japanese-class enabled/disabled setting
- Japanese-class start time
- lesson duration
- break duration
- number of periods
- number of classes
- class names
- date-specific overrides

After generation, confirm the first day, at least one normal weekday, any override date, and the programme end period.

Manual items are designed to remain when automatic items are regenerated. Nevertheless, review the full schedule after regeneration.

## 7. Export procedure

### JSON
Use for backup and transfer between browser environments.

### CSV
Use for spreadsheet review, sorting, operational worksheets, and controlled downstream processing.

### ICS
Use for calendar import. Confirm representative event times after import into the target calendar application.

### HTML
Use for human-readable calendar preview, printing, or controlled file sharing.

## 8. English output

English output may use manually entered English fields or a limited built-in fallback replacement routine. For external distribution, staff must review names, places, titles, and notes. The fallback routine is not a substitute for human translation or proofreading.

## 9. Browser-local data boundary

The working data is stored in browser localStorage under the known key `rsjp_schedule_mvp_state_v2`.

Therefore:

- another computer does not automatically have the same data;
- another browser profile does not automatically have the same data;
- normal logout does not delete working data;
- clearing site data or selecting the explicit delete-and-logout action removes the working dataset;
- Vercel deployment history is not a schedule-data backup;
- private/incognito browsing should not be used for operational work unless JSON is exported before closing.

The current core ScheduleApp does not use IndexedDB and does not send the working schedule dataset to a server API.

## 10. Information-security precautions

Authentication controls access to the application UI, but the repository remains public and the schedule data remains local to the browser. Operational use must therefore continue to avoid unnecessary sensitive personal data.

Do not enter:

- passport or visa document contents;
- medical or disability details;
- date of birth or personal address;
- private phone numbers or personal email addresses unless there is a separately approved need and control;
- authentication credentials;
- payment information.

## 11. Production authentication configuration

Production always runs in restricted mode. It must fail closed if the required server-side values are missing.

Required Vercel environment variables:

- `SCHEDULE_AUTH_SESSION_SECRET` — at least 32 characters;
- `SCHEDULE_AUTH_ALLOWED_EMAILS` — comma-separated allowed email addresses;
- `SCHEDULE_AUTH_SHARED_PASSWORD_HASH` — scrypt hash of the RSJP Schedule-specific PoC password.

Preview/local defaults to demo mode with a valid email format plus `12345`. Production never falls back to demo mode.

The PoC login uses an 8-hour signed HttpOnly, Secure-in-production, SameSite=Strict cookie. Restricted sessions re-check that the email remains in the current allowlist.

## 12. Human approval before circulation

Before a schedule is circulated externally or treated as operationally final, a staff member should confirm:

- dates and weekdays;
- start/end times;
- rooms and meeting points;
- transport information;
- participant counts where used operationally;
- English wording if applicable;
- whether bookings/reservations were separately confirmed in their authoritative systems.

## 13. Incident handling

If data appears missing or incorrect:

1. stop editing the affected programme;
2. export the current JSON if possible;
3. locate the latest known-good JSON backup;
4. compare the backup date with the current application timestamp;
5. restore only after confirming which version is authoritative;
6. verify representative entries after restore;
7. record the correction in the normal team communication channel if the schedule had already been circulated.

If login fails unexpectedly, first confirm the Production environment variables and allowlist before redeploying.

## 14. Change management

Production changes must be made through a GitHub pull request. Before Preview, GitHub CI should already pass dependency audit, security-boundary validation, lint, typecheck, and build. Preview is reserved for browser/runtime UAT, not routine compilation debugging.

Documentation-only, Markdown-only, and `.github/**` changes should skip Vercel builds.
