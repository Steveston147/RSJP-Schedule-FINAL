# RSJP Schedule — Operations Manual

Updated: 2026-08-18

## 1. Operating position

RSJP Schedule is an operational support application for preparing and distributing programme schedules. It is not the institutional source of truth for reservations, student records, contracts, payments, or attendance.

## 2. Daily operating rules

- Use one clearly named programme record for each operational programme.
- Confirm the programme period and participant count before generating schedule items.
- Treat automatically generated entries as draft templates until checked by a staff member.
- Enter English-specific title/location/notes when an English schedule will be circulated externally.
- Do not store unnecessary personal or sensitive information in notes.
- Export a JSON backup after material changes and before importing another JSON file.
- Keep operationally circulated CSV/ICS/HTML files under normal departmental document-control rules.

## 3. Recommended backup procedure

At minimum, export JSON:

1. after the programme framework is first completed;
2. after major schedule revisions;
3. before importing any JSON file;
4. before clearing browser data or changing devices;
5. before a programme enters final operational use.

Recommended filename retention pattern:

`ScheduleData_YYYY-MM-DD.json`

If multiple backups are created on the same date, add a departmental version suffix after download.

## 4. Restore procedure

1. Open the application in the intended browser/profile.
2. Export the current JSON first if it may still be needed.
3. Select JSON import.
4. Choose the intended backup file.
5. If the application warns that the imported file appears older, verify the timestamps before proceeding.
6. After import, check programme name, period, participant count, and several representative schedule entries.
7. Export a fresh JSON after confirming the restored state.

## 5. Automatic generation procedure

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

## 6. Export procedure

### JSON
Use for backup and transfer between browser environments.

### CSV
Use for spreadsheet review, sorting, operational worksheets, and controlled downstream processing.

### ICS
Use for calendar import. Confirm representative event times after import into the target calendar application.

### HTML
Use for human-readable calendar preview, printing, or controlled file sharing.

## 7. English output

English output may use manually entered English fields or a limited built-in fallback replacement routine. For external distribution, staff must review names, places, titles, and notes. The fallback routine is not a substitute for human translation or proofreading.

## 8. Data-loss precautions

The working data is stored in browser localStorage. Therefore:

- do not assume another computer has the same data;
- do not assume another browser profile has the same data;
- do not clear site data before exporting JSON;
- do not rely on Vercel deployment history as a data backup;
- do not use private/incognito browsing for operational work unless JSON is exported before closing.

## 9. Information-security precautions

The application repository is public and the application has no built-in authentication. Operational use must therefore avoid sensitive personal data. Use programme-level information and generic operational notes only.

Do not enter:

- passport or visa document contents;
- medical or disability details;
- date of birth or personal address;
- private phone numbers or personal email addresses unless there is a separately approved need and control;
- authentication credentials;
- payment information.

## 10. Human approval before circulation

Before a schedule is circulated externally or treated as operationally final, a staff member should confirm:

- dates and weekdays;
- start/end times;
- rooms and meeting points;
- transport information;
- participant counts where used operationally;
- English wording if applicable;
- whether bookings/reservations were separately confirmed in their authoritative systems.

## 11. Incident handling

If data appears missing or incorrect:

1. stop editing the affected programme;
2. export the current JSON if possible;
3. locate the latest known-good JSON backup;
4. compare the backup date with the current application timestamp;
5. restore only after confirming which version is authoritative;
6. verify representative entries after restore;
7. record the correction in the normal team communication channel if the schedule had already been circulated.

## 12. Change management

Production changes should be made through a GitHub pull request and should not be treated as complete until the build/lint checks and a representative browser UAT are complete.
