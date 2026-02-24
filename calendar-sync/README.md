# Calendar Sync Script

This directory contains a Google Apps Script used to sync a Google Sheet with a Google Calendar for `our206`.

- Script file: `our206-calendar-sync.gs`
- Sync direction: Sheet -> Calendar (for upcoming events), plus maintenance helpers

## Purpose

The script:

1. Reads the `Concerts` sheet as the source of truth for upcoming shows.
2. Creates/updates/deletes tagged calendar events so the calendar matches the sheet.
3. Generates and persists stable UIDs (in a `UID` column and event description marker) to match rows to events across runs.
4. Moves past events from `Concerts` to `Past Concerts`.
5. Provides utility functions for importing past concerts and purging future events (manual use only).
6. Normalizes venue aliases (for example `Nectar` -> `Nectar Lounge`) and writes canonical names back to the sheet.

## Expected Sheet Structure

The spreadsheet should include:

- `Concerts` sheet
- `Past Concerts` sheet (created automatically if missing)

Header detection is automatic, but headers should include at least:

- `Date`
- `Artist`
- `Venue`
- `Skoi` (rating)
- `Notes`
- `Ticket`

The script also ensures a `UID` column exists.

## Setup

1. Open your Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. Replace existing code with `our206-calendar-sync.gs`.
4. Save the project.
5. In Apps Script, set project timezone to match your sheet/calendar timezone (for example `America/Los_Angeles`).
6. In `Project Settings -> Script properties`, add:
   - Key: `OUR206_CALENDAR_ID`
   - Value: your calendar ID (for example `our206wa@gmail.com`)
7. Run `setUpOur206` once from the Apps Script editor.
8. Approve requested permissions.

After setup, reload the sheet and use the `Our206` custom menu.

## Daily Use

From the `Our206` menu:

- `Sync now`: live sync to calendar
- `Dry run sync`: logs proposed creates/updates/deletes without calendar writes
- `Move past events to Past Concerts`: archives old rows
- `Move past events + Sync now`: archive then sync
- `Show last run log`: view the latest sync summary

The script also installs:

- On-edit debounce trigger (`our206_onEdit` -> delayed sync)
- Daily maintenance trigger (`our206_dailyMaintenance`)

## UID Matching

Each event is tagged in description with:

- `[our206_uid]:<hash>`

That UID is also written to the `UID` sheet column. Matching uses this marker, so event titles/notes can change while identity stays stable.

## Venue Normalization

The script normalizes venue names before:

- UID generation
- location/address lookup
- calendar create/update

Normalization rules live in `our206-calendar-sync.gs`:

- `VENUE_ALIASES` for exact/slang mappings
- `VENUE_REGEX_RULES` for fuzzy matching variants

Matching order is:

1. exact canonical match
2. exact alias match
3. regex fuzzy rules (ordered; first match wins)
4. exact normalized canonical fallback

Examples:

- `Nectar` -> `Nectar Lounge`
- `Sodo Showbox` -> `Showbox SoDo`
- `Shitbox` -> `Showbox SoDo`

During sync, if a row uses an alias, the canonical venue name is written back into the `Concerts` sheet automatically.

## Safety Notes

- `purgeAllFutureEvents_our206_paced()` is destructive: it deletes future events on the configured calendar.
- Use purge only for recovery/reset workflows.
- Keep `OUR206_CALENDAR_ID` in Script Properties rather than hardcoding it in source.

## Troubleshooting

- Wrong dates/day shifts:
  - Confirm Apps Script project timezone matches the spreadsheet timezone.
  - Run `Dry run sync` and compare logged `CREATED: YYYY-MM-DD` values with sheet dates.
- Unexpected large create/delete counts:
  - Run one live sync, then run dry run again; counts usually converge after reconciliation.
- Missing menu:
  - Reload the sheet after saving script and running setup.
