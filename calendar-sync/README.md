# Our206 Concert Calendar Automation

A lightweight Google Apps Script system for managing concert and event
tracking in Google Sheets and synchronizing upcoming events into a
Google Calendar.

This project uses:
- Google Sheets as the canonical data store
- Google Forms as an optional intake UI
- Google Calendar as a downstream projection/sync target
- Apps Script triggers for automation orchestration

The system supports:
- Event intake from a Google Form
- Venue normalization
- UID-based calendar reconciliation
- Debounced calendar synchronization
- Automatic movement of past events into an archive sheet
- Processing of messy contributor updates via an ingestion sheet
- Automatic event sorting and maintenance tasks

---

## Sheet Structure

### Concerts

Primary canonical event store.

Expected columns:

| Date | Artist | Venue | Skoi Rating | Notes | Ticket Link | UID | Added On |
|------|------|------|------|------|------|------|------|

Default header row: 3

### Past Concerts

Archive sheet for events whose date is before today.

Must match the schema of `Concerts`.

### Incoming Raw

Optional ingestion sheet used for batch contributor updates and parsing
semi-structured source lines.

Typical source format:

```text
5/23: SOME BAND @ Neumos ✅
```

The ingestion processor:
- parses dates/artists/venues
- normalizes venues
- de-shoutifies artist names
- deduplicates against canonical events
- appends valid rows into `Concerts`

### Venue Map

Optional normalization/reference sheet.

Expected columns:

| Raw Venue | Normalized Venue | Address (optional) |

Used to:
- normalize venue names
- enrich calendar event locations with addresses

Fallback venue mappings are also embedded in source.

---

## Trigger Model

### onEdit → debounced sync

Edits to the canonical event sheet schedule a delayed sync operation.

Rapid edits coalesce into a single eventual calendar synchronization.

Flow:

```text
edit
  → our206_onEdit
      → maybeSetAddedOn_
      → scheduleDebouncedSync_
          → create delayed trigger
              → our206_debouncedSync
                  → syncUpcomingEvents
```

Debounced sync triggers self-clean after execution.

### Daily maintenance

A scheduled maintenance trigger:
- moves past events into `Past Concerts`
- compacts/sorts sheets
- runs a calendar sync

---

## Calendar Sync Behavior

Upcoming events are synchronized into the configured Google Calendar
using deterministic UID reconciliation.

UIDs are derived from:
- event date
- normalized artist name
- normalized venue name

The sync process:
- creates missing events
- updates changed events
- removes orphaned events
- preserves event identity across edits

Calendar event descriptions contain an embedded hidden UID marker.

---

## Google Form Intake

A Google Form can append events into the canonical `Concerts` sheet.

Current intake mapping:

| Form Field | Concerts Column |
|------|------|
| Artist/Event | Artist |
| Date | Date |
| Venue | Venue |
| Ticket Link | Ticket Link |
| Additional Notes | Notes |

The intake flow:
- normalizes venues
- applies row formatting
- writes Added On timestamps
- preserves UID handling for downstream sync

---

## Local Development

This project is managed locally using `clasp`.

Useful commands:

```bash
clasp pull
clasp push
clasp open
```

Recommended:
- use Git locally
- avoid editing large logic changes directly in the Apps Script editor

`.clasprc.json` should not be committed.

---

## File Organization

The Apps Script project is split into multiple source files for sanity.
Apps Script still executes everything in a shared global namespace.

Suggested conceptual boundaries:

- config.js
- triggers.js
- calendar.js
- ingest.js
- venues.js
- sheets.js
- util.js

---

## Operational Notes

This system intentionally favors:
- low operational overhead
- free/near-free infrastructure
- recoverability
- human readability
- lightweight automation

It is not intended to be a high-scale multi-user application.

The current implementation works well for:
- personal concert tracking
- small-group collaborative event management
- lightweight calendar publishing workflows
