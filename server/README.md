# our206 server

The our206 platform service (Phase 1 of [the roadmap](../docs/roadmap.md)):
a single Go binary that owns the canonical event database and projects it
outward. It replaces the Google Sheet as the source of truth while keeping
the workflows the spreadsheet proved — the contributor shorthand, venue
normalization, duplicate detection, and one-way Google Calendar sync — with
test-pinned behavior compatibility.

## Layout

- `cmd/our206d` — service + admin CLI entrypoint
- `internal/store` — SQLite canonical store (venues, aliases, events,
  submissions, proposals) plus seed data carried over from the prototype
- `internal/parse` — contributor shorthand grammar
  (`3/20: Band @ Venue ✅✅`) and spreadsheet date parsing
- `internal/normalize` — artist de-shoutifying, comparison keys,
  excitement score (✅/! → 🔥, max 4)
- `internal/importer` — intake pipeline: parse → normalize → venue
  resolution → dedup → events + proposal records
- `internal/gcal` — one-way Google Calendar reconciliation, UID-marker
  compatible with calendars populated by the Apps Script system
- `internal/web` — JSON API, authenticated import endpoint, minimal HTML
  event listing (placeholder until the Phase 2 website)

## Running

```bash
go test ./...
go build ./cmd/our206d

# import a shorthand batch (from a file or stdin)
./our206d import --submitter someone < batch.txt

# one-time migration from a spreadsheet CSV export
# (columns: Date, Artist, Venue, Skoi Rating, Notes, Ticket Link)
./our206d import-csv concerts.csv
./our206d import-csv past-concerts.csv

# inspect
./our206d events
./our206d pending

# calendar sync (see below for credentials)
./our206d sync --dry-run
./our206d sync

# HTTP server
OUR206_ADMIN_TOKEN=$(openssl rand -hex 16) ./our206d serve
```

Every invocation applies the schema and (idempotently) seeds the venue and
alias tables, so there is no separate migrate step.

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `OUR206_DB` | `our206.db` | SQLite database path |
| `OUR206_ADDR` | `:8080` | HTTP listen address |
| `OUR206_ADMIN_TOKEN` | *(unset)* | Bearer token for `POST /api/import`; endpoint is disabled when unset |
| `OUR206_CALENDAR_ID` | *(unset)* | Google Calendar id for `sync` |
| `GOOGLE_APPLICATION_CREDENTIALS` | *(unset)* | Service-account key file for `sync` |

## HTTP surface

- `GET /` — minimal HTML listing of upcoming shows
- `GET /healthz`
- `GET /api/events?from=YYYY-MM-DD&to=YYYY-MM-DD` (defaults: today → +2y)
- `GET /api/venues`
- `POST /api/import` — text/plain shorthand lines; requires
  `Authorization: Bearer $OUR206_ADMIN_TOKEN`

## Calendar sync

Sync is one-way (database → calendar) and reconciles by the hidden
`[our206_uid]:` marker in event descriptions — the same scheme the Apps
Script used, so it takes over an existing calendar cleanly. Events without
a marker are never touched. To grant access: create a Google Cloud service
account, download its key, and share the target calendar with the service
account's email address with "Make changes to events" permission.

## Deploying (Coolify)

The Dockerfile builds a static binary in a distroless image. Point Coolify
at this directory, attach a persistent volume at `/data` (the SQLite
database lives there), and set the environment variables above. The
container listens on 8080.
