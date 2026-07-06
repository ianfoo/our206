# our206 Architecture

## System shape

```
Intake channels                    Projections
(Signal DM, web form,              (website, Google Calendar,
 shorthand paste, CLI,              JSON API, ICS, digests,
 automated discovery)               chat answers)
        │                                ▲
        ▼                                │
  Classification ──► Extraction ──► Enrichment
        │                                ▲
        ▼                                │
  Proposal queue ──► Human review ──► Canonical store ──► Sync workers
```

Everything revolves around the canonical database. Messaging platforms,
calendars, and the website are clients of the platform, not the platform.

Clean separation between five concerns:

1. **Ingestion** — accept raw submissions from any channel; record
   provenance; never write canonical records directly
2. **Canonical storage** — events, venues, artists, recommendations, users
3. **Enrichment** — augment canonical/proposed records with external
   knowledge, always with confidence and source
4. **Publication** — project canonical data outward (site, calendar, feeds,
   digests)
5. **User interaction** — web UI, chat, review tooling, auth

## Runtime & storage

- **Application: a single Go service.** HTTP API, background jobs, sync
  workers, and bot integrations in one deployable binary. (See
  decisions.md D1.)
- **Database: SQLite.** Expected to be sufficient indefinitely at this
  community's scale; favors simplicity, testability, portability, trivial
  backup, and near-zero operational overhead.
- **Python retrieval sidecar (later phase only).** When semantic retrieval
  arrives, a sidecar owns web retrieval, extraction, chunking, embeddings,
  vector storage, and reranking. The Go app keeps business logic, APIs,
  auth, persistence, and orchestration. Not built until a phase needs it.

## Domain model

### Venue

Canonical venue: `id, name, address, city, state, postal_code, lat, lng,
website, place_id, place_confidence, last_checked_at`.

### VenueAlias

Many aliases → one venue (`croc` → The Crocodile, `sodo showbox` → Showbox
SoDo). Seeded directly from today's `Venue Map` sheet and the baked-in
fallback maps in `calendar-sync/config.js`.

### Artist

Canonical performer: `name, aliases, genres, biography, website`, plus
cached enrichment (representative videos, interviews, articles). Reused
across events so enrichment work compounds.

### Event

Canonical event: `id, uid, date, title/artist, venue_id, score, notes,
ticket_url, status, created_at, updated_at`. Future: genres, event type,
supporting acts, enrichment status, discovery source, confidence.

- `uid` preserves the current deterministic identity
  (date + normalized artist + normalized venue) so calendar reconciliation
  carries over unchanged.
- `score` is the excitement score: `✅`/`!` count capped at 4, displayed as
  🔥 (as the sheet does today).
- "Past" is a status/view derived from date — not a row moved to another
  table (replaces the `Past Concerts` archive sheet).

### Submission

One intake action: `source` (signal, web, paste, form, cli, discovery),
`submitter_id`, `raw payload` (text and/or image refs), `received_at`.
The original message and images are retained for review and audit.

### EventProposal

The unit of review. One submission may fan out into many proposals (e.g. a
ten-date tour poster). Carries: extracted fields, per-field confidence,
enrichment results and sources, reasoning notes, link back to submission,
and lifecycle state:

`pending → approved | rejected | expired | superseded`

- Pending proposals whose date passes auto-expire.
- Expired/rejected proposals are kept for audit but excluded from active
  queues.
- Approval creates or updates a canonical Event.

The current `Incoming Raw` dispositions (appended / exact duplicate / fuzzy
duplicate / ignored / needs review / error) map onto this: an
unambiguous parse with no duplicate can be auto-approved per channel trust
policy; duplicates and ambiguity become pending proposals with the conflict
attached.

### MusicRecommendation

Shared music that isn't an event: `url, platform, artist, album, title,
notes, contributor, original_sharer, created_at`. Searchable; linked to
Artist when resolvable.

### User & Identity

`User` (display name, roles: submitter/editor/administrator, trust level)
with linked `Identity` rows per platform (signal, discord, telegram, web).
Many identities → one user. Auth = one-time challenge delivered over a
linked messaging identity → short-lived web session. No passwords.

### CalendarSync

Per (event, calendar): external event ID, sync hash, last synced at.
Direct replacement for the UID-marker-in-description mechanism, but stored
relationally instead of hidden in the calendar event body.

## Pipelines

### Classification

A forwarded message is classified before extraction:

- **Event** (performer + venue + date signals) → event extraction
- **Music** (YouTube/Bandcamp/SoundCloud/Spotify/Relisten URL) →
  recommendation record
- **Unknown** ("everyone should check out this band") → assisted
  clarification workflow with the submitter

### Extraction

- Text: port of the existing shorthand grammar (tolerant of missing colon,
  arbitrary capitalization, `- commentary`, parentheticals, `!` and `✅`),
  plus LLM extraction for free-form text.
- Images (screenshots, posters, flyers): multimodal extraction of artist,
  tour, venue, city/state, date, time, on-sale/presale dates, restrictions.
  Multiple events in one image fan out to multiple proposals.
- Geographic pass: identify the local show among tour dates; offer nearby
  regional shows (Portland, Vancouver) as optional proposals.

### Normalization & dedup

Ports today's behavior:

- Artist: de-shoutify (all-caps → title case; mixed case untouched), known
  typo corrections.
- Venue: alias resolution (below).
- Duplicates: exact = date + normalized venue + normalized artist; fuzzy =
  same date/venue with overlapping primary artist ("Tom Hamilton" vs
  "Tom Hamilton x Swindler"). Fuzzy matches are never silently merged —
  always presented to an editor.

### Venue canonicalization

Not a RAG problem. The pipeline:

```
raw name → alias table hit? → done
         → LLM normalization ("Croc" → "The Crocodile Seattle music venue")
         → Places API lookup (authoritative)
         → candidate scoring
         → auto-accept if high confidence, else editor review
         → store canonical venue + new alias
```

- The Places provider is the source of truth for name, address, lat/lng,
  and a stable place ID. The LLM only does query expansion/normalization
  and is never trusted for factual address data.
- Refresh: re-fetch each venue by stored place ID infrequently (~annually)
  to detect moves, closures, renames. Cheap because it's once per canonical
  venue.

### Enrichment

After canonicalization, per-proposal/per-event enrichment: ticket links,
venue/artist pages, presales, opening acts, age restrictions, seated vs GA,
accessibility, parking/transit, representative videos, interviews, bios.

- Every enrichment carries confidence + source.
- Cached per artist/venue with refresh timestamps, so repeat performers are
  free.
- Sources: structured APIs and Places first; general web search (e.g. Brave
  Search API) for the long tail; full retrieval pipeline only in the
  semantic-search phase.

### Review

Heavy review happens in the web app; messaging stays lightweight. Queue
capabilities: list/filter/search/sort pending, bulk approve/reject, approve
all from one submission, inspect original message/images, compare extracted
vs enriched values. Ambiguity produces a clarification request to the
submitter ("No Seattle show found, but Portland and Vancouver — add
either?").

### Synchronization

One-way, application → Google Calendar, preserving today's semantics:
create missing, update changed (via sync hash), delete orphans, stable UID
identity. Runs debounced after approvals/edits and on a daily schedule.
Other projections (ICS, JSON feed, site pages) render straight from the
canonical store.

### Semantic retrieval (future phase)

When conversational discovery needs it:

```
Go app → Python sidecar → search → fetch → extract → chunk → embed
       → vector store → hybrid retrieval (semantic + keyword) → rerank
       → evidence → LLM
```

Vector records carry both embedding and metadata payload (source URL,
title, text, fetched_at, expires_at, source type, domain, tags) — metadata
drives filtering, freshness, authorization, provenance. Guiding principle
throughout: **use the simplest tool that is authoritative for the task**
(Places for facts, LLM for messy input, embeddings for similarity, SQL for
filtering, reranker for evidence selection).
