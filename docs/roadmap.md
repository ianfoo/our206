# our206 Roadmap

Each phase ships something usable on its own and demotes one more
responsibility of the spreadsheet. The Apps Script system keeps running
until Phase 1 replaces it — nothing breaks mid-migration.

## Phase 0 — today (baseline)

Google Sheet is canonical; Apps Script handles ingestion, normalization,
dedup, calendar sync, archival; our206.com is an iframe of the Google
Calendar. Works, but storage/UI/workflow/sync are all fused into one
spreadsheet.

## Phase 1 — canonical core *(in progress — see `server/`)*

**Goal: the application database becomes the source of truth.**

- Go service + SQLite with the core domain model (Event, Venue, VenueAlias,
  Submission, EventProposal, CalendarSync)
- One-time migration: import `Concerts`/`Past Concerts` rows, `Venue Map`
  entries, and the baked-in alias/address maps from
  `calendar-sync/config.js`
- Port the shorthand parser, artist/venue normalization, and exact/fuzzy
  dedup (behavior-compatible with `incoming-raw.js`, covered by tests)
- Port one-way Google Calendar sync with the same UID scheme, tracked in
  the `CalendarSync` table instead of hidden description markers
- Minimal admin surface (CLI and/or barebones web) to paste shorthand
  batches, review fuzzy duplicates, and edit events

**Exit criteria:** spreadsheet is read-only/retired; calendar stays
correct; a contributor batch flows paste → review → publish → calendar.

## Phase 2 — website as the primary experience

**Goal: our206.com stops being an iframe.**

- Public read API (JSON) + ICS feed from the canonical store
- Fast, polished, responsive site: upcoming events list, search/filter,
  event detail with ticket links and venue info, venue pages with maps
- Keep the existing visual identity (skoiberg backdrop, Space Needle
  favicon) as a starting point; UX/UI design passes for discovery flow
- About/documentation section (contribution guide, formats, scoring,
  privacy, FAQ)

**Exit criteria:** the site is where people actually look things up; the
embedded calendar is gone or secondary.

## Phase 3 — Signal intake + AI extraction

**Goal: forwarding a message is all a contributor has to do.**

- Signal bot via DM (bot does not join the group; see decisions.md D5)
- Message classification: event / music / unknown
- LLM extraction for free-form text; multimodal extraction for
  screenshots/posters/flyers with fan-out to multiple proposals
- Proposal queue + web review UI (filter, bulk approve/reject, inspect
  original message and images, compare extracted vs enriched)
- Venue canonicalization pipeline: alias table → LLM normalization →
  Places API → confidence-scored auto-accept or review
- Clarification requests back to the submitter for ambiguous items
- Auto-expiry of stale pending proposals

**Exit criteria:** a forwarded poster becomes reviewed, published,
calendar-synced events without anyone typing structured data.

## Phase 4 — identity, enrichment & notifications

**Goal: the platform knows its people and communicates proactively.**

- User accounts with linked messaging identities; passwordless magic-link/
  one-time-code auth over Signal; roles and trust levels
- Attribution (forwarder vs original author vs source) and contributor
  recognition (history, accepted submissions, leaderboard) with
  privacy-by-default
- Enrichment pipeline with confidence + caching: ticket links, artist
  pages, presales, representative videos, bios, age restrictions, GA/seated
- Artist pages on the site backed by cached enrichment
- Digests and notifications (Signal first): new events, presales opening,
  on-sale today, weekly what's-new

**Exit criteria:** a weekly digest goes out automatically and event pages
are visibly richer than a calendar entry.

## Phase 5 — discovery, music pipeline & conversational access

**Goal: the platform finds things on its own and you can talk to it.**

- Automated discovery pipeline (venue websites/calendars, newsletters,
  socials) with per-source provenance, always landing in editorial review
- Music recommendation pipeline: shared links become searchable records
  that attach to future events by the same artists
- Python retrieval sidecar: embeddings, hybrid retrieval, reranking
- Chat interfaces: built-in website chat (authenticated), then Discord/
  Telegram — lookup, recommendations, contribution, admin where authorized
- Abuse controls: quotas, rate limiting, moderation
- Optional: bot promoted into the Signal group with a strict
  mention-only policy

**Exit criteria:** events appear in review that no human submitted, and
"what's happening this weekend?" gets a good conversational answer.

## Sequencing notes

- Phases 1–2 involve no AI at all — they're pure engineering and de-risk
  everything after.
- Phase 3 is the biggest experience jump and depends only on Phase 1.
- Enrichment (4) is deliberately after review (3): enrichment is most
  valuable when a human is looking at the proposal.
- The retrieval sidecar is deferred to the last phase on purpose — venue
  resolution explicitly does not need RAG (Places API is authoritative).
