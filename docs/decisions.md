# Decisions & Open Questions

The three source briefs were produced in separate threads and mostly agree.
Where they diverge, this file records the reconciliation. Format is
ADR-lite: decision, rationale, status.

## D1 — Implementation language: Go (single service)

**Status: recommended, awaiting confirmation.**

The product brief intentionally leaves the language open (Go, Python, or
other); the venue/retrieval brief assumes a Go application with a Python
sidecar. Recommendation: **Go** for the application.

- Single static binary fits the "low operational overhead, straightforward
  deployment, portability" requirements
- Strong fit for long-running bots, background sync workers, and an HTTP
  API in one process
- The Python ecosystem advantage (ML/retrieval tooling) is confined to the
  Phase 5 sidecar, which the venue brief already assigns to Python

If the implementation team prefers Python end-to-end, the architecture
holds — only this decision changes.

## D2 — Database: SQLite (not Postgres/pgvector)

**Status: decided for Phases 1–4.**

The product brief says SQLite is expected to be sufficient indefinitely;
the retrieval brief discusses pgvector. Reconciliation: canonical data
lives in SQLite. Vector storage is a Phase 5 concern and belongs to the
retrieval sidecar (its own store, or sqlite-vec if we want one file). We do
not adopt Postgres just to get pgvector; revisit only if Phase 5 workloads
demand it.

## D3 — Venue resolution: Places API, not RAG

**Status: decided.**

LLM normalizes messy input ("Croc" → "The Crocodile Seattle music venue");
a Places/Geocoding provider is authoritative for name, address,
coordinates, and stable place ID. The LLM is never the source of factual
address data. Store the place ID and refresh ~annually to catch moves,
closures, renames. General web search / embeddings / reranking are for
Phase 5 semantic features, not venue resolution.

**Open:** which Places provider (Google Places vs alternatives) — pick on
pricing and place-ID stability when Phase 3 starts.

## D4 — Google Calendar is a one-way projection

**Status: decided (both briefs agree; matches current code).**

Application → Google Calendar only. Direct calendar edits are not
authoritative and will be overwritten by reconciliation. The existing UID
scheme (date + normalized artist + normalized venue) carries over; sync
state moves from hidden description markers into a `CalendarSync` table.

## D5 — Signal: DM-forwarding first, group later, mention-only

**Status: decided.**

The bot does not join the main Signal group initially (privacy, trust,
minimal permissions). Members forward messages via DM. The architecture
keeps future in-group promotion cheap, under a strict policy: the bot
ignores every group message except explicit mentions. When someone tags the
bot on another member's message, attribution goes to the original author.

## D6 — Proposals subsume import dispositions

**Status: decided (reconciliation).**

One brief models review as EventProposals with lifecycle states
(pending/approved/rejected/expired/superseded); the other models import
items with dispositions (appended/exact duplicate/fuzzy duplicate/ignored/
needs review/error). These merge: every parsed submission item becomes a
proposal; disposition-like outcomes are how a proposal resolves. Trusted
channels (e.g. editor shorthand paste) may auto-approve unambiguous,
non-duplicate proposals; everything ambiguous waits for review. Fuzzy
duplicates are never silently merged.

## D7 — Preserve prototype behavior as the compatibility bar

**Status: decided.**

Phase 1 ports, with tests, the exact behaviors contributors rely on today:
the shorthand grammar, de-shoutifying, the ✅/! → score (max 4, displayed
🔥) conversion, alias-based venue normalization, exact/fuzzy dedup keys,
and UID-stable calendar sync. The existing `Venue Map` sheet plus the
baked-in maps in `calendar-sync/config.js` seed the venue and alias tables.

## Open questions

1. **Hosting/runtime target** — the app needs a persistent process (bots,
   jobs), so GitHub Pages no longer suffices for the whole product. Small
   VPS vs Fly.io vs home server; affects Phase 1 setup only, not
   architecture. The static site can stay on Pages until Phase 2 needs
   server rendering (or the site stays static and reads a published JSON
   feed).
2. **Signal transport** — signal-cli / signald / libsignal wrapper; needs a
   dedicated number. Evaluate at Phase 3 start.
3. **LLM provider & multimodal model** — needed at Phase 3 for
   classification/extraction; choose then, keep behind an interface.
4. **Website rendering strategy** — static generation from the canonical
   store (keeps Pages hosting) vs server-rendered from the API. Decide at
   Phase 2; the read API exists either way.
5. **Web framework / frontend stack** — deliberately open until Phase 2
   design work starts.
