# our206 Product Vision

## One sentence

our206 is a community-curated event intelligence platform for the Seattle
music scene: AI-assisted intake and enrichment, human editorial review, and
polished multi-surface publishing — of which the Google Calendar is just one
output.

## From prototype to platform

The Google Sheet + Apps Script prototype proved the editorial workflow, but
it is currently responsible for everything at once: canonical storage,
editorial UI, import queue, venue normalization, calendar sync, publication,
and logging. Those responsibilities move into a purpose-built application
while preserving the low-friction contributor workflow that made the
prototype successful.

Two identity shifts define the platform:

1. **The application owns the truth.** An internal database becomes the
   canonical event store. Everything else — website, Google Calendar, ICS
   feeds, JSON API, chat interfaces, digests — is a projection of it.
   Calendar sync is one-way (application → Google Calendar); edits made
   directly in Google Calendar are not authoritative.
2. **The website is the primary experience, not a calendar viewer.** Fast,
   responsive, visually polished, optimized for discovery: rich event cards,
   venue pages, artist pages, maps, ticket links, representative media.

## Core principles

- **Editorial first.** Automation exists to reduce repetitive work — "is
  this a duplicate?", "is this venue already canonical?", "is this worth
  surfacing?" — but final publication decisions are human. AI proposes;
  humans decide. Ambiguity is surfaced, never hidden.
- **Submissions never publish directly.** Every intake channel feeds the
  same pipeline: submission → proposal(s) → review → canonical event →
  synchronization. Each proposal carries extracted fields, confidence,
  enrichment sources, and the original message/images for inspection.
- **Historical preservation.** Future events are mutable; past events become
  historical records and are never modified automatically.
- **Privacy by default.** The community originates on Signal. Messaging
  identities are never publicly exposed unless a user opts in. The bot
  initially stays out of the group (members forward via DM); a future
  in-group bot ignores everything except explicit mentions.
- **Low-friction contribution.** The contributor shorthand
  (`3/20: Umphrey's McGee @ Showbox ✅✅✅✅`) keeps working forever. New
  channels (forwarded messages, screenshots, posters, links) reduce the
  effort further — the sender never has to type structured data; the system
  infers it.

## Product surfaces

### Website (primary)

- Discovery-oriented: upcoming shows, search, filtering, rich event detail
- Venue pages (address, map, upcoming events)
- Artist pages (bio, representative videos, interviews, related events,
  recently shared music)
- Ticket links, presale/on-sale info
- About/documentation section: contribution workflow, message formats,
  scoring, privacy, attribution, FAQ

### Chat interfaces

Conversational access with comparable capabilities across:

- built-in website chat (authenticated)
- Signal (first), Discord, Telegram (later)

Capabilities: event/venue/artist lookup, upcoming events, recommendations,
contribution workflows, and administrative actions where authorized. Because
chat incurs LLM cost, it is gated by authentication, trust levels, quotas,
rate limiting, and moderation.

### Intake channels

- Forwarded Signal messages (text, links, screenshots, posters) via DM
- Shorthand batch paste (the current `Incoming Raw` workflow)
- Web submission form (successor to the Google Form)
- CLI / API
- Automated discovery (venue sites, calendars, newsletters, socials) —
  always lands in editorial review, with provenance recorded

### Publication targets

- the website (primary)
- Google Calendar (one-way sync, as today)
- JSON API and ICS feed
- digests and notifications: newly added events, presales opening, on-sale
  today, weekly "what's new" — delivered to Signal first, later Discord,
  Telegram, email, RSS

## The intelligence layer

What makes this more than a CRUD calendar:

- **Multimodal extraction.** A forwarded poster or screenshot is inspected
  by a multimodal model to extract artist, tour, venue, city, date, time,
  on-sale/presale dates, restrictions. One image announcing ten tour dates
  fans out into multiple independent proposals.
- **Geographic intelligence.** From a multi-city tour poster, identify the
  local show, flag nearby regional shows (Portland, Vancouver), and ask the
  submitter whether to add them.
- **Venue canonicalization.** LLM normalizes messy names ("Croc" → "The
  Crocodile Seattle music venue"); a Places API is the authoritative source
  for address, coordinates, and a stable place ID. The LLM is never trusted
  for factual address data.
- **Enrichment with confidence.** Ticket links, artist pages, opening acts,
  age restrictions, seated/GA, accessibility, parking/transit — each with a
  confidence score and source. Enrichments are cached per artist/venue so
  future events reuse prior work.
- **Music knowledge base.** Shared YouTube/Bandcamp/SoundCloud/Spotify/
  Relisten links become searchable recommendation records instead of
  vanishing into chat history, and automatically attach to future events by
  the same artists.

## Community & recognition

- Internal user records link one person's Signal/Discord/Telegram
  identities to a single application account with roles (submitter, editor,
  administrator) and trust levels.
- Passwordless auth: a one-time challenge sent over the user's messaging
  platform proves identity and establishes a web session. Passkeys/OAuth
  may come later.
- Attribution is preserved: forwarding user vs. original author vs.
  ingestion source. If someone tags the bot on another person's message,
  credit goes to the original author.
- Contributors get recognition: contribution history, accepted submissions,
  discovery credits, leaderboard — always respecting privacy defaults.
