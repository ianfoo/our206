# our206 Direction Docs

This directory codifies the product and engineering direction for evolving
our206 from a Google Sheets + Apps Script + shared-calendar prototype into a
community event platform.

It synthesizes three planning briefs (platform vision, venue resolution &
semantic retrieval research, and the product/domain brief) into a single
coherent direction, reconciled against the code that exists today in this
repo.

## Reading order

| Doc | What it covers |
|-----|----------------|
| [product-vision.md](product-vision.md) | What we're building and why; product surfaces; principles |
| [architecture.md](architecture.md) | System architecture, domain model, pipelines, venue resolution, sync |
| [roadmap.md](roadmap.md) | Phased plan from the current spreadsheet system to the full platform |
| [decisions.md](decisions.md) | Key decisions made, with rationale, plus open questions |

## Current state (baseline)

What exists in this repo today:

- `calendar-sync/` — Google Apps Script bound to a spreadsheet. The
  spreadsheet **is** the canonical datastore. The script provides:
  - shorthand-line ingestion (`5/23: SOME BAND @ Neumos ✅`) with parsing,
    artist de-shoutifying, venue normalization, and exact/fuzzy dedup
  - venue normalization via a `Venue Map` sheet plus baked-in fallback
    alias and address maps
  - deterministic UID-based one-way sync to Google Calendar (create,
    update, delete orphans), debounced on edit
  - daily maintenance (archive past events, sort, sync)
  - Google Form intake
- `website/` — a static page on GitHub Pages (our206.com) that embeds the
  Google Calendar in an iframe.

The prototype proved the editorial workflow. The direction docs describe how
its responsibilities move into a real application without losing the
low-friction workflow that made it work.
