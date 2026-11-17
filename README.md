# Event Launch RSVP

A single-page RSVP web app for an event launch. Guests reply yes/no, say how
many people are coming, pick a slot and number of days, and the host sees
every reply plus a live headcount on one page.

**Live site:** https://jamielimjw-afk.github.io/Event-launch-App/

## Run it

Open `index.html` in any browser — that's it. No install, no build, no
backend, no dependencies. The GitHub Pages deployment above serves the same
file.

## What's on the page

One screen, three stacked sections (no tabs, no navigation):

1. **RSVP Form** — guest fills in their name, picks Yes/No, total pax, a slot
   (Morning / Afternoon / Evening), number of days (1–3), and who invited
   them, then submits. The reply is added to the list below.
2. **Guest List** — every reply, newest first, showing name, a clear YES/NO
   badge, pax, slot, days, and invited by.
3. **Headcount Summary** — live totals: confirmed pax from yes replies,
   number of yes parties, number of no replies, plus a pax breakdown per
   slot. Updates the moment a reply is added.

The page opens with 3 made-up Malaysian sample entries so it's never empty.
Everything is in-memory only — data resets when the page reloads.

## Tech

- One self-contained file: `index.html` (HTML + CSS + JavaScript, vanilla).
- No framework, no npm, no backend, no login, no payments.
- All sample data is fictional.

See [GUIDE.md](GUIDE.md) for full documentation of the code structure and
project status.
