# GUIDE.md — Event Launch RSVP

## What this project is

A single-page RSVP web app for an event launch, written for invited guests and
the host. Guests reply yes/no, say how many people (pax) are coming, pick a
slot and number of days, and name who invited them. The host sees every reply
and a live headcount on the same page.

**Live URL:** https://jamielimjw-afk.github.io/Event-launch-App/
**Repo:** https://github.com/jamielimjw-afk/Event-launch-App
**Local copy:** `C:\Users\Acer\Documents\Event launch ai\index.html`

## Technical shape

- **One file.** Everything lives in `index.html` — HTML, CSS, and JavaScript.
  No backend, no database, no login, no payments, no framework, no build step.
- **Runs by opening the file in any browser**, or via GitHub Pages (enabled,
  deploying from `main`, branch root `/`).
- **No persistence.** State lives only in a JavaScript array (`guests`) and is
  re-rendered on every change. Data resets when the page reloads. This is
  intentional per the requirements — nothing is saved between visits.
- Vanilla JS only. No external fonts, CDNs, or assets; it works fully offline.

## Page structure (all 3 sections stacked on ONE page — no tabs, no routing)

1. **RSVP Form** (`#rsvpForm`) — name, Yes/No toggle buttons (`#btnYes`,
   `#btnNo`), total pax (`#pax`), slot (`#slot` select: Morning
   10am–12pm / Afternoon 2pm–4pm / Evening 7pm–9pm), number of days (`#days`
   select: 1–3), invited by (`#invitedBy`). Submit adds the reply to the list.
   Name + a yes/no choice are required; a red error message (`#formError`)
   shows otherwise. "No" replies store pax/slot/days as "—".
2. **Guest List** (`#guestList`) — newest first. Each entry shows name, a YES
   (green) or NO (red) badge, pax, slot, days, and invited by.
3. **Headcount Summary** — stat boxes: `#totalConfirmed` (sum of pax across
   yes replies), `#totalParties` (count of yes replies), `#totalDeclined`
   (count of no replies); plus a per-slot pax breakdown table
   (`#slotBreakdown`).

All rendering goes through one function: `render()`. It rebuilds the guest
list and summary from the `guests` array. Submit handler calls `render()` at
the end — that is what makes the summary "live".

## Seed data

The app never opens empty: `guests` starts with 3 made-up Malaysian sample
entries (Farid Amran — yes/4 pax/morning/2 days; Siti Aminah Bakar — no;
Rajesh Kumar — yes/2 pax/evening/1 day). All sample data is fictional.

## Current status (as of 2026-11-17)

- ✅ Working: all 3 sections, yes/no RSVP, pax/slot/days capture, live
  summary, seed entries, form validation.
- ✅ `index.html` committed on `main` (commit `db4e723`) and pushed to GitHub.
- ✅ GitHub Pages enabled, deploying from `main` at branch root; site
  verified serving HTTP 200 at the live URL above.
- ✅ Git 2.56.0 installed on this machine via winget (was not present before).
- ⚠️ Credentials: the GitHub push token is stored by Git Credential Manager
  (Windows credential store), invoked as
  `git-credential-manager get` from `C:\Program Files\Git\ucrt64\bin\`.
  Note: GCM requires `git.exe` on PATH (use `C:\Program Files\Git\bin`)
  and its subcommand is `get`, not `fill`.
- ⚠️ No `.gitignore`, no tests, no CI. Nothing else has been added on purpose
  — the spec calls for exactly one self-contained file.

## Known limitations / possible next steps

- No editing or deleting of submitted replies (spec didn't ask for it).
- Number of days is a fixed 1–3 select; slots are hardcoded in both the form
  and the summary breakdown (`slotNames` array in the script — keep these in
  sync if slots change).
- Pax is capped at 20 by the input's `max` attribute (client-side only).
- If you change `index.html`, push to `main` and Pages redeploys
  automatically; if a redeploy ever 404s, a build can be forced via
  `POST /repos/{owner}/{repo}/pages/builds` on the GitHub API.
