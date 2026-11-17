# AGENTS.md

Single-page RSVP app for an event launch. **The entire app is `index.html`** — HTML, CSS and vanilla JS in one file. No backend, no database, no persistence (`guests` is an in-memory array that resets on reload), no dependencies, no build step. Keep it that way unless asked: no CDNs, external fonts, assets, or frameworks.

There is no `package.json`, test suite, linter, CI, or codegen. Do not invent npm/test commands.

## Verifying changes

- Open `index.html` directly in a browser (or the in-app browser) and exercise it: submit a yes and a no, confirm the reply appears at the top of the guest list, and check the three stat tiles plus the per-slot table.
- `render()` is the single render path. It rebuilds the guest list and summary from `guests`; the submit handler mutates the array then calls `render()`. Any new state must be re-rendered there.
- Seed data: three fictional guests ship in the `guests` array, so a fresh load shows **6 confirmed pax, 2 yes parties, 1 no**.
- Rendering builds markup with `innerHTML` but writes guest values with `textContent`. Follow that pattern; never interpolate user input into HTML.

## Slot labels are duplicated (easy to break)

Slot strings appear in three places in `index.html`: the `<option>`s in `#slot`, the seed `guests`, and `slotNames` inside `render()`. The exact string — **including the `–` en dash**, e.g. `Morning (10am – 12pm)` — is the grouping key for `#slotBreakdown`. Change one copy and you must change all three, or confirmed pax silently falls into the `Other` row.

Other fixed limits: days is a 1–3 select; pax is capped at 20 only by the input's `max` attribute (no JS validation).

## Deploy

- Push to `origin/main` → GitHub Pages redeploys automatically from the branch root. Live at https://jamielimjw-afk.github.io/Event-launch-App/
- No Actions workflow or build is involved. There is no `.gitignore`; `build-final.ps1` and the `.pptx` are currently untracked.

## User manual deck

- `Event Launch RSVP - User Manual.pptx` is **generated**, not hand-edited. `build-final.ps1` opens PowerPoint via COM (Windows + installed PowerPoint required), creates a shell deck, then injects 8 hand-written DrawingML slides and reopens the result to verify.
- The script hardcodes the output directory at the top (`$dir = 'C:\Users\Acer\Documents\Event launch ai'`) and uses `$env:TEMP\_shell.pptx`. Update that path if the repo moves.
- Slide copy mirrors the UI: field names, slot labels, sample guests, and the palette. The palette is duplicated between the CSS `:root` vars in `index.html` and `$PRIM`/`$NO`/… in the script. When behavior or wording changes, edit the script and rebuild rather than editing the deck.
- It overwrites the `.pptx` in place and launches real PowerPoint; if verification fails it can leave a PowerPoint process running. Run it deliberately, not as part of a quick check.

## Docs

`GUIDE.md` holds the architecture/status notes and is worth updating when behavior or deployment status changes. It is prose, not a source of truth — trust `index.html`.
