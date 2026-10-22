<div align="center">

<img src="public/icon-192.png" width="96" height="96" alt="" />

# Tessera

**The spreadsheet that keeps working when your browser doesn't.**

A full spreadsheet in a browser tab — formulas, charts, `.xlsx`,
scenarios and simulations — that never makes you wait on it.

### [Open Tessera →](https://gesso-sheets.vercel.app)

No account. No install. Your workbooks stay in your browser.

[Why it's different](#every-browser-spreadsheet-freezes-this-one-doesnt) · [What's inside](#everything-you-expect) · [What others don't do](#and-what-other-spreadsheets-dont) · [The proof](#proof-not-promises) · [Run it yourself](#run-it-yourself) · [Under the hood](#under-the-hood)

</div>

---

## Every browser spreadsheet freezes. This one doesn't.

You know the moment. You paste a column, or type a formula that a
thousand cells depend on, and the tab stops. The cursor sticks. The
scrollbar won't move. You wait for the spreadsheet to finish thinking
before you're allowed to keep working.

That happens because almost every web spreadsheet does everything —
the formulas, the drawing, your typing — on the one thread the browser
also uses for everything else. When the maths gets heavy, the whole
page waits.

**Tessera splits the work across three threads.** One recalculates
your formulas. One draws the sheet. The page itself does almost
nothing but pass along your keystrokes. Heavy recalculation never
stands between you and the next scroll, click or keystroke.

## Try to break it

Open **[gesso-sheets.vercel.app/proof](https://gesso-sheets.vercel.app/proof)**.
It's the same spreadsheet with an instrument panel across the top, and
three things to try:

1. **Scroll.** The panel shows the cost of every frame. Scroll one
   screen or ten thousand rows — the work per frame is the same,
   because Tessera only ever handles the cells you can see.
2. **Recalculate 200,000 cells** (toolbar, *Data* menu, or F9), and
   keep scrolling while it runs. The frame times don't move.
3. **Block the main thread for 5 seconds.** This deliberately freezes
   the page: hover dies, the tab won't close, DevTools stops answering.
   Behind the freeze, Tessera finishes recalculating all 200,000 cells
   and keeps drawing. When the page comes back, the sheet is already
   up to date.

A spreadsheet that does its work on the page can't survive that
button, because there the page *is* the spreadsheet.

## Everything you expect

Tessera is a spreadsheet you could keep a budget in, not a demo of a
fast grid.

- **Over 150 functions.** `SUMIFS`, `XLOOKUP`, `INDEX/MATCH`,
  `TEXTJOIN`, `EOMONTH`, `STDEV`, `PERCENTILE`, `IFERROR` and the rest
  of the logic, maths, statistics, text, lookup, date, finance and
  array functions you reach for. Dates work the way Excel's do.
- **Errors that explain themselves.** Land on a `#REF!` and a line
  under the cell tells you which cell went where.
- **A formula bar that helps.** Autocomplete with each function's
  signature, references coloured in the text and outlined on the grid,
  click-to-insert ranges, F4 to cycle `$A$1`, matching brackets, and a
  searchable list of every function under *Help ▸ Functions*.
- **Editing that just works.** Undo and redo, copy and paste to and
  from other apps, fill down and right, find and replace, sort, filter,
  hide, freeze, merge and autofit. Insert or delete a row and every
  formula that pointed at it is rewritten — exactly the ones that did.
- **Formatting the way you learned it.** Number, currency, percent and
  date formats, fonts, bold and italic, alignment, fills, borders,
  conditional formats and colour scales, and dropdowns that only accept
  the values you allow.
- **Charts.** Line, column, bar, stacked, area, pie and scatter, dragged
  wherever you want them. A chart over fifty thousand rows is drawn at
  the detail its width can show, so it stays as quick as a small one.
- **Workbooks.** As many sheets as you like, with tabs to rename,
  reorder, duplicate and colour, cross-sheet references, and named
  ranges.
- **A head start.** *File ▸ New* opens a blank sheet, a template — a
  monthly budget, an invoice, a project tracker, a weekly timesheet —
  or a worked example.
- **Light and dark.** Follows your system, or pick one in *View ▸
  Theme*. Your colours are adjusted so a pale header stays readable in
  dark mode.
- **Fully keyboard-driven.** Every menu, dialog and toolbar button
  works without a mouse. Ctrl+/ shows the shortcuts.

## And what other spreadsheets don't

- **Scenarios.** Name a way the year might go — *Optimistic*,
  *Pessimistic* — and change its inputs. *Scenario ▾* shows the whole
  workbook as that scenario would have it, with every cell it moved
  tinted. Put two versions side by side and they scroll together,
  highlighted wherever they differ.
- **Uncertainty, built in.** Type a guess as a guess:
  `=NORMAL(0.02, 0.01)`, `=UNIFORM(40, 60)`,
  `=TRIANGULAR(5%, 6.1%, 8%)`. *Data ▸ Run a simulation* draws it five
  thousand times, and every cell it reaches gets a histogram and its
  P10, P50 and P90. Stop pretending your forecast is one number.
- **Your own functions.** `LET` and `LAMBDA`, with `MAP`, `REDUCE`,
  `SCAN` and `BYROW` — or write a function in JavaScript and call it
  from a formula, `=TAX(B2, 0.2)`, run in a sandbox with its own time
  and memory limits. Functions in a file somebody else sent you are off
  until you turn them on.
- **Scripts.** A small, typed API to read, write and format ranges and
  add sheets, run from *Data ▸ Scripts* on a worker of its own, and
  undone in one step. The security model is written down in
  [`SCRIPTS.md`](SCRIPTS.md).
- **Version history.** *File ▸ Version history* puts a workbook back
  the way it was when you opened it, ten minutes into an edit, or when
  you last saved it.

## Your work stays yours

- **Saved as you type**, in your own browser. The title bar always says
  where your work is — *Saved in this browser*, *Saved to Q3.gsheet* —
  and turns red if it isn't being kept.
- **No account, no server.** Tessera is a static page. Your workbooks
  never leave your machine unless you save them somewhere.
- **Real files.** Open and save `.gsheet` files, and keep a file
  updated on every change. CSV in and out. **`.xlsx` in and out** with
  formulas, formats, merges, names, conditional formats, validation and
  charts — and if a file uses something Tessera doesn't have yet, it
  tells you in a sentence instead of dropping it quietly.
- **Print and PDF**, from *File ▸ Print* and *Export as PDF*.

## Proof, not promises

The speed claim is tested on every build, not just demonstrated.
`pnpm proof` drives the real app in headless Chrome with real wheel
events and real clicks, and **fails the build** if a frame budget is
missed. The most important budget doesn't depend on the machine: *a
frame while 200,000 cells recalculate may cost no more than 4 ms over
a frame while idle.*

|                                           | median frame | worst  | cells measured |
| ----------------------------------------- | ------------ | ------ | -------------- |
| scrolling an idle sheet                   | 1.9 ms       | 7.4 ms | 604            |
| scrolling while 200,000 cells recalculate | 2.0 ms       | 5.5 ms | 233            |

Measured on one Linux machine under software rendering. The point is
the shape: the two rows are the same row.

A million-cell sheet (10,000 × 100) scrolls at 60 fps in both
directions at a 9,000 px/s fling, with no blank cell in any frame. The
engineering behind that is in [`PHASE0.md`](PHASE0.md).

More than 2,700 specs back the rest, and the ones that count work —
cells sent per scroll, formulas rewritten per insert — assert the exact
number, not an upper bound.

## How it works

| Thread             | Does                                                                         |
| ------------------ | ---------------------------------------------------------------------------- |
| Main (the page)    | the canvas, forwarding input, typing in other scripts (IME), the clipboard  |
| Application worker | the cells, the formula parser, the dependency graph, recalculation, saving   |
| Render worker      | the grid, the selection, the editor, the menus — everything you see          |

The render worker never hears about a cell that isn't on screen. It
tells the application worker where you're looking, and gets back just
that window, already formatted. That's why the cost of a frame doesn't
grow with the size of the sheet, and why a conditional format over a
million cells scrolls like no rule at all.

Tessera is built on [Gesso](https://github.com/kevinpbaker/gesso), a UI
framework that lays out, paints and hit-tests a whole interface in a
worker. There is no DOM under the sheet.

The full build log — forty-two phases, each with its exit criterion
and what testing it in a real browser found — is in
[`ROADMAP.md`](ROADMAP.md).

## What's not here yet

Pivot tables, real-time collaboration, rich text within a single cell,
images, and flash fill. VBA macros in `.xlsm` files are not run, on
purpose: running a stranger's macros is the security problem everyone
already knows.

## Run it yourself

```bash
pnpm install
pnpm dev
```

Open the address Vite prints. A first visit opens *Northwind Trading*,
a small business's year — a dashboard, a ledger of orders, a forecast
and charts — with `VLOOKUP` across sheets, `SUMIF`, `INDEX/MATCH`,
`NETWORKDAYS`, `=SUM(Revenue)` through a named range, a colour scale
down the revenue column, and a dropdown that refuses a word it doesn't
know. Every number on it is a formula reading a formula. Add `/proof`
to the url for the instrument panel.

---

## Under the hood

Everything below is for someone working on the code.

```bash
pnpm dev          # the spreadsheet, with hot replacement of the screen
pnpm build        # a production bundle
pnpm preview      # serve it
pnpm typecheck    # tsc, no emit
pnpm test         # vitest, headless
pnpm proof        # build, serve, drive in headless Chrome, check the budgets
pnpm icons        # re-lift the Heroicons paths this app uses
pnpm app-icon     # redraw Tessera's own icon and its PNGs
```

`src/sheet` is the engine and imports nothing from the framework, which
`src/sheet/boundaries.spec.ts` enforces. `src/app` is the contract, the
application worker, the grid, the editor and the chrome. `src/script`
is the scripting API and the worker scripts run on. The proof strip is
`proofPanel` from `gesso-devtools`.

Called Gessosheet until October 2026, and still `gesso-sheets` as a
repository and in the roadmap. What the browser has already saved keeps
the old name too — the `gessosheet:` settings and `gessosheet.json` in
the origin's private files — so the rename did not lose anyone's work.

### The two pages

`/` is the spreadsheet. That is the whole of it: a grid, a menu bar, a
toolbar and a formula bar, and nothing on the screen that is about the
screen.

`/proof` is the same spreadsheet with the instruments attached — the
black strip along the top, and the one command the strip exists to
measure:

- a pulse driven by the main thread's own `requestAnimationFrame`, so
  it is alive exactly when that thread is;
- a button that holds the main thread in a busy loop for five seconds,
  which freezes the page and not the sheet;
- the layout heatmap, which washes every node the engine measured this
  frame and explains whatever is under the pointer;
- the render worker's frame rate, its worst gap, and how many nodes
  were re-measured;
- **Recalculate 200,000 cells** — on the toolbar, in the Data menu and
  on F9, and on this page only.

That list is why there are two urls. Nobody opening a spreadsheet
wants a black bar of instrumentation across the top of it, and a menu
item that exists to be photographed belongs on the page that
photographs it. `pnpm proof` drives `/proof`, because everything it
reads is there. The strip is the only DOM on the page, on the main
thread on purpose: a claim that a thread is idle cannot be made from
inside it.

The routes are declared in `src/app/Routes.tsx`, in the render worker,
because a route names a component and a component cannot cross a
`postMessage`; the shell's half is `history: { mode: 'path' }` in
`createApp` and nothing else. The strip is the one part that cannot
work that way — it is DOM on the main thread, which is the entire
reason it is believable — so the shell reads the url for itself too,
through `src/route.ts`, the one module both threads share.

### The three threads in code

| File                | What it is                                                         |
| ------------------- | ------------------------------------------------------------------ |
| `src/main.ts`       | The main thread: create the app, mount it, attach `/proof`'s strip |
| `src/worker.ts`     | The render worker: name the root component and the routes          |
| `src/AppWorker.ts`  | The application worker: the sheet and everything it computes       |

`main.ts` names no worker, and that is `gesso-vite-plugin` in
`vite.config.ts`. It finds `worker.ts` and `AppWorker.ts` beside
`main.ts` and writes the `new Worker(new URL(...))` construction for
each, which has to be a literal because a bundler emits a chunk only for
a worker it can see constructed. While the dev server runs, the plugin
also replaces the screen on save instead of reloading the page, and
draws what the render worker threw, source-mapped, over the
application.

The render worker publishes a viewport and the application worker
answers with that window plus an overscan band, as already-formatted
display strings in a map keyed by row and column — because Phase 0
measured a row-major array at eighty times the bytes over Gesso's
structural differ. Formats travel on a separate key from values so that
editing one cell does not walk the format map. A conditional format is
resolved at the window on publish and folded into the palette index
already on the wire.

### Icons

[Heroicons](https://heroicons.com/) is this project's icon set — the
24×24 outline half of it. It is MIT licensed, and the licence travels
with the path data in `src/app/heroicons.ts`.

Nothing in that package is imported at runtime. There is no DOM under
this interface, so an `<svg>` has nowhere to go; what Gesso's `Icon`
draws is the path itself, and `pnpm icons` lifts the paths this
application uses out of the package into `src/app/heroicons.ts`:

```bash
pnpm add -D heroicons@latest && pnpm icons
```

`scripts/icons.ts` holds the list of glyphs — add a line there and run
it again. `src/app/icons.ts` is the layer above, where a glyph gets an
application's name for it (`undo`, not `arrowUturnLeft`), and it is the
file to read to see which picture each button is wearing. A spec checks
the lifted data against the installed package, so an upgrade that skips
`pnpm icons` fails the build rather than quietly drawing last year's
glyph.

Tessera's own icon is drawn by `scripts/app-icon.ts`, which writes
`public/favicon.svg` and renders the PNGs with headless Chrome. On
macOS, point `CHROME_BIN` at Chrome's binary first.

### Deploying

The sheet is a static build — `dist/` is a page, its icons and
manifest, and a folder of hashed assets with one wasm — so
hosting it is hosting a directory. `vercel.json` is the whole of the
configuration: the Vite preset, the build, and one rewrite that sends
every path to `index.html`, because `/proof` is a route the router
knows and the filesystem does not. Hashed assets are sent with a
year's `immutable`; `index.html` is not, so a deploy is visible on the
next load.

Connecting it to GitHub is done once, in the dashboard:

1. **Add New… → Project** at [vercel.com/new](https://vercel.com/new),
   and import `kevinpbaker/gesso-sheets`.
2. Leave every build setting alone. `vercel.json` sets them, and what
   is written in the file wins over what is typed in the form.
3. **Deploy.**

After that every push to `main` deploys to production and every push
to any other branch, and every pull request, gets a preview url of its
own. Nothing has to be run by hand and no token has to be stored: the
GitHub App does the triggering, which is why there is no workflow file
in `.github/` here.

The install is a plain `pnpm install --frozen-lockfile`, with Gesso's
six packages coming from npm like everything else. `pnpm-workspace.yaml`
exempts those six from pnpm's one-day wait on new releases, so a Gesso
release can be deployed the hour it is published.
