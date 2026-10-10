<div align="center">

# gessosheet

**A million-cell spreadsheet that keeps recalculating while your browser is frozen solid.**

Built on [Gesso](https://github.com/kevinpbaker/gesso), a UI framework that lays out, paints and hit-tests the whole interface in a worker, and lets the page's own thread do nothing at all.

TypeScript · Canvas · three threads · zero DOM under the sheet · a frame budget that fails the build

[Sixty seconds](#sixty-seconds) · [What it does](#what-it-does-today) · [The numbers](#the-numbers) · [How](#how-it-works) · [Not yet](#what-it-is-not-yet) · [Under the hood](#under-the-hood)

</div>

---

Everyone has felt a browser spreadsheet die. Type a formula, watch the
cursor stop, wait. The DOM is the wrong place to put a hundred thousand
cells, and every web spreadsheet knows it, which is why the serious ones
draw on a canvas and still run the whole thing on the one thread the
browser also uses for everything else.

This one does not. The grid is painted by a **render worker**. The
formula engine runs in an **application worker**. The main thread
creates a canvas, forwards input, and is otherwise idle — so idle that
the page carries a button to prove it:

> **Block the main thread for 5 s.** The page freezes. Hover does
> nothing, the tab will not close, DevTools stops answering. Behind the
> freeze the application worker finishes recalculating 200,000
> dependent cells, and the render worker keeps laying out and drawing
> on its own clock. When the main thread comes back, the sheet is
> already up to date.

That is the claim, and the rest of this repository exists to make it
checkable by a stranger with a browser.

## Sixty seconds

```bash
pnpm install
pnpm dev
```

Open the address Vite prints. You are looking at a three-sheet workbook
of a quarter's orders: `VLOOKUP` across sheets, `SUMIF`, `INDEX/MATCH`,
`NETWORKDAYS`, a `=SUM(Revenue)` through a named range, a colour scale
down the revenue column, and a review column whose dropdown refuses a
word it does not know. Type into it. Everything you see is a formula
reading a formula.

Then add `/proof` to the url. The black strip along the top is the
instrument panel, and it is the only DOM on the page, on the main
thread on purpose — a claim that a thread is idle cannot be made from
inside it. Three things to try:

1. **Scroll.** A pulse in the strip beats with the main thread's own
   `requestAnimationFrame`. The frame readout beside it is the render
   worker's: median cost, worst gap, and how many cells were measured
   this frame. Scroll a screenful and the count is a screenful. Scroll
   ten thousand rows and the count is still a screenful.
2. **Recalculate 200,000 cells** (toolbar, Data menu, or F9) while you
   are still scrolling. The frame readout does not move.
3. **Block the main thread.** Watch the pulse stop and the sheet not.

`pnpm proof` does all three in headless Chrome with real wheel events
and real clicks, reads back the same frames the strip shows, and fails
the build if any of six budgets is missed. The one that matters most is
machine-independent: *a frame while 200,000 cells recalculate may cost
no more than 4 ms over a frame while idle.*

## What it does today

A spreadsheet somebody would keep a budget in, not a demo of a grid.

- **The engine.** A parser, a dependency graph and an incremental
  recalculator on their own thread. Over a hundred and fifty functions
  across logic, maths, statistics, text, lookup, dates, finance,
  arrays and conditional aggregates: `SUMIFS`, `XLOOKUP`, `INDEX/MATCH`, `TEXTJOIN`,
  `EOMONTH`, `STDEV`, `PERCENTILE`, `IFERROR`, and the `">10"` criteria
  grammar they share. Dates are serials with a format, as in Excel.
  Errors explain themselves: land on a `#REF!` and a line under the
  cell says which cell went where.
- **Editing.** Undo and redo. Cut, copy and paste as a rectangle,
  including to and from other applications as TSV. Fill down and
  right. Find and replace across the sheet. Go to a typed address.
- **The formula editor.** Autocomplete with the signature and the
  current argument highlighted. References coloured in the text and
  outlined on the grid as you type. Click or drag a range into a
  formula mid-typing. F4 cycles `A1 → $A$1 → A$1 → $A1`. Matching
  brackets light up.
- **Structure.** Insert and delete rows and columns, and every formula
  that pointed at them is rewritten — exactly the ones that did, by a
  spec that says `toBe` and not `toBeLessThan`. Sort a range. Hide,
  freeze, merge, autofit, filter.
- **Formatting.** Number, currency, percent and date formats. Bold,
  italic, alignment, fills, per-edge borders. Conditional formats,
  including colour scales, resolved for the visible window only, so a
  rule over a million cells costs the same scroll as no rule. Data
  validation with a dropdown when the rule is a list.
- **Charts.** Line, column, bar, stacked, area, pie and scatter, over
  the sheet and dragged where you want them, with the value axis in
  the cells' own format. A chart of fifty thousand rows is sent to the
  render worker at the resolution its width can draw, and says so in
  its corner.
- **Workbooks.** Many sheets, with tabs to add, rename, reorder,
  duplicate and colour. `Sheet2!A1` and `'Q3 Budget'!A1:B9` in the
  parser. Named ranges from the name box or `Insert ▸ Name`. A formula
  on one sheet reading 50,000 cells on another publishes nothing while
  that sheet is out of view.
- **Persistence.** The workbook is saved in the browser as you type
  and is there when you come back.
- **Files.** Open and save `.gsheet` through the file system, several
  documents at once, recent files, a file dropped on the window. CSV
  in and out. `.xlsx` in and out with its formulas, formats, merges,
  names, conditional formats, validations and charts; what a file has
  that this sheet does not is named in a sentence rather than dropped
  without a word.
- **Formulas that are functions.** `LET`, and `LAMBDA` kept under a
  name and called like any other function, with `MAP`, `REDUCE`,
  `SCAN`, `BYROW` and their kin to hand one to.
- **Scripts.** A small typed API — read a range, write it, format it,
  add a sheet — run from *Data ▸ Scripts* on a worker of its own, one
  undo step a run, behind a written security model
  ([`SCRIPTS.md`](SCRIPTS.md)). And functions written in JavaScript
  that a formula calls, `=TAX(B2, 0.2)`, run by an interpreter inside
  the evaluator with its own time and memory limits, off by default in
  a file somebody else wrote.
- **Scenarios.** Name a way the year might go and type its inputs
  differently: *Scenario ▾* beside the tabs shows the whole workbook
  as it would be, with what the scenario typed and what that moved
  tinted, and the base's value for the cell you are on. The example
  opens with an Optimistic and a Pessimistic one. *Side by side with*
  puts another version in a second pane that scrolls with the first,
  tinted where the two differ. A scenario is the workbook forked in the
  engine, so it recalculates only what its inputs reach.
- **How sure.** Write a guess as a guess — `=NORMAL(0.02, 0.01)`,
  `=UNIFORM(40, 60)`, `=TRIANGULAR(5%, 6.1%, 8%)` — and it is its
  likeliest value until *Data ▸ Run a simulation* draws it five
  thousand times. Every cell it reaches gets a histogram under its
  figure, and the cell you are on reads out its P10, P50 and P90. The
  example's forecast growth and loan rate are guesses.
- **Light and dark.** The sheet follows the system's appearance, and
  View ▸ Theme overrides it, remembered in this browser. A cell's
  fill is its author's colour in light; in dark, a pale fill — a grey
  header, a colour scale's pastels — is drawn as the dark shade of
  the same hue, and a strong one like a navy title bar as it is.
  Text with no colour of its own is written in whichever ink reads on
  the fill as drawn.
- **Keyboard.** Every menu, every dialog and the whole toolbar are
  reachable without a pointer, and the specs for the chrome contain no
  pointer event to prove it. Ctrl+/ opens the sheet of shortcuts.

## The numbers

Measured on one Linux machine under software rendering. What matters is
the shape: the two rows are the same row.

|                                             | median frame | worst  | cells measured |
| ------------------------------------------- | ------------ | ------ | -------------- |
| scrolling an idle sheet                     | 1.9 ms       | 7.4 ms | 604            |
| scrolling while 200,000 cells recalculate   | 2.0 ms       | 5.5 ms | 233            |

The million-cell sheet (10,000 × 100) scrolls at 60 fps on both axes
with every visible value arriving from the other worker, at a 9,000
px/s fling, with no blank cell in any frame. The band that keeps it
clean, and why it belongs on the fetch side rather than the mount side,
is in [`PHASE0.md`](PHASE0.md).

The suite is 2,643 specs, most of them running the engine headless in
node. A handful are budget specs that count: patches per
scroll, formulas rewritten per insert, cells published for a
cross-sheet reference. They assert the number, not an upper bound.

## How it works

| Thread            | Owns                                                                           |
| ----------------- | ------------------------------------------------------------------------------ |
| Main (the shell)  | the canvas, input forwarding, the editing proxy for IME, the clipboard, `/proof`'s strip |
| Application worker | the cell store, the parser, the dependency graph, recalculation, persistence   |
| Render worker     | the grid, the selection, the cell editor, the menus, everything painted        |

The render worker never learns about a cell that is not on screen. It
sends the application worker a viewport, and the application worker
publishes that window plus an overscan band, as already-formatted
display strings, as a map keyed by row and column — because Phase 0
measured a row-major array at eighty times the bytes over Gesso's
structural differ. Formats travel on a separate key from values so that
editing one cell does not walk the format map. A conditional format is
resolved at the window on publish and folded into the palette index
already on the wire.

The three-way split is what the block button demonstrates. The page
freezes because `requestAnimationFrame` exists on the main thread
alone, so the render worker's frames spread to a timer's cadence and
no input arrives. What does not happen is the part people expect: the
sheet does not stop computing, and it does not stop drawing.

The longer story, with the exit criterion for each of thirty-seven phases
and what running it in a real browser found that the suite could not,
is [`ROADMAP.md`](ROADMAP.md).

## What it is not yet

Charts and `.xlsx` were the next two phases the last time this was
written. They are above now, and scripting, scenarios and a Monte
Carlo simulation with them. What is still not here: pivot tables; collaborative editing; rich text within
a single cell, where a bold word inside a cell is a different text
model from formatting per cell; images in a cell or over the sheet;
flash fill. An `.xlsm`'s VBA macros are not run: a different language
with a large runtime, and running a stranger's is the security
problem everybody already knows.

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
```

`src/sheet` is the engine and imports nothing from the framework, which
`boundaries.spec.ts` enforces. `src/app` is the contract, the
application worker, the grid, the editor and the chrome. `src/shell` is
the proof strip.

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
reads is there.

The routes are declared in `src/app/Routes.tsx`, in the render worker,
because a route names a component and a component cannot cross a
`postMessage`; the shell's half is `history: { mode: 'path' }` in
`createApp` and nothing else. The strip is the one part that cannot
work that way — it is DOM on the main thread, which is the entire
reason it is believable — so the shell reads the url for itself too,
through `src/route.ts`, the one module both threads share.

### The three files

| File            | What it is                                               |
| --------------- | -------------------------------------------------------- |
| `src/main.ts`   | The main thread: create the app and mount it into `#app` |
| `src/worker.ts` | The render worker: name the root component               |
| `src/App.tsx`   | The screen                                               |

There are three rather than two because a component cannot cross
`postMessage`, so the root has to be named on the worker's side of the
barrier.

`main.ts` names no worker, and that is `gesso-vite-plugin` in
`vite.config.ts`. It finds `worker.ts` beside `main.ts` and writes the
construction, which has to be written out as a literal because a bundler
emits a chunk for a worker it can see constructed and cannot see through
a variable holding the URL. The plugin also, while the dev server is
running:

- replaces the screen when you save `App.tsx`, instead of reloading the
  page, and puts the keyboard focus back where it was;
- draws what the render worker threw over the application that was
  running when it threw it, source-mapped, with the node named as a path
  through your components rather than as an id;
- says so when a save is about to reload the page, which happens when a
  module is reachable from the main thread as well as from the worker.

If you would rather say it out loud, write it and the plugin leaves the
construction alone:

```ts
createApp({
  renderWorker: () => new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' })
});
```

```json
"jsx": "react-jsx",
"jsxImportSource": "gesso-framework"
```

Those two lines in `tsconfig.json` are the whole of what makes `<row>`
and `<text>` work. JSX here is a spelling rather than a runtime: it
compiles onto the element factories and produces the identical tree, so
`Row({ gap: 8 }, Text({ text: 'Ready' }))` is the same thing written the
other way.

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

### Reporting errors from a production build

The overlay is a development tool and a production build carries no
reference to it. What a shipped application wants instead is its own
`onError`, which takes the same three arguments:

```ts
createApp({
  onError: (message, stack, source) => {
    reportToYourService({ message, stack, source });
  }
});
```

`source` is one of `message`, `uncaught`, `renderer`, `listener` or
`channel`, and it says what the failure cost: a `listener` error means
one handler did not run, a `renderer` error means the surface stopped
being updated, and a `channel` error means the data behind an intact
view has stopped arriving.

### Where to go next

- `App.tsx` is commented with what each part of it is doing.
- `gesso-components` has the controls: inputs, overlays, structure,
  data and media. `Switch` in `App.tsx` is one of them.
- Every prop takes a value or an Observable of that value. That is the
  whole binding model, and it is why the component body runs once.

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
