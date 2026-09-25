# Roadmap

A spreadsheet built on [Gesso](../gesso): a million cells, frozen panes,
and a formula engine that recalculates on a thread the scroll never
touches.

It exists to be evidence. Gesso's claim is that owning the whole stack —
component model, layout, input, rasterization — buys an application
something the DOM cannot, and that the main thread should be doing
almost nothing. A spreadsheet is the hardest honest test of that claim
available: everyone has felt a browser spreadsheet die, the failure is
legible without a profiler, and reviewers will try to break it
themselves rather than take a benchmark's word for it.

**Status:** Part One is done — eight phases, eight exit criteria met.
[Part Two](#part-two--a-spreadsheet-rather-than-a-demonstration) is the
plan for turning the proof into a spreadsheet somebody would keep a
budget in, and none of it has started. Phase
0's findings are in [`PHASE0.md`](PHASE0.md); the sheet model is in
`src/sheet`, the contract, the application worker, the grid and the
editor in `src/app`, the proof strip in `src/shell`, and `pnpm test` is
its 279 specs. `pnpm dev` is a spreadsheet you can type into, copy out
of and paste into, which remembers what you typed. `pnpm proof` is the
frame budget: it drives the built application in headless Chrome and
fails the build when scrolling stops being free.

---

## The thread split

The formula engine lives in the **application worker**. The render
worker never learns about a cell that is not on screen.

| Thread           | Owns                                                                        |
| ---------------- | --------------------------------------------------------------------------- |
| Shell (main)     | the canvas, input forwarding, the editing proxy for IME, the clipboard      |
| App worker       | the cell store, A1 references, the parser, the dependency graph, recalc     |
| Render worker    | the grid, selection, the cell editor, everything painted                    |

Concretely: the render worker sends `setViewport({ r0, r1, c0, c1 })` as
a command, and the app worker publishes only that window plus an
overscan band, as already-formatted display strings.

Not as arrays, though: Phase 0 measured the two shapes and a row-major
array costs eighty times the bytes of a map keyed by absolute row and
column, because `diffArray` finds no common prefix or suffix in a
window that scrolled and falls through to one patch per cell. See
[`PHASE0.md`](PHASE0.md) §4.

This is not an optimization, it is forced. Gesso's channel wire carries
plain data only — `requirePlainData` in `channel/plainData.ts` rejects
anything whose prototype is not `Object.prototype`, so no typed arrays —
and `diffProjection` in `channel/StorePatch.ts` walks the projection
structurally on every publish and emits one patch object per changed
leaf. A sheet as the projection would benchmark the differ. A window as
the projection is a few hundred patches a frame.

Values and formatting are separate channel keys, for the reason
`NotesContract.ts` gives for splitting `rows` from `open`: editing one
cell should not walk the format map.

**The risk this creates.** Scrolling runs at 60fps in the render worker,
but refilling the window is a round trip. Scroll past the overscan band
and cells are blank for a frame or two. The mitigations are a generous
band, keeping last-known values instead of clearing them, and a pending
cell treatment that is not a flash. Phase 0 exists to find out how bad
this is before anything is built on top of it.

Phase 0's answer: the band is 8 rows and 2 columns at a 9,000 px/s
fling, and it belongs on the **fetch** side. Widening the band the
render worker mounts buys the same lookahead, costs four times the
nodes, and past eight rows makes the frame slow enough to lose more
coverage than it gained. What the round trip cannot survive is a recalc
that blocks the publish: 30 ms of application thread leaves the sheet
blank in 89% of frames while the scroll itself stays at 60fps.

## What Gesso does not have yet

Three gaps, two found by reading the engine and one by scrolling the
Phase 0 spike by hand, all of which this project will have to close. Both are plausibly worth upstreaming rather than working
around, and that decision should be made once the shape is known, not
now.

**Two-axis virtualization.** `LazyGrid` in
`core/src/composition/UiLazyList.ts` virtualizes rows only; columns are
declared tracks shared with the header through `subgrid: 'columns'`.
Nothing windows columns. This is the main engine work, and the reason
Phase 0 is a spike rather than a start.

*Closed by Phase 0* as `UiVirtualSheet` / `LazySheet`, a peer of
`UiLazyList` rather than a second axis on `UiVirtualWindow`: a sheet is
told its geometry, so there is nothing to measure, and without
measurement the correction table, the scroll anchoring and the
per-frame extent walk all go away. It does not use `subgrid` either —
sharing tracks costs a full-content measure of every mounted cell on
every layout, and a sheet's column widths are given rather than
derived. The shape is now known, so the upstreaming decision this file
deferred can be taken.

**Two-axis scroll input.** *Found in Phase 0, and not one of the two
this file started with.* `UiWheelController` asked a scroll container
for a single axis, derived from its flex direction, so a container that
overflows both ways dropped every wheel delta on the axis it was not
classified as — a trackpad could not scroll the sheet sideways at all.
Fixed by asking each axis whether it has room. The scrollbar thumb's
grab target was also the six pixels it paints, which made a press page
the track instead of dragging. `UiTouchScroller` is still one-axis and
will need the same treatment before a tablet is in scope.

**Paste into a range.** `UiEditingController.paste` routes pasted text
into the focused editable. A spreadsheet needs to intercept a paste when
the grid owns a rectangular selection and no cell is being edited —
pasting a TSV block out of Excel is table stakes. Copy is fine: the
write path already exists through `ShellService.copyText`.

*Closed by Phase 5*, in two places rather than one. The engine offers
a `Paste` event to whatever holds focus when nothing editable does;
and the shell listens for the browser's paste on the canvas and on the
accessibility mirror, which is where focus actually is while an app is
being used. The first without the second changes nothing, and only a
browser says so.

One thing that is already there, and worth not rebuilding:
`EditableTextModel` in `core/src/editing/` has the caret, selection, IME
composition and coalescing undo, and is deliberately free of layout. It
drops into a cell editor as it is. The semantics roles a grid needs —
`grid`, `row`, `columnheader`, `rowheader`, `cell` — are all in
`UI_ROLES` already, so the accessibility story is available from the
first phase rather than retrofitted. (It is `cell`, not `gridcell`.)

---

## Phases

Each phase ends at something observable, not at "done". Where a phase
borrows an exit criterion from Gesso's own gates, it says so.

### Phase 0 — Burn down the two risks — **done**

Timeboxed, throwaway code. Spike two-axis windowing and the viewport
round trip together, with fake data and no formulas.

**Exit:** a 1,000,000-cell empty grid scrolls at 60fps on both axes with
values arriving from the app worker, and the overscan band that keeps it
clean is a known number. If this cannot be hit, the rest of this file
changes.

**Met.** 10,000 × 100 cells, both axes, at 3,000 and 9,000 px/s: a
16.6–16.7 ms frame gap on 4.2–5.2 ms of work, and no visible cell
without a value in any frame of any of the four runs. The band is 8
rows and 2 columns. The findings, including the three that change later
phases, are in [`PHASE0.md`](PHASE0.md); the harness that produced them
retired with the spike in Phase 3, and rebuilding it against the real
grid is owed to Phase 7.

### Phase 1 — The sheet model, headless — **done**

Pure TypeScript, no framework import, all under vitest in node: a sparse
cell store keyed by packed row/column, A1 references with `$` absolutes
and ranges, a tokenizer and parser, the dependency graph, topological
incremental recalc, cycle detection, and the error values (`#REF!`,
`#DIV/0!`, `#NAME?`, `#VALUE!`, `#CIRC!`). The function library starts
deliberately small: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`,
`ROUND`, `ABS`, `CONCAT`.

**Exit:** a recalc budget spec in the style of Gesso's
`LayoutEngine.budget.spec.ts` — editing one cell with N dependents
evaluates exactly the transitive closure, asserted as a count and not as
a timing. This is the spec that makes the headline claim checkable.

**Met.** `Sheet.budget.spec.ts`: editing one cell with 50,000
dependents evaluates 50,000 cells, `toBe` and not `toBeLessThan`, in
29 ms. Beside it, the counts that say the closure is the *right* one
rather than merely the right size — a diamond evaluates its shared
dependent once and not once per path, an edit nothing reads evaluates
nothing, an edit inside a `SUM` range evaluates the formula once and
not once per cell, and a rewritten formula stops being woken by the
reference it dropped.

Two things the phase added to its own description. Recalculation is a
**resumable queue**, not a call: `recalculate(budget)` does at most
that many cells and says whether more remain, because Phase 0 measured
that 30 ms of uninterrupted application thread leaves the sheet blank
in 89% of frames, and a recalc written as one long call cannot be cut
into slices afterwards. And the no-dependency rule is a spec rather
than a promise — `boundaries.spec.ts` fails the build, naming the file
and the import, if anything under `src/sheet` reaches for the
framework, RxJS or the browser.

### Phase 2 — Contract and application worker — **done**

`SheetContract.ts` with the windowed view and the commands
(`setViewport`, `setCell`, `setSelection`, `undo`, `redo`), and the
worker serving it.

**Exit:** a spec that drives commands and asserts patch counts per edit.
One keystroke in a cell with 50,000 dependents should emit patches
proportional to the visible window, not to the dependents. That single
assertion is the whole thesis.

Phase 1 leaves `Sheet` ready for it: `display(row, column)` is the
already-formatted string the window carries, `entries()` is what a
repository writes out, and `recalculate(budget)` is what lets the
worker publish a window between slices of a recalc.

Two things Phase 0 adds to this phase. The window is a map keyed by
absolute row and column, not arrays — the shape decides the patch count
more than the contents do. And a second spec is owed beside the
patch-count one: editing a large graph *while the viewport moves*, so
that a recalc which blocks the publish is caught here rather than felt
as a blank sheet.

**Met, and by the real machinery.** `SheetChannel.spec.ts` drives
`provide` over a recording port with the framework's own differ in
between, so the numbers it asserts are the ones that would cross a
`postMessage`. A column of 50,000 formulas, thirty rows of it in view,
one keystroke at the top: **50,000 cells evaluated, 30 patches sent** —
one per visible cell, `toHaveLength` and not a ceiling. The same edit
with the dependents scrolled off screen sends **nothing at all**, which
is the claim stated the other way round. A scroll of one row costs
four patches: the row that entered, the row that left, and the two
bounds.

The second spec is there too. The recalc is a pump: `recalculate` is
called a slice at a time and the thread handed back between slices, so
a `setViewport` arriving mid-recalc is answered on the spot, ahead of
the arithmetic — asserted with 20,000 cells still pending. The
scheduler is injected, so the slices are deterministic in a spec and a
task rather than a microtask in the worker; a microtask would never
let the command in, and the slicing would look like it was working
while the sheet stayed blank.

### Phase 3 — The grid surface — **done**

Frozen row and column headers through `position: 'sticky'`, which is
already conformance-tested against Chrome — though not against a
two-axis window, which is the piece Phase 0 did not cover. Selection
rectangle, active cell ring, column resize by drag, A/B/C and 1/2/3
headers.

Column resize turns every offset from a multiplication into a prefix
sum, which `UiVirtualSheet` does not do yet. And cell-element identity
is a performance contract here, not a detail: Phase 0's whole frame
budget went on allocating cell bindings until the elements were
memoized by the cell they hold.

**Exit:** specs querying the semantics tree as `grid` / `row` /
`columnheader` / `cell`, and `toHaveBox` assertions pinning the frozen
panes.

**Met, with one correction to the criterion itself.** The semantics
half is as written: the surface is a `grid`, its columns are
`columnheader`s named A and B, its rows are `rowheader`s numbered from
one as a person counts, and its cells carry the value rather than the
formula — `=1+2` is a cell named `3`.

The box half could not be written as stated. `toHaveBox` reports where
a node was *laid out*, and a sticky header is laid out at the top of
its content and stays laid out there however far the container
scrolls; four specs asserting `toHaveBox({ y: 0 })` after a scroll
passed against a header that was never sticky at all, and the bug was
found in a browser, by eye, afterwards. So `getVisibleBox` and
`toHaveVisibleBox` were added to `gesso-testing` and the frozen panes
are pinned with those, on each axis and on both at once.

Three pieces of engine work, all in the sibling checkout: columns of
different widths (a prefix sum, because a drag moves every offset after
it), the header row and the gutter as content the window accounts for,
and the testing gap above.

### Phase 4 — Editing — **done**

An in-cell editor over `EditableTextModel`, a formula bar bound to the
same buffer, and the commit semantics people have muscle memory for:
Enter commits and moves down, Tab commits and moves right, Esc reverts,
F2 enters edit mode, typing over a selected cell replaces it. IME
composition inside a cell.

**Exit:** a keyboard spec in the style of `gesso-components`'
`Keyboard.spec.ts` — navigate, type, commit, undo, entirely through the
semantics tree with no synthetic mouse.

**Met.** `Keyboard.spec.tsx`: twenty specs, no pointer event anywhere
in the file, every assertion read back from the semantics tree or from
what the application worker ended up holding. Tab from nothing reaches
the grid; arrows, Enter, Tab, Home, End and PageDown move the
selection; typing replaces a cell starting from the first character;
F2 opens it with the formula rather than the value; Enter commits and
moves down, Tab commits and moves right, Escape puts back what was
there; ctrl-Z takes it back and recalculates what it fed.

Two things worth keeping in the notes.

The selection **leads** on the render thread and the channel follows.
Read back across the barrier it lagged a frame or two, and every arrow
pressed inside that window moved from the same stale cell — hold an
arrow down and the selection travels one cell and stops. The draft is
render-side only and never crosses until it is committed, which is
what makes Escape possible at all.

The character that opens a cell arrives **once**. The key handler
seeds the draft and calls `preventDefault`, which is what stops the
browser delivering the same character again to the editor it has just
focused. Driving this from CDP with a separate `char` event produced
`==D3*2` and a `#VALUE!`, which is a harness artefact rather than a
bug — but the doubling is real enough to be worth the comment it now
carries.

### Phase 5 — Clipboard and fill — **done**

Copy a range to TSV. Paste a range, which needs the second gap closed.
Then the fill handle, with relative references adjusted as it extends.

Phase 4 leaves the seam it needs: `SheetEditing` owns the selection on
the render thread, so a copy knows its rectangle without asking, and
`EditableTextModel` is already handling the clipboard inside an open
cell. What is missing is the engine gap this file has carried from the
start — a paste that the grid takes when it owns a rectangle and no
cell is open.

**Exit:** a round trip inside the sheet, and a paste of real TSV copied
out of Excel or Google Sheets landing correctly.

**Met.** `Clipboard.spec.tsx` drives both halves through the real
paths: a copy reaches `ShellService`, which is the only thing with a
clipboard, and every paste arrives as the engine's `Paste` event
rather than as a command called by hand. The round trip carries
*formulas*, not the numbers they showed — copying a column of totals
and pasting it one column over pastes sums that compute from where
they landed, which is the difference between a spreadsheet's copy and
a screenshot of one. The Excel half is a block with CRLF endings, a
quoted cell holding a tab, and the trailing newline Excel adds, which
must not arrive as a row of blanks that wipes a row of the sheet.

The engine gap this file opened with is closed, and closing it took
two changes rather than one. `UiEditingController.paste` now offers
the text to whatever holds focus when nothing editable does — but that
alone did nothing in a browser, because a real paste never reached the
engine at all: with no field focused the editing proxy blurs, and the
browser fires `paste` at the canvas or, once the accessibility mirror
exists, at whichever of *its* elements has focus. Neither had a
listener. That second half was only findable by pasting into a real
browser; every spec passed without it.

### Phase 6 — Persistence — **done**

A repository writing through `FileSystemSyncAccessHandle` in the app
worker, mirroring `OpfsNotesRepository` in Gesso's notes example.

**Exit:** the notes example's own proof, borrowed — type a marker,
reload, it is still there.

**Met.** `OpfsSheetRepository`, `SheetFile.ts` and `Persistence.spec.ts`.
Inputs are written, not values: a formula is not derivable from its
answer, and reloading recalculates, which is the cheapest check that the
engine still agrees with itself.

### Phase 7 — The proof surface — **done**

This is what the application is *for*, so it is a phase and not a
nicety. Recalculate 200,000 dependent cells while scrolling, and show
that the scroll re-measured nothing. A "block the main thread for five
seconds" button. `engine.explain` wired to a cell inspector. The
re-measure heatmap visible.

**Exit:** each of those is a thing a stranger can do in a browser in
under a minute, and a frame budget in CI that fails when one of them
regresses.

**Met.** The strip along the top of the page is the only DOM in the
application, and it is on the main thread on purpose: a claim that the
main thread is idle cannot be made from inside the thread that is not.
It carries the block button, the heatmap toggle, and the frame readout;
the recalculate button is in the sheet's own toolbar, because that
command has to cross to the application worker. Hovering with the
heatmap on shows `engine.explain` for the cell under the pointer,
floated over the sheet so that reading it does not move what is being
read.

`pnpm proof` is the budget. It builds, serves, and drives the result in
headless Chrome with real wheel events and real clicks, and reads back
the same frames the strip is showing. Measured on this machine:

| | median frame | worst | most re-measured |
|---|---|---|---|
| scrolling an idle sheet | 1.9 ms | 7.4 ms | 604 |
| scrolling while 200,000 cells recalculate | 2.0 ms | 5.5 ms | 233 |

The two rows being the same row is the phase. The fourth budget is the
one that survives a different machine: a frame while recalculating may
not cost more than 4 ms over a frame while idle.

Three things the phase learned by being run rather than by being
written.

**The heatmap tells the truth and reads as the opposite.** Sweep the
sheet and every cell washes red, correctly — a virtualised row is
created and measured once when it comes into view, and a long sweep
replaces the screen. The claim was never that a scroll measures
nothing; it is that a scroll measures a screenful and stops, whatever
the sheet's height and whatever is recalculating behind it. So the
strip shows the peak next to the last frame, and a short scroll shows
the picture plainly: the rows that stayed are cold, only the incoming
band is red.

**The block button demonstrates something narrower and better than the
slogan.** The page freezes solid for five seconds — measured at 5,006
ms of a DevTools round trip going unanswered. The application worker
does not notice, and finishes a second 200,000-cell recalculation
during the freeze. The render worker does not stop either: it keeps
laying out and drawing on its own clock, 139 frames through the five
seconds. What it loses is the display's cadence and only that, because
`requestAnimationFrame` exists on the main thread alone, so its frames
spread out to a timer's interval. And no input reaches the sheet at all
while the shell is gone. "The sheet keeps scrolling at sixty" would
have been both untrue and weaker than that.

**The demo was saving itself into people's files.** Every spec passed.
Pressing recalculate once and reloading showed a sheet that had quietly
written a quarter of a million formulas to disk and recalculated all of
them on the way back up. The chain now lives one row past the end of
the sheet, where nothing can be scrolled to, selected or typed in, and
`snapshotOf` writes only what is inside the sheet. Sixth phase running
in which a real browser found something the suite could not.

---

## What was not in v1

Charts, pivot tables, multiple sheets, conditional formatting, and
merged cells were all held back, and all but the pivot tables come
back in Part Two below. The objection recorded against merging —
that it fights `subgrid: 'columns'` — is stale: Phase 3 stopped using
subgrid, for unrelated reasons, and the real cost turned out to be
somewhere else. See Phase 10.

---

# Part two — a spreadsheet rather than a demonstration

The eight phases above proved the claim. Nobody keeps a budget in a
proof.

Part two is the ordinary work: a top bar, formats, insert and delete,
a function library worth the name, several sheets, charts, and files
that came from somewhere else. None of it is novel and all of it is
required, and the interesting question is not how to build any one of
these — it is whether the thing Part One bought survives having them
built on top of it. A spreadsheet with a toolbar bound to the
selection, formats on every cell, conditional rules over a million
rows and a chart open is the sheet that usually stops being fast.

Somebody could stop after Phase 11 and have a spreadsheet they would
keep a budget in. The phases after it are what make them keep a
second one.

## The rule Part Two runs under

Every phase below inherits three exit criteria on top of its own, and
a phase that cannot meet them is a phase whose design is wrong rather
than a budget that needs raising:

1. **Patches are proportional to the viewport, never to the sheet.**
   Phase 2's spec is the template, and each phase that adds a view key
   owes one of its own.
2. **`pnpm proof` does not move.** The fourth budget — a frame while
   recalculating may not cost more than 4 ms over a frame while idle —
   holds with the whole chrome mounted, not just with a bare grid.
3. **`src/sheet` still imports nothing.** `boundaries.spec.ts` fails
   the build over a stray framework import, and three of the phases
   below are large enough to be tempted.

---

### Phase 8 — The top bar

A menu bar (File, Edit, Insert, Format, Data, View, Help), a toolbar
beneath it, the name box beside the formula bar, and a status bar
along the bottom carrying the selection's Sum, Average and Count —
the readout people check before they trust a column.

The bar is painted in the render worker, like everything else, and
that is the point of doing it rather than a reason to dread it. The
grid was one role repeated across a lot of cells; the bar is fifteen
roles with a traversal model — Alt opens the menu bar, arrows walk
across the menus and down their items, Escape closes, type-ahead
jumps, and an item shows its own shortcut. `menubar`, `menuitem`,
`menuitemcheckbox` and `menuitemradio` are in `UI_ROLES` already, so
the semantics are available; `Menu` in `gesso-components` is a popup
that knows nothing about a bar above it, so traversal *between* menus
is engine work, and probably upstream.

The trap is performance and it is Phase 0's trap wearing a hat. A
toolbar bound to the selection re-renders on every arrow key, so its
buttons must be memoized by what they show exactly as Phase 3
memoized cells by the cell they hold. A bar that reallocates its
bindings on every keystroke is the cell-binding bug again, in the one
part of the screen that is on top of everything.

Find and replace lands here too, on `FindBar`, along with the
distinction between clearing contents and clearing formats, and a
shortcut sheet that is generated from the keymap rather than written
beside it.

**Exit:** the whole bar driven from the keyboard through the semantics
tree with no pointer event anywhere in the spec — Phase 4's standard,
applied to the chrome. And `pnpm proof` with the bar mounted, plus a
fifth budget: the median frame while a menu is open over a scrolling
sheet.

### Phase 9 — The format axis

The format model in `src/sheet`: number formats (general, number with
places, currency, percent, scientific, date, time, text, and a custom
pattern), weight, italic, underline, size, text colour, fill,
horizontal and vertical alignment, wrap, indent, and borders.

The contract gains a seventh view key, `formats`, a window keyed
exactly as `window` is — carrying **an index into a palette, not a
record**. A column formatted as currency is one palette entry and one
small integer per visible cell; a record per cell would double the
bytes on the wire and give the differ a nested object to walk for
every cell in the window. The palette changes when somebody formats
something and the indices change when the viewport moves, and those
being different rates is the same argument that split `rows` from
`open`.

Formatting itself stays on the application worker. `display()`
already returns a formatted string and now consults the cell's format
to build it; the render worker must never learn a number format,
because a locale-aware formatter on the frame path is precisely the
work this architecture exists to keep off it. The consequence to
accept rather than engineer around: changing a format republishes the
window's *values* as well as its indices. That is bounded by the
viewport, which is the whole answer.

File format v2 — `formats` and the palette beside `cells` — and a
format change is a `Step` on the existing undo stack, not a second
one.

**Exit:** a patch-count spec. Formatting a column of 50,000 cells
emits one palette patch and one patch per visible cell,
`toHaveLength` and not a ceiling. And `pnpm proof` over a sheet where
every cell is formatted: the same budget, not a new one.

### Phase 10 — Rows, columns, and cells that span

Insert and delete rows and columns, hide and unhide, autofit, freeze
at an arbitrary cell rather than only at the headers, sort a range,
filter, and merge.

The hard half is references, and it is a different rewrite from the
one `Rewrite.ts` already does. A fill moves relative references and
pins absolute ones; an insert moves *both*, because `$A$1` is
absolute against a fill and not against a column that appeared to its
left. References into a deleted region become `#REF!`; a range that
spans the boundary grows or shrinks rather than moving. That one
distinction is what every implementation gets wrong exactly once.

The cost to watch is that an insert rewrites every formula below the
line and rebuilds their edges — on a 200,000-cell chain, that is the
whole graph in one call. It has to go through the same resumable
queue the recalc uses, or an insert is a five-second freeze on the
one thread in this application that is not allowed to have one.

Merging is here rather than in its own phase because merging is a
structural edit, and because the objection recorded against it is no
longer the real one. The real cost is that `UiVirtualSheet` mounts
the cells the window covers, and a merge whose anchor has scrolled
off the top must still paint into the window — so the mounted set has
to be widened by the merges intersecting it. That is engine work, and
it is the first thing in this project that makes what is mounted
depend on what the document says.

**Exit:** a rewrite-count spec in Phase 1's style — inserting a row
above a column of 50,000 formulas rewrites exactly the formulas that
reference it, `toBe` and not `toBeLessThan`. And an insert performed
mid-scroll that does not drop a frame, asserted by `pnpm proof`.

### Phase 11 — The library, and dates

Headless, in `src/sheet`, under vitest in node. `boundaries.spec.ts`
guards the whole phase.

Dates first, because a spreadsheet without them is a toy. A date is a
number with a format — Excel's serial from 1899-12-30 — and the only
new thing in the value model is that typing `2026-09-24` into a cell
sets a number *and* a format. That is the single point where the
engine and the format axis touch, and it is why this phase comes
after Phase 9 rather than before it.

Then roughly sixty functions in six families: logic (`AND`, `OR`,
`NOT`, `XOR`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`); maths (`SQRT`,
`POWER`, `MOD`, `INT`, `TRUNC`, `CEILING`, `FLOOR`, `SIGN`, `EXP`,
`LN`, `LOG`, `RAND`, `RANDBETWEEN`, `SUMPRODUCT`); statistics
(`MEDIAN`, `MODE`, `STDEV`, `VAR`, `COUNTA`, `COUNTBLANK`, `LARGE`,
`SMALL`, `RANK`, `PERCENTILE`); conditional aggregates (`SUMIF`,
`SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and the `">10"`
criteria grammar all of them share); text (`LEFT`, `RIGHT`, `MID`,
`LEN`, `FIND`, `SEARCH`, `TRIM`, `UPPER`, `LOWER`, `PROPER`,
`SUBSTITUTE`, `REPLACE`, `REPT`, `TEXTJOIN`, `VALUE`, and `TEXT`,
which is Phase 9's formatter called from inside a formula); lookup
(`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `OFFSET`,
`INDIRECT`, `CHOOSE`); and dates (`TODAY`, `NOW`, `DATE`, `DATEDIF`,
`YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `WEEKDAY`, `EDATE`,
`EOMONTH`, `NETWORKDAYS`).

Three of those are dangerous and are worth naming now rather than
discovering. `INDIRECT` and `OFFSET` compute their references while
being evaluated, so the graph cannot know their edges beforehand and
has to re-derive them after each evaluation; getting that wrong
produces a cell that is stale and never woken, which is the worst bug
a spreadsheet can have because it is silent. `RAND`, `NOW` and
`TODAY` are volatile: they recalculate whenever anything does, which
is a set the graph carries beside its edges.

Full-column references (`A:A`) arrive here, and they are the stress
this phase exists to apply. `=SUM(A:A)` must not create a million
edges — Phase 1 already proved a range edge wakes its formula once
and not once per cell, so this should hold, and the spec is what says
it did rather than that it was assumed to.

**Exit:** a conformance table — input, expected value, asserted as a
literal — covering every function and its error cases; and a budget
spec saying a `VLOOKUP` down a full column evaluates once per edit
inside that column and not once per row.

### Phase 12 — The formula editor

Everything a person uses while actually typing a formula: function
autocomplete with the signature and the current argument highlighted,
references coloured in the text and outlined in the grid as they are
typed, clicking or dragging a range into a formula mid-typing, F4
cycling `A1 → $A$1 → A$1 → $A1`, matching-paren highlight, an error's
`explain` shown under the cell rather than in a panel, and named
ranges — `Insert ▸ Name`, and the name box defining one.

It is separate from Phase 11 for a reason worth keeping: Phase 11
adds nothing that imports the framework and this phase is nothing
else. Keeping the seam is what lets the library stay runnable in node
and keeps `boundaries.spec.ts` meaningful instead of ceremonial.

The one genuinely hard piece is reference picking. While a cell is
open and the caret sits just after an operator, a click in the grid
must insert an address rather than move the selection — that is a
mode, and modes are where spreadsheets keep their worst bugs. It
belongs in pure logic with a spec, in the shape `SheetKeys` already
has, and not in a click handler.

**Exit:** Phase 4's standard again — driven through the semantics
tree, keyboard-only wherever keyboard will do — and the reference
picking written as a table of (caret context, click) → result.

### Phase 13 — Many sheets

Tabs along the bottom: add, rename, delete, reorder, duplicate,
colour. `Sheet2!A1` and `'Q3 Budget'!A1:B9` in the parser.
Three-dimensional ranges left out.

The key widens everywhere. The store is keyed by a packed row and
column; it becomes a sheet id and that. The dependency graph keys the
same way. `SheetContract` gains a `sheets` view key — the list and
the active one — and the viewport names its sheet, which is cheaper
than putting a sheet argument on every command and says the same
thing.

The claim this phase has to defend, and why it carries a patch-count
spec of Phase 2's kind: a formula on Sheet 1 depending on 50,000
cells on Sheet 2 publishes **nothing** while Sheet 2 is not the sheet
in view. Cross-sheet references are exactly where a naive
implementation starts publishing the whole workbook, and the failure
is invisible until the workbook is large.

Deleting a sheet turns every reference to it into `#REF!` and
renaming one rewrites them, which is Phase 10's rewrite machinery
aimed down a third axis. File format v3.

**Exit:** the cross-sheet patch-count spec above, and Phase 6's
reload proof run over a three-sheet workbook with cross-references in
both directions.

### Phase 14 — Formats that think

Conditional formatting and data validation.

The design is worth writing down before it is built, because the
obvious implementation is the expensive one. **A conditional format
is resolved at the window, not in the graph.** A rule is a formula
over a range; a rule covering a million cells does not need a million
graph nodes, because a format nobody can see does not exist. The
application worker evaluates each rule for the cells in the window at
publish time and folds the answer into the palette index Phase 9 is
already sending. A rule over the whole sheet costs the viewport,
which is the sentence this file keeps writing.

The catch, stated honestly rather than waved at: a rule that reads
other cells must re-resolve when those change. The window republishes
then anyway, so the answer is to re-resolve the window's rules on
every publish and then *measure* whether that is affordable rather
than assuming it. It is a few hundred formula evaluations per frame
at worst, which is the same order as a slice of the recalc pump the
worker already runs between publishes — but that is a prediction, and
this phase's job is to replace it with a number.

Data validation is the same shape: a per-cell predicate resolved at
the window, drawn as a marker, enforced at commit, with a dropdown
when the rule is a list.

**Exit:** a colour-scale rule over a million cells, scrolled, with
`pnpm proof` unchanged; and a patch-count spec saying the rule
published palette indices for the window and nothing else.

### Phase 15 — Charts

`<paint>` is the element — a box the application draws into over a
full path surface — so a chart is a component that subscribes to a
series and strokes it. Line, bar, column, stacked, area, pie and
scatter, with axes, a legend and a title.

This phase has an admission to make, and making it plainly is better
than breaking the invariant quietly. A chart is the first thing that
lets the render worker know about cells that are not on screen. It
does not learn the *cells*: it learns a **series**, built by the
application worker from the range and published on its own view key
at its own rate. And a chart 400 px wide should not be sent 50,000
points, so the worker downsamples to the chart's own pixel width
before it sends anything. "The render worker never learns about a
cell that is not on screen" becomes "the render worker learns a
series, at the resolution it can draw" — which is a weaker sentence
and a true one.

Charts float over the grid in the sheet's coordinate space: they
scroll with it and are hit-tested against it, and a chart is
selectable, movable and resizable. That is a small floating-object
layer the grid does not have, and it is what images and shapes would
later reuse.

**Exit:** a 50,000-point chart open while the sheet scrolls, with
`pnpm proof` unchanged; and a spec asserting the series published for
a chart 400 px wide holds at most 400 points.

### Phase 16 — Files that leave the tab

CSV import and export, `.gsheet` open and save-as through the File
System Access API, several documents open at once, recent files, and
a file dropped on the window.

**On `.xlsx`, which was an open question: CSV and local files now,
`.xlsx` import as the phase after this one, `.xlsx` export not yet.**
Reading xlsx is a zip reader and a handful of OOXML parts — sheets,
shared strings, styles, the calc chain ignored — and it is tractable
precisely because a reader is allowed to drop what it does not
understand. Writing xlsx is a fidelity contract with Excel, and every
part that cannot be round-tripped is somebody's lost work. Import is
what makes this a spreadsheet people can bring their data to; export
is what makes it a spreadsheet people can lose their data with. Do
the first, and let the second wait until something has actually been
imported.

CSV is not TSV with a different separator, and Phase 5's `Tsv.ts`
will not cover it: quoting, embedded newlines, a BOM, semicolon
locales, and the question of whether `=1+2` arriving in a CSV cell is
a formula. It is text. A CSV is data and not a program, and treating
it as one is how spreadsheet injection works.

**Exit:** Phase 6's proof, widened — import a CSV exported from
somewhere else, save it to the local disk as `.gsheet`, close the
tab, reopen the file, and the formulas added since recalculate.

---

## Still not in it, after all sixteen

Pivot tables, macros and scripting, collaborative editing, `.xlsx`
export, rich text runs *within* a single cell (formatting is
per-cell, and a bold word inside a cell is a different text model),
and touch.

Touch is the one with a date on it. `UiTouchScroller` is still
one-axis, noted in Phase 0 and still true, so a tablet cannot scroll
this sheet sideways at all. Whichever phase first wants a tablet
closes it.

## What Gesso still does not have

Carried forward from the head of this file, with what Part Two adds:

- **Menu bar traversal.** `Menu` is a popup with no notion of a bar
  above it. The roles exist; the arrow-across-the-menus model does
  not. Phase 8, probably upstream.
- **A mounted set that depends on the document.** `UiVirtualSheet`
  mounts what the window covers, and a merged cell anchored above the
  window still has to paint into it. Phase 10, and it is the deepest
  of the three.
- **A floating object layer over a scroll surface.** Selectable,
  movable, resizable things in a scrolled coordinate space, which
  charts need and images would reuse. Phase 15.
- **Two-axis touch scroll.** Still open, from Phase 0.

---

## Working against a local Gesso

This project installs Gesso from the sibling checkout rather than from
the registry — the packages are packed into `vendor/` and the manifest
points at the tarballs. Phases 0 and 5 expect to change the engine, so
the loop has to be cheap:

```bash
./scripts/vendor-gesso.sh              # build what is stale, pack, install
./scripts/vendor-gesso.sh --no-build   # pack whatever dist/ holds now
```

`vendor/` is gitignored: it is 2MB of tarballs that change on every
engine iteration, and the script regenerates them in seconds. A fresh
clone runs the script once, before `pnpm install`.

Each tarball's name carries a hash of its own bytes, and that is not
cosmetic. pnpm resolves a `file:` dependency by path and version and
then trusts its store, so an engine change that leaves the version
alone — which every change during a phase does — used to be re-packed
under the same name and never unpacked. `pnpm install` said "Already up
to date" and the app went on running the previous build.

Do **not** follow the generated `README.md`'s advice to re-run
`create-gesso-app --local --force` over this directory. It rewrites
`src/`, `README.md` and `package.json` from the template, which would
delete this project. That is what the script above is for.

To move to published packages once the engine work has landed upstream:
delete `vendor/`, delete `pnpm-workspace.yaml`, delete `overrides`, and
put version ranges back in `dependencies`.
