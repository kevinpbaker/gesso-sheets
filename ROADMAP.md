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

**Status:** Part One is done — eight phases, eight exit criteria met —
and [Part Two](#part-two--a-spreadsheet-rather-than-a-demonstration),
which turns the proof into a spreadsheet somebody would keep a budget
in, is three phases into nine: the top bar, the format axis and the
structural edits are done — insert, delete, borders, sort, hidden
rows and columns, frozen panes, merged cells, autofit and filtering.
`pnpm test` is 625 specs and `pnpm proof` is six budgets. Phase
0's findings are in [`PHASE0.md`](PHASE0.md); the sheet model is in
`src/sheet`, the contract, the application worker, the grid, the
editor and the chrome in `src/app`, and the proof strip in
`src/shell`. `pnpm dev` is a spreadsheet you can type into, copy out
of and paste into, find and replace across, fill down, format, rule
with borders, sort, insert and delete rows and columns, freeze a
pane, merge cells, fit a column to its contents, filter to what the
cursor is on, and navigate by typing an address — and which remembers
what you typed.
[Part Three](#part-three--quality-of-life) is the list of what a hand
that knows Excel reaches for and does not find yet.
`pnpm proof` is the frame budget: it drives the built application in headless
Chrome and fails the build when scrolling stops being free.

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
the track instead of dragging. `UiTouchScroller` had the same bug and
now has the same fix, through a `hasScrollRoom` the two paths share
rather than write twice.

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

### Phase 8 — The top bar — **done**

A menu bar, a toolbar beneath it, the name box beside the formula
bar, and a status bar along the bottom carrying the selection's Sum,
Average and Count — the readout people check before they trust a
column. Find and replace, fill down and fill right, go-to-cell, and a
sheet of shortcuts.

**Exit:** the whole bar driven from the keyboard through the semantics
tree with no pointer event anywhere in the spec — Phase 4's standard,
applied to the chrome. And `pnpm proof` with the bar mounted, plus a
fifth budget: the median frame while a menu is open over a scrolling
sheet.

**Met.** `TopBar.spec.tsx` is thirty-two specs and contains no pointer
event: F10 reaches the bar, letters open menus, arrows walk them,
Enter runs what it lands on, Ctrl+G jumps to a typed address, Ctrl+F
finds and Enter steps through the matches. `pnpm proof` grew the
fifth budget and holds all five. Measured on this machine:

| | median frame | worst |
|---|---|---|
| scrolling an idle sheet | 1.9 ms | 19.7 ms |
| scrolling while 200,000 cells recalculate | 2.0 ms | 6.1 ms |
| scrolling with a menu open | 2.2 ms | 5.1 ms |

The whole chrome costs about a tenth of a millisecond of median
frame, checked against a build of the previous commit rather than
against Phase 7's recorded numbers — the worst-frame figures moved on
this machine for reasons that have nothing to do with this code, and
the same run over the old build says 22.3 ms where the new one says
19.7 ms.

**Three menus and not seven.** File, Insert and Format would be menus
of things that do not work yet, and a menu of disabled items is worse
than no menu: it advertises, and then it refuses. They arrive with
Phases 16, 10 and 9.

Four things the phase learned by being built and then used.

**It is not `Menu` from the component set, and the reason is focus.**
`Menu` owns its keyboard and traps focus inside itself, which is
right for a popup opened by a button and wrong for a bar: with focus
trapped in the popup, ArrowLeft has nowhere to go, and walking to the
menu next door with the arrows is most of what makes a bar a bar. So
the traversal is a table — `MenuBarModel.ts`, nineteen specs, no
rendering — and the panel is an overlay that draws without taking the
keyboard. The shape is now known, which is the point at which the
upstreaming decision can be taken.

**Two tables answer the keyboard, and a spec keeps them honest.**
`SheetKeys` says what a key press does; `SheetCommands` says what the
menu advertises. Nothing in the language makes those agree, and the
failure is the quiet kind — the menu goes on printing `Ctrl+Z` next
to Undo long after the key has been rebound, and somebody learns the
wrong thing from the application itself. So every accelerator marked
`viaKeyTable` is pressed in `SheetCommands.spec.ts`, through the real
key table, and asserted to mean what the menu says. The shortcut
sheet is generated from the same table, and the navigation keys it
advertises are pressed too.

**Paste cannot be a menu item that pastes.** `ShellService` can put
text *on* the clipboard — that is the path Phase 5's copy takes — and
has no matching read, because reading the clipboard is gated on a
gesture inside a document and the render worker has no document. The
three options were a menu item that silently does nothing, no Paste
in the Edit menu at all, and a dialog that says which key to press.
The first is the worst thing software can do and the second sends
people looking.

**A dialog of plain text takes the keyboard and gives it to nothing.**
`Dialog` traps focus into itself as it opens, and
`UiFocusManager.settleScope` *blurs* when the scope it is settling
into holds nothing focusable — so the shortcut sheet, which is text,
left focus on no node at all. Escape reached neither the sheet nor
the dialog and the thing could not be dismissed without a mouse.
Found by pressing Escape in a browser; every spec passed without it,
because a spec that opens a dialog and asserts it is open never asks
what has the keyboard, and `Overlays.spec.ts` upstream gives every
dialog a button. Both dialogs here now carry a Close button, which is
better UI anyway and makes the focus settle. Seventh phase running in
which a real browser found something the suite could not.

### Phase 9 — The format axis — **done**

Number formats (general, number, currency, percent, scientific, date,
time and text), weight, italic, underline, size, text colour, fill,
alignment and wrap. A Format menu, a toolbar that shows itself
pressed, and the accelerators every spreadsheet binds.

**Exit:** a patch-count spec. Formatting a column of 50,000 cells
emits one palette patch and one patch per visible cell,
`toHaveLength` and not a ceiling. And `pnpm proof` over a formatted
sheet: the same budget, not a new one.

**Met.** `FormatChannel.spec.ts`: fifty thousand cells formatted at
once with thirty in view sends **one palette patch and thirty index
patches**, and the same edit with the column scrolled off screen
sends nothing at all. `pnpm proof` over the now-formatted seed reads
2.00 / 2.20 / 2.30 ms against Phase 8's 2.00 / 2.10 / 2.30 — the
format axis costs nothing measurable.

**The palette is the design.** `formats` carries a palette *index*
per visible cell and `palette` carries the entries, on two keys
because they move at different rates: a scroll would otherwise put
the whole palette on the wire and one click on Bold would put the
whole window back. Only the *paint* crosses; the number format stays
on the application worker, so what the render worker receives is the
finished string and it never learns a locale, a currency symbol or a
thousands separator.

Four things the phase learned, two of them from a browser.

**A format is applied as a change, not as a format.** `format` takes
`{ bold: true }` and every absent field means "leave it alone",
because a selection holding one bold cell and one plain one, told to
go italic, has to end up bold-italic and italic. Sending a whole
format would make every button on the toolbar destroy what the others
had done, which is the bug every naive formatting model has.

**The toolbar had to stop being fifteen tab stops.** Adding a Format
section put the grid fifteen presses away from the keyboard, and the
specs said so before a person could — `Tab never reached the grid` is
what Phase 9 got for the fourth button. It is now one stop with the
arrows moving inside it, which is ARIA's toolbar pattern and what the
menu bar already did. `focusable: false` is the opt-out that makes it
possible, and it is the same `tabindex="-1"` this file wanted for
`Dialog` in Phase 8 — so that gap is narrower than it looked.

**Every spec passed while nothing was bold.** They all asserted the
*document* — `formatAt(0, 0).paint.bold` — and the document was
right; the palette was simply never published on the load path, so
every cell pointed at an entry the render worker had never been sent
and fell back to plain. Number formats looked perfect throughout,
because those are applied on this side and cross as strings.
`FormatPaint.spec.tsx` now reads the properties the renderer draws
with, which is the assertion that was missing rather than the fix.

**Ctrl+A, then ctrl+B, wrote thirty megabytes to somebody's disk.**
A format stored per cell is a million entries for two keystrokes, a
million-entry undo step, and a file parsed back on every load for the
rest of that sheet's life. Phase 7 learned this about the proof
chain; this is the same lesson on the other axis. A region is now
stored **as a region** — sheet, row and column defaults resolved
cell, then row, then column, then sheet — and the file went from
30.32 MB to 2,365 bytes with nothing on screen changing. The wire is
untouched, because the render worker asks what a *cell's* id is and
the resolution happens behind that question.

And the correction inside the correction: the first version of the
region write *cleared* the cell overrides inside it, which looked
right and threw away the currency in column C the moment anybody made
the sheet bold. A change applies to the region and to each override
under it; "make this bold" has nothing to say about somebody's
currency symbol. One keystroke in a browser showed both.

Two things deliberately left out. **Borders**, because `borderWidth`
is a single number in the engine and per-edge borders need either
engine support or four child boxes per cell — see the gaps below.
And **a custom number-format pattern language**, Excel's
`#,##0.00;[Red](#,##0.00)`, which is a parser, a spec and a class of
bugs; the eight named formats cover what people pick.

### Phase 10 — Rows, columns, and borders — **done**

Insert and delete rows and columns with reference rewriting, per-edge
cell borders, sorting a range, hiding rows and columns, freezing a
pane, merged cells, autofit, and filtering.

**Exit:** a rewrite-count spec in Phase 1's style — inserting a row
above a column of 50,000 formulas rewrites exactly the formulas that
reference it, `toBe` and not `toBeLessThan`. And an insert performed
mid-scroll that does not drop a frame, asserted by `pnpm proof`.

**Met, both halves.** `Structure.budget.spec.ts`: 50,000 formulas, a
row inserted at the top, **49,999 rewritten** — `toBe` — and a row
inserted below everything rewrites none. `pnpm proof` grew a sixth
budget: a row inserted into the two-hundred-thousand-formula chain,
with the scroll running across it, costs 2.6 ms against 2.0 ms idle
where 4 ms is the allowance. The most expensive thing the application
thread does does not reach the frame.

**`Shift` is not `Rewrite`, and the difference is the thing every
implementation gets wrong exactly once.** A fill pins `$A$1`; an
insert moves it. The dollar sign says "do not move when I am copied",
not "do not notice the sheet" — and a reference that ignored an
insert would quietly point at somebody else's data. Shifting is
positional rather than relative, too: a formula in row 1 reading
`=A900` is rewritten by an insert at row 500 although it did not
move, which is why the walk is over the whole store. Ranges grow when
a row goes in, shrink when one comes out, clamp when a corner goes,
and break only when every cell they named is gone.

The store and the graph are **rebuilt rather than patched**, because
after a shift every key on the moved side is wrong and a patched
version can only ever be half right. Undo records what a shift
destroyed — the cells in a deleted row and the formulas it turned
into `#REF!` — because an opposite shift brings back neither.

**Borders needed no engine change, and this file was wrong to say
they would.** A border in Gesso is paint-only and the `decorated`
modifier already takes an Observable of arbitrary coloured
rectangles, drawn in the node's own paint pass with nothing to lay
out and nothing to hit test. A bordered cell is four draw instances
and no extra nodes. The shapes are *pushed* into a per-cell subject:
piped as a `combineLatest` per cell they cost 0.2 ms of median frame
and five milliseconds of input latency, measured — the same lesson
the value and the standing learned in Phases 0 and 3.

Three things found by using it rather than by writing it.

**A single-cell sort widened to the whole sheet.** It swept three
unrelated tables into one ordering and dragged formulas across each
other until some pointed off the sheet and said `#REF!`. One press of
ctrl-Z put it back, and it should never have been offered: "sort the
table I am standing in" means the *current region*, the block that
stops at the blank row, and where that block stops is a question only
the application worker can answer — the render worker holds the rows
it has mounted and the block may be bigger or smaller.

**A sort moved the values and left the formats behind**, so the
sorted table kept its bold total row where the total had been,
against somebody else's numbers. Formats travel with their rows now.

**`wrap` shipped in Phase 9 and did nothing.** The format was stored,
the toolbar showed itself pressed, and `Grid` drew `textWrap: 'none'`
regardless — and even wired, `LazySheet` takes one row height for
every row, so a wrapped cell has nowhere to put its second line. The
property is bound correctly now and the control is out of the
toolbar, on this file's own rule: a control that silently does
nothing is worse than one that is missing. *Back after Phase 17,*
once rows could be of different heights — see "Wrapped text and row
heights" there.

**Freezing a pane took three bugs to finish, and two of them were
the engine's.** `frozenRows` and `frozenColumns` mount the pane; the
renderer sticks it there with the `position: 'sticky'` the header row
and the gutter have always used. Which was the trouble: every sticky
node in the framework sticks at *zero*, so nobody had noticed that
`assignBox` offset a sticky node by its inset and
`resolveStickyOffset` then subtracted the inset again — an inset
applied twice, cancelling only at zero. A frozen second column, which
sticks at the width of the gutter, was drawn one gutter too far
along. An inset is a threshold, not a displacement, and it is fixed
upstream with specs on both axes.

The third was this application's. The frozen cells were correctly
placed, correctly stuck, and **empty**: the application worker
publishes the scrolled window, and a column frozen at A while the
sheet is scrolled to CL is not in it. The window is a *list* of rows
and columns now rather than a rectangle, so the pane and the window
both cross and the gap between them costs nothing — bounding it
instead would fetch five hundred rows to show one.

And the way all three were found is the phase's own lesson repeating:
the first specs asserted the frozen cell's `left` *property*, and
passed. Phase 3 added `toHaveVisibleBox` because a header asserted by
property passed while it scrolled off the screen; the frozen pane is
asserted by box now, and that is what caught every one of them.

**Autofit is the one thing neither thread can do alone**, and saying
so is the clearest statement of the split this project is about. The
application worker knows every string in a column and nothing about
fonts; the render worker knows the font and holds thirty rows. So the
application worker narrows a million cells to a shortlist — by
character count, which picks the right *candidates* even though in a
proportional font it picks the wrong *winner* — and the render worker
measures those exactly and sets the width.

Measuring needed an engine change, and a general one: layout measures
text constantly and an application could not ask, so anything wanting
a width before there is a node to read it from had no answer.
`TextService` hands out the runtime's *own* measurer, so the width a
column is given is the width its cells are laid out at — the same
warm cache, and the same invalidation when a font finishes loading.
The alternative every application reaches for is mounting the text
invisibly and reading its box a frame later, which is a frame of
latency, a node in the tree and a race with the layout it is trying
to inform.

**A filter is a snapshot, not a rule.** It hides the rows in the
current region whose cell in the active column is not the one the
cursor is on — the filter people actually use, and the one that needs
no dialog. Editing a cell afterwards does not re-run it, which is
what every spreadsheet does and what keeps an edit from making rows
vanish under somebody's hands. The filtered rows are a *second* set
beside the hidden ones, because the two are undone by different
things: clearing a filter must not reveal a row somebody hid on
purpose.

**A merged cell is drawn by its anchor and by nothing else.** The
anchor is as wide as the columns it covers and as tall as the rows,
and it overflows its own row downwards to reach them — rows are not
merged, only cells are, so there is nothing else a vertical merge can
be. The cells it covers are given no width and no height at all,
which keeps every other cell in the row at the offset the window put
it and is the only version that does not need the window to know
about merges.

What the window *does* need to know is where the anchors are, and
that is `extendRange`: a merge spanning C3:E3 is drawn by C3, so a
window starting at D has nothing to draw and the merge disappears at
the edge of the screen. This is the consumer that engine hook was
added for, which is the same discipline the row heights got — an
engine change with no consumer is one nobody has run.

Merging is destructive and knowingly so: the cells it covers have
nowhere to show what they held, so they are emptied. Every
spreadsheet warns about that; this one makes it one step of ctrl-Z
instead, which is the same promise kept differently.

**The two engine changes, and the first thing built on them.**
Merged cells and freeze panes needed the mounted set to be able to
include something outside the window; hiding rows, autofit and a
visible `wrap` needed a row height per row. Both are done — see the
list at the end of this file — and hiding rows is the first thing
standing on one of them, which is also how it was checked: an engine
change with no consumer is an engine change nobody has run.

And it took a browser to finish it. A hidden row is one of height
zero, and a zero-height row whose cells are also zero-height still
*paints* them — nothing in the engine clips a node to its box unless
it is asked to, so the text of the hidden row went on drawing over
its neighbours. Every spec passed: they asked what the cell's height
property was, and zero is exactly what they got. The row clips itself
now, and only the hidden ones do, because clipping every row would
cut off the fill handle that deliberately hangs outside its cell.

### Phase 11 — The library, and dates — **done**

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

**Met, both halves.** `Functions.spec.ts` is 210 cases over eighty-odd
functions, every answer a literal, run through parse → graph →
recalculate → read rather than against the function table, because
that is the path a formula actually takes and the wiring is where the
mistakes were. Its last spec fails the build if a name is added to
the library with no case asserting what it does.
`Library.budget.spec.ts` holds the other half: a `VLOOKUP` down a
full column is **one evaluation per edit** — `toBe`, not
`toBeLessThan` — and a thousand edits are a thousand, not a million.
All six frame budgets are unchanged at 2.0–2.6 ms.

**A date is a number with a format, and that is the whole design.**
No date type, because one would break the three things that make
dates work: `=B2-B1` giving a count of days, a date sorting with the
numbers, and `EOMONTH` returning something you can still add 7 to.
The serials are asserted against Excel's, including the phantom
1900-02-29 that this declines to invent — a serial that disagreed
would make every file leaving here wrong by a day, in silence,
forever. Typing one is a single undo step writing both halves, and
the format goes on only over `General`: a cell somebody deliberately
formatted has been answered already.

The parse refuses a date with no year. Excel reads `1/2` as this
January, which makes the stored value depend on the day it was typed.
`3/4/2026` is March the fourth because with no locale there is no
evidence in the text — but `24/9/2026` is read day-first, because the
other reading is not a date at all.

**The three dangers this file named were all real, and all three were
found by a spec rather than by reasoning.**

*Volatiles* wake on every edit, and so does everything downstream of
them — the half that was missing at first. A `=TODAY()` redone while
the `=A1+1` beside it is not leaves two cells disagreeing about the
date, which is worse than either being stale because one of them
looks right. `NOW()` is read once per recalculation rather than once
per call: the first version handed out a live clock and two cells in
one pass disagreed by however long the pass took.

*`INDIRECT` and `OFFSET`* run with their reads recorded and their
edges rebuilt from what they actually touched, so editing the cell an
`INDIRECT` landed on wakes the formula that read it. A formula that
read something still dirty is redone once it settles — *once*,
because two of them pointing at each other would otherwise chase each
other for as long as anybody watched. Making that terminate meant
teaching `recalculate()` with no budget to keep going when new work
appears during a pass, which it did not do.

*`A:A`* is watched, not expanded: 1,048,576 edges is not a thing to
store, so the graph keeps one entry per column per formula and
`dependentsOf` unions the watchers in. That is the interval index
this repository has wanted since Phase 1, built for the one case that
cannot live without it. Kahn's in-degree had to be counted forwards
through `dependentsOf` rather than backwards through `precedentsOf`,
because the backwards spelling cannot see a watch and there is no
entry for it to find. Reads stop at a high-water mark of what has
ever been written, which errs high and only ever costs a few blank
cells.

**`#N/A` is a sixth error value**, and the count in `Values.ts`
changed with it. It came with the lookups because a `VLOOKUP` that
found nothing has not failed — the formula is fine and the table
simply has no such row — and calling that `#VALUE!` sends somebody to
debug a formula that is correct. `IFNA` is only a coherent function
if the code is its own. `#NUM!` is still absent on purpose: `SQRT(-1)`
is `#VALUE!`, which says the true thing, and a seventh code earns
less than it costs.

**What `TEXT` cannot do is stated rather than half-done.** Excel takes
a format language here and this application deliberately had none —
see `NumberFormat`. So `TEXT` accepted the pattern strings that name
the formats it has and answered `#VALUE!` to anything else. A
half-implemented pattern language that ignored the parts it could not
do would produce text that is wrong rather than absent, which is worse
in a cell nobody is checking. *Superseded after Phase 17:* `TEXT` now
reads Excel's format codes in full, and is checked against Excel's own
answers — see "Checked against Excel" there.

### Phase 12 — The formula editor — **done**

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

### Phase 13 — Many sheets — **done**

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

### Phase 14 — Formats that think — **done**

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

### Phase 15 — Charts — **done**

`Paint` is the element — a box the application draws into over a
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

**Both met, and the numbers are these.** The proof grew a seventh
scenario: a chart of fifty thousand readings open over a scroll costs
a median frame of 4.30 ms against 4.00 ms without it, a difference of
0.30 ms against a budget of 2. The series claim is asserted twice —
in `Series.spec.ts` at every width that could fall off an end, and in
`Charts.spec.ts` end to end through the service, where the limit the
thinning is given is the chart's own width in pixels.

Two things are worth carrying forward.

**The thinning keeps extremes, not samples.** Every nth point loses a
spike one sample wide, and a spike is what somebody drew the chart to
find. Points are bucketed and the smallest and the largest of each
bucket are kept in the order they occur, so the envelope survives:
two per bucket, hence `limit / 2` buckets, hence never more than the
limit. On the proof page every peak of the sine still reaches exactly
±1000 through four hundred points.

**The floating object layer is a child of a row.** A chart hangs off
the row it starts in, so it travels with the scroll and its node is
never rebuilt by one; `extendRange` keeps that row mounted while any
part of the chart is on screen, which is the hook a merge reaching up
out of the window already used. The row holding one is lifted for the
same reason a merge's anchor is — without a stacking context the rows
below paint over it, and a chart came out twenty-four pixels tall in
a browser, which is exactly the height of one row.

Placement is in sheet pixels rather than anchored to a cell, which is
a choice with a cost: widen a column above a chart and the cells move
under it. An anchor is a translation into this coordinate space
rather than a different one, and the space had to exist first.

### Phase 16 — Files that leave the tab — **done**

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

**Met, in Chrome, and in node.** A CSV dropped on the window became a
sheet of its own; four formulas were added over it; Ctrl+S wrote a
`.gsheet` through a real `FileSystemFileHandle`; the tab was closed;
a new, blank tab reopened the file from *Open recent*, landed on the
same document by its handle, and recalculated 13.5, 30, 57.75 and
101.25 — "Ready · 4 evaluated". The one thing replaced was the native
dialog, by a function returning an OPFS handle, since automation
cannot click one; everything past it is the shipped code. The same
story runs in `Documents.spec.ts` against a second, empty library.
`pnpm proof` did not move: 4.30 ms median with the chart open, as
before.

Most of it was the engine's, and it is worth saying which parts.
Gesso could already turn a file drop into a drag — and nothing posted
one, so the shell half is new (gesso `625e31d`), and the first spec of
it found that every dropped file would have arrived with no bytes
(`1efb2d3`). Files through the shell did not exist at all: `openFiles`,
`saveFile`, `reopenFile`, `recentFiles` and `forgetFile` are new on
`ShellService` (`0fee892`), and Ctrl+S reaching the application rather
than Chrome's "Save page as" needed `interceptKey` (`059976c`).

Four decisions to carry forward.

**A tab is a document.** "Several documents open at once" is several
tabs, each at `/d/<id>`, each with its own copy in OPFS and an index
of them beside it. Holding many in one tab would have meant a service
that is not bound to one workbook, and a browser already has tabs.
`/` is the last used, `/d/new` a blank one, and both are replaced with
the real id as soon as it is known, so a reload reopens what was
there. The first document keeps `gessosheet.json`, so a profile from
before this phase opens where it was.

**A document is not a file.** The copy in OPFS is saved on every edit
whether or not anybody chose a file; a file is where it was last saved
*to*. That is how a tab closes without a Save and loses nothing, while
Save still means something — and why the status line says *edited*
only for a document that has a file to differ from.

**A handle crosses as a number.** The shell keeps every
`FileSystemFileHandle` in IndexedDB and hands out its key, which is
all "recent files" is. A download has no handle, so the next Save asks
again — the honest answer in Firefox and Safari, which cannot write
back to a file they did not open.

**A CSV is data.** `=1+2` arriving in one is text, formatted Text so it
stays text through an edit and a reload; on the way out, a text value
another program would run gets OWASP's apostrophe, and a negative
number, which is not text, does not.

**A read that fails is not an empty file.** Found after the phase, by
opening one workbook in two tabs: a sync access handle is exclusive
across the origin, a tab refused one read "nothing stored", and the
service saved a fresh, empty workbook over the real one. The index
did the same with its list. Both now wait for a held file and throw on
one they cannot read, and nothing is opened or saved over it; a
document is claimed with a Web Lock, and a second tab shows it without
keeping it.

Found on the way and fixed with it: a viewport naming a sheet that is
not showing used to *switch* to it, which made two tab changes in one
turn a ping-pong neither thread could leave. Two files dropped at once
found it. A viewport now answers for the sheet that is showing, and
switching is `activateSheet`'s alone.

### Phase 17 — `.xlsx` import — **done**

The phase Phase 16 named: an Excel workbook opened as a document of
its own, and nothing written back. Reading is a zip, a handful of XML
parts and a mapping, and it is tractable for the reason given there —
a reader may drop what it does not understand.

**Kept:** cells, formulas, number formats, bold, italic, underline,
font size and colour, fills, borders, alignment and wrap, column
widths, hidden rows and columns, frozen panes, merges and defined
names. **Dropped, deliberately:** charts, pivot tables, conditional
formats, validations, comments and images, each of which this sheet
would show wrongly rather than not at all. Theme and indexed colours
were on this list at first and came off it: Excel writes nearly every
colour a person picks as a theme slot and a tint, so a workbook from
Excel lost most of its colour until they were read.

**Exit:** a workbook written by another spreadsheet opens with its
formulas calculating, across sheets and through names. **Met.** The
fixture is `src/sheet/fixtures/orders.xlsx`, written by LibreOffice's
own Excel filter from `scripts/xlsx-fixtures.py`, so the reader is
tested against a producer and not against its author's idea of the
format. In node it opens as two sheets and calculates $141.23 through a
`SUM`, $11.30 through a defined name and again through `Rates!B1`, and
shows `2026-09-24` and `007` as they were. In Chrome, the same file
dropped on the window opened as its own document at its own route and
drew its merged title, fills and formats.

Four things worth carrying forward.

**A formula that cannot run keeps its value, and says so.** The
fixture's due dates are `=DATEVALUE(...)`, and this sheet has no
DATEVALUE. Written as a formula it would show `#NAME?`; kept as the
serial LibreOffice calculated it shows the date, and the status line
says four formulas became values. A formula that silently turned into
a number is a sheet that stops updating without saying so.

**An `.xlsx` becomes a snapshot**, the shape a `.gsheet` already has,
and opens by Phase 16's path. Nothing downstream knows it was ever an
Excel file — which is also why Save afterwards writes a `.gsheet` and
asks where: there is no writer, and a Save that silently changed the
format of the file it was opened from would be worse than one that
asks.

**The reader carries its own XML and zip.** The application worker is
not a window, so it has no `DOMParser`; the XML reader is a small one,
safe because it expands no entity it did not define, and the zip reader
takes its inflate from the caller so that `src/sheet` still imports
nothing. A sheet of 100,000 cells and 10,000 formulas reads in about a
quarter of a second.

**A file crosses as base64.** A command carries plain data and an
`ArrayBuffer` is not, so the bytes travel a third larger than they are.
That is the cost of the barrier for a thing the size of a file, paid
once per open, and it is written down here rather than hidden.

Not measured, at first: a file Excel itself wrote. Excel's shared
formulas, `_xlfn.` prefixes and inline strings were specced from the
format's own shapes, but no Excel was available to write one.

**Checked against Excel, after the phase.** Apache POI keeps several
hundred `.xlsx` files in its test data, most written by Excel, and an
Excel file holds each formula's last answer beside the formula. So the
check is: open each, calculate, and compare with what Excel saved. Of
352 files, 323 now agree in every formula. Seventeen are refused, all
rightly — fuzzer output, a truncated zip, and three encrypted files,
which now say they are password-protected rather than "not a zip". The
twelve that still differ do so for reasons known and left: `TODAY` in a
file saved years ago, a sheet wider than this one, a circular formula
Excel resolves by iteration, array formulas, an external workbook, and
a few of POI's deliberately odd arguments.

What it found, and is fixed: a single-cell reference passed to a
function is a range (so `SUM(A1)` of text is 0); implicit intersection;
a blank equal to `""`, 0 and FALSE; `""` in arithmetic is `#VALUE!`;
the 1900 phantom leap day, so every serial before March 1900 was a
day out; `0^0` and friends; `ROUND` at fifteen significant digits;
`XLOOKUP`'s every argument, and its binary search bisecting as Excel's
does even on unsorted data; zip entries with Windows backslashes; and
the 1904 date system of Mac workbooks, whose dates are moved onto this
sheet's calendar as they come in.

The largest was `TEXT`. POI's format workbooks hold about eight hundred
codes with the text Excel produced for each, and `FormatCode.ts` now
matches all but eleven of them: sections and conditions, literals,
padding and fills, currencies, placeholders, thousands and scaling
commas, percent, scientific and engineering notation, fractions,
dates, twelve-hour times and elapsed time. The eleven are fraction
codes with literals wedged between the placeholders, where Excel's
answers contradict each other. A sample of the cases is
`FormatCode.spec.ts`; POI is Apache-2.0 and is credited there.

**Wrapped text and row heights, after the phase.** A row is as tall
as what it holds, and a row somebody dragged is as tall as they made
it. The engine had carried a sparse height per row since Phase 10 and
hiding rows was its only consumer; this is the second, and it needed
no engine change.

Fitting a row is autofit turned on its side, and split the same way:
the application worker knows which cells wrap, how wide their columns
are and what they say, and knows nothing about fonts; the render
worker can measure and holds thirty rows. So the application worker
sends the rows that could be taller than one line — a wrapped cell or
a large font, with something in it — and the render worker measures
them at the column's width through the layout's own measurer and
sends the heights back. A row sent with nothing in it goes back to
the default, which is how emptying a wrapped cell brings its row
down. It is asked after an edit, a format, a width, an undo and a
recalculation; a sheet with no wrapping and no large fonts anywhere in
its palette answers with one look at the palette, and that is nearly
every sheet. The fitted heights are saved, so a sheet opens at the
heights it had rather than growing a frame after it appears, and an
answer to a question since replaced is dropped by its serial.

A height set by hand wins, which is Excel's rule: a row told what it
is keeps it whatever goes in it, until "Fit rows to contents" gives it
back. Excel's own row heights come in from an `.xlsx` only where the
file says somebody set them (`customHeight`); the rest were Excel
fitting its rows, and this sheet fits its own.

Found on the way: an undone row insert put the widths back and left
the hidden rows where the insert had moved them, so undoing an insert
above a hidden row revealed it and hid its neighbour. Row heights and
hidden rows are held by the step now, and both come back. And pressing
the new grip swept a selection as well, since the gutter is over a
row of cells as far as the offsets are concerned; it stops the pan
the way the fill handle always has. Checked in Chrome: a sentence
wrapped in B21 grew the row to four lines and moved everything below
it, kept that height through a reload, and a drag on the row's foot
resized it without moving the selection.

**A bar too wide for its window, after the phase.** The rules bar is
wider than a window of 900 pixels, and its value field was the one
thing in it allowed to shrink — so on anything much narrower than a
full-width monitor the value a rule had been given was drawn zero
pixels wide, and the bar looked as if it had not read the rule at all.
That was the report "it only shows up after I click another cell",
and it had survived an engine fix for a different zero-height bar
because a spec asking whether the field *existed* passed. The rules
and find bars wrap onto a second line now, their fields keep their
widths, and the specs ask for the width.

**Arrays, after the phase.** A formula whose answer is more than one
value spills it into the cells beside and below its own, as Excel 365
does: `=A1:A3`, `=A1:B3*2`, `=IF(A1:A9>1, "big", "small")`,
`TRANSPOSE`, `SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`, and an `XLOOKUP`
that returns a whole record. A cell in the way stops it with
`#SPILL!`, and clearing that cell lets it through. Arithmetic over
ranges handed to a function arrives as an array, which is what
`SUMPRODUCT((A1:A9>1)*B1:B9)` and `SUM(A1:A3*B1:B3)` need.

The cells an array spills into hold nothing of their own. Reading one
reads the array, and each carries one edge in the dependency graph —
to the formula — so everything that reads a spilled cell is ordered
after the formula that fills it by the same topological sort as every
other cell, and nothing about spilling is saved: the formula is the
cell, and the rest is its answer.

The one real decision was what an older file means. In an `.xlsx`
written before Excel 365, `=B1:B9*2` in C5 means `B5*2` — implicit
intersection — and this sheet would spill it. So a formula read from a
file is given the `@` Excel 365 itself shows when it opens one,
wherever the difference would show: before a range or a named range
that is an operand, or the test of an `IF`, and never inside
`SUMPRODUCT`, which has always worked on arrays. `_xlfn.SINGLE` is
how Excel 365 writes `@` into a file, and it reads back as one.
Excel's own array formulas (`t="array"`) come in as the formula alone,
and the values Excel kept in the cells they cover are dropped so the
formula can fill them. Checked against POI's files: both of the ones
that were failing on arrays now agree with Excel, 325 of 352 in all.

**`.xlsx` export, after the phase.** File ▸ Download as Excel workbook
writes the whole workbook: values and formulas, number formats, fonts,
fills, borders, alignment and wrap, column widths, row heights, hidden
rows and columns, merges, frozen panes, defined names and iterative
calculation — what the reader brings in, taken back out, so a workbook
from Excel goes back with what it came with. Conditional formats,
validations and charts are left out, as the reader leaves them out,
and the line that reports the download says which of them the
workbook had.

Formulas are written as Excel reads them. A function newer than the
format carries Excel's `_xlfn.` prefix; a formula that spills is an
array formula over the cells it fills, with their values beside it as
Excel writes them; and one that does arithmetic over ranges without
spilling — `=SUM(A1:A3*B1:B3)` — is an array formula of one cell,
because written plainly an Excel from before dynamic arrays would take
one value from each range. It is the reader's `@` rule run the other
way, and uses the same code to decide.

The engine had to learn one thing: `saveFile` wrote text, and an
`.xlsx` is a zip. It takes bytes now, which cross the barrier as they
already did on the way in. The zip writer is this project's, beside
the reader, and deflates through the platform's `CompressionStream`.

Checked twice. The seeded workbook, written and read back through the
reader, is cell for cell what it was, formats, widths and panes
included. And LibreOffice, given the same file headless, recalculated
it and showed every cell of the sales sheet — dates, currency,
percentages, a spilled array, an array `SUM`, an `XLOOKUP` — exactly
as this sheet does, with no error anywhere in the workbook.

**Iterative calculation, after the phase.** A circle is `#CIRC!`,
which is almost always what a circle is — a mistake. A workbook can
ask otherwise, as Excel's can: Data ▸ Work out circular formulas by
iteration goes round each circle in Excel's order, sheet by sheet and
row by row, from the values the last round left, until nothing moves
by a thousandth or a hundred rounds have gone. An `.xlsx` that turns
it on (`calcPr iterate`) keeps it on, with its own count and step, and
so does the `.gsheet` it is saved as. The circle is gone round after
the rest of the recalculation, because it may read cells that have
just changed. POI's one file that uses it now agrees with Excel to
within its step and not to the last digit, and cannot: the answer
depends on how many times the file was recalculated before it was
saved, which the file does not say.

Not yet: `A1#` for "the whole spill", and a function lifted over an
array (`ABS(A1:A9)` takes one value, where Excel 365 spills nine). The
formula bar over a spilled cell came in Phase 21.

---

## Still not in it, after all seventeen

Pivot tables, macros and scripting, collaborative editing, rich text
runs *within* a single cell (formatting is
per-cell, and a bold word inside a cell is a different text model),
and touch.

Touch was the one with a date on it, and the date has passed:
`UiTouchScroller` took one axis per container, noted in Phase 0 and
true for ten phases, so a tablet could not scroll this sheet sideways
at all. That is fixed in the engine now. What touch still wants before
a tablet is in scope is this application's own: targets sized for a
finger, a long press that means what a right click means, and a
selection that can be dragged without a hover to show what it would
select.

## What Gesso still does not have

### ~~A bar switched in by a menu choice is never laid out~~ — **fixed**

`Edit ▸ Find…` had been giving a find bar of zero height since Phase
8, for anybody using a mouse; Ctrl+F gave the same bar at full size.
Phase 14's conditional formatting bar inherited it, and the menu was
its only route in.

The cause was in the engine. `markLayoutDirty` walks a changed node's
ancestors setting `measureDirty` and stops at the nearest relayout
boundary — sound only while the pass really does start at that
boundary. It does not always: another node dirtied in the same frame
can drag the layout root into the set, and then the pass is a
`fullLayout` that measures from the root, down through ancestors the
walk left clean, every one of which short-circuits on its own measure
memo. A menu's own teardown is exactly what drags the root in, which
is why the menu was the route that failed and an accelerator was not.

Fixed in gesso by carrying the walk past the boundary to the root
while still recording the boundary as where a bounded relayout
starts. Both bars now measure 1400x32 by either route.

It took a browser to find and a browser to prove: the shape does not
reproduce in either repo's harness, so there is no unit spec, and
that is written down in the engine's commit rather than papered over
with one that passes when the fix is reverted.

Carried forward from the head of this file, with what Part Two adds.
Most of it is struck through now. Two of these were closed by the
engine while Phase 10 was running, and the rest in one pass over the
list afterwards, which is the argument for keeping a list like this in
the application rather than in the engine: every one of them was found
by trying to build something, and none of them would have been found by
reading the engine.

What is genuinely left is first-class per-edge border properties,
which `borders()` may well have made unnecessary. The floating object
layer that used to be named here turned out not to need the engine at
all: see the struck-through entry below.

- ~~**Menu bar traversal.**~~ *Closed, and upstreamed.* `Menu` is a
  popup that traps focus, so the arrows could not reach the bar to
  move along it, and Phase 8 worked around that by owning the
  traversal here. `menuBarStep` is now a peer of `Menu` in
  `gesso-components`, generic in the command type, with its own specs
  against its own menus; `MenuBarModel.ts` is the four lines that bind
  it to `CommandId`, and its spec still runs against this
  application's real menus, which is the second consumer the upstream
  version needed. The *chrome* stays here: a third of `MenuBar.tsx` is
  this application's accelerator column, mnemonic underlines and
  theme, and generalising a component against one consumer is how a
  library gets props nobody wants.
- ~~**A focus trap with nothing to focus.**~~ *Closed, and this file
  had guessed the fix right.* `UiFocusManager.settleScope` blurred
  when the innermost scope held nothing focusable, so a `Dialog` whose
  content is plain text handed the keyboard to nothing and could not
  be dismissed with Escape. Both halves this file named turned out to
  be needed and neither was enough on its own: `tabStop`, which is
  `tabindex="-1"` — focusable, reachable by a press and by `focus()`,
  skipped by the cycle — and a `settleScope` that falls back to the
  scope root before blurring. `Dialog` sets `focusable: true,
  tabStop: false` on its body. The prediction that the one-line
  version breaks `Overlays.spec.ts` was right, and the two-property
  version does not.
- ~~**A mounted set that depends on the document.**~~ *Closed.*
  `UiVirtualSheet` mounted what the window covered, so a merged cell
  anchored above the window could not paint into it and a frozen row
  the sheet had scrolled past was not there to be frozen. Two
  additions rather than one, because the two cases want different
  things: `extendRange` widens the window the viewport implies, which
  is what a merge needs and costs nothing when nobody widens
  anything; `frozenRows` and `frozenColumns` are a *second* mounted
  set, because widening the window back to row 0 from row 5,000 would
  mount five thousand rows to show one. Both are the header row and
  the gutter — which have always done this — with the count turned
  up, and keeping them visible stays the renderer's job with sticky
  positioning.
- ~~**One row height for every row.**~~ *Closed.* `columnWidth` was a
  number or an array and `rowHeight` was only a number, so a column
  could be hidden by setting its width to zero and a row could not.
  `rowHeights` is a **sparse map** rather than an array, and the
  asymmetry is the point: a sheet has a few hundred columns and up to
  a million rows, so a height per row would be the largest allocation
  in the application to describe a sheet where every row but two is
  the same. The offsets stay a multiplication plus a binary search
  over the exceptions, so a sheet with none pays what it always paid.
- ~~**A floating object layer over a scroll surface.**~~ *Closed, and
  it needed no engine change at all.* The expectation was a new
  primitive; what it took was `Paint` — which already existed — and
  the observation that a child of a row scrolls with the row. The two
  hooks it leans on, `extendRange` and a lifted row, were both grown
  for merges. Images and shapes would reuse the same three lines.
- **Per-edge borders — and the note this used to be was wrong.**
  `borderWidth` is one number and `borderColor` one colour, so a cell
  cannot have a heavy bottom edge and a hairline top. This file then
  said the only way round it was four child boxes per cell, four times
  the nodes on the one surface whose node count is a budget. That is
  not true. A border in Gesso is paint-only — `toHtml.ts` says so
  outright, and `paintBorder` strokes *inside* the box — so it touches
  no layout at all, and the `decorated` modifier already takes an
  Observable of arbitrary coloured rectangles drawn in the node's own
  paint pass with **nothing to lay out and nothing to hit test**. Four
  thin rects per bordered cell is four draw instances and zero nodes.
  Phase 10 can have borders without an engine change.

  What was genuinely missing is *closed*, and it was not an anchoring
  convention. `DecorationBox` now takes `right` and `bottom`, and the
  rule is CSS's for an absolutely positioned box: any two of near
  edge, size and far edge fix an axis. An anchor alone would have
  covered "the bottom edge" and not "the span between the two
  horizontal borders", which is what a side edge is and what decided
  the shape. `borders()` in `gesso-core` is the general modifier that
  needed it — this application still does its own, because it pushes
  per-cell shapes into a subject and a modifier per cell is not the
  same thing, but anything that is not a hundred thousand cells can
  now just ask.

  First-class `borderTopWidth` props are a different and much larger
  change, and the cost is not where it looks: the property, the paint
  state and the Canvas2D stroke are easy, but the WebGPU instance
  packs `radius, opacity, borderWidth` as one `float32x3`, the
  fragment shader draws the border as an isotropic SDF band, and one
  instance carries one colour — so four widths need a wider vertex
  format, an anisotropic inner rect in the shader, and up to four
  instances when the colours differ.
- ~~**Two-axis touch scroll.**~~ *Closed.* `UiTouchScroller` picked one
  axis per container from its flex direction, exactly as
  `UiWheelController` did before 38e70a3, so a finger could not move
  this sheet sideways at all. The same fix on the same shared state,
  through a `hasScrollRoom` the two paths now share rather than write
  twice, and a fling that passes the threshold per axis so a vertical
  throw does not drift sideways by the thumb.
- ~~**The engine on the thread that does not run it.**~~ *Closed, and
  it was never on this list because nobody had weighed a bundle.*
  `createApp` took either the worker options or a root component, and
  the overload that took a component reached `GessoAppBuilder` →
  `GessoApp` → `GessoRuntime`: the layout engine, both renderers and
  the hit-tester, statically, in every shell. A bundler cannot see
  which half of one function a call reaches, so this application's
  main thread downloaded and parsed all of it in order to create a
  canvas and forward input. The single-thread configuration is now
  `createSyncApp`, and `check-bundle-size.ts` holds the line with a
  budget and with a look for Canvas2D calls in the shell's bytes.
  Measured here: **152.7 kB gzipped to 45.0**, and what is left is the
  proof panel, which is the one thing on this thread that is supposed
  to be.
- ~~**A fan-out that scales.**~~ *Closed, and this application is why
  it exists.* Every prop taking an Observable is the whole binding
  model and it has a cliff: N cells each reading their own slice of
  one channel is N pipelines per emission, and RxJS removes an
  observer by scanning a list, so tearing down a window is quadratic.
  This file's Phases 0, 3 and 10 each learned that separately and each
  wrote the same `Map` of subjects with one writer — 0.2 ms of median
  frame and five milliseconds of input latency, measured, for one
  piped binding per cell. `fanOut` in `gesso-framework` is that
  pattern with the lesson attached, and its `changed` hint is where the
  frame is won: one key changing costs one read rather than ten
  thousand. The grid here can drop its hand-rolled version whenever
  somebody wants to; it has not, because the hand-rolled one works and
  a rewrite of the hot path wants its own phase.

---

# Part three — quality of life

Part Two made a spreadsheet somebody would keep a budget in. Part
Three is about the second hour of keeping it: the gestures that
people's hands make before they have decided to, and which this sheet
answers with nothing.

None of it is a feature anybody would put on a list, and that is the
problem. Clicking a column's letter selects the column in every
spreadsheet ever made; here it does nothing, and the only way to hide
a column is to drag a selection across it first. Ctrl+Down goes to the
end of the data everywhere else; here it goes nowhere. Each one is
small, and each one is the moment someone who came from Excel decides
this is a toy. A reviewer will not report them. They will just stop.

So the bar for this part is not "does it exist" but "does the hand
that already knows Excel find it where it reaches". Where Excel and
Google Sheets agree, do what they do. Where they disagree, do Excel's,
and say so. Where Chrome owns the key — Ctrl+=, Ctrl+N, Ctrl+T, Ctrl+W
— the Ctrl+Alt fallback Phase 10 already uses is the answer, and the
shortcuts sheet has to say it.

**The rule Part Two runs under still holds**, all three of it. Two of
these phases put new handlers on the header row and the gutter, which
are mounted for every visible row and column, and a handler per header
is exactly the per-cell binding cliff this file has now learned three
times. The pointer work goes through one handler on the header strip
that works out what was pressed, the way the grid's own press does.

---

### Phase 18 — The pointer — **done**

Clicking a column letter selects the column, and a row number the
row. Shift+click extends from where the selection was, and a drag
across the headers sweeps several — the same sweep the cells already
have. Clicking the corner above row 1 selects the sheet. Every command
already written for "the selected columns" starts working from the
mouse the moment this does.

A double-click on a column's grip fits it to its contents, and on a
row's grip fits the row: Phase 10's autofit, reached where everybody
looks for it. A double-click on the fill handle fills down as far as
the column to its left goes, which is the single most-used fill in
anybody's sheet and today takes a drag across however many thousand
rows there are.

Right-click opens a context menu, built from the command table like
every other menu so that it cannot drift: cut, copy, paste; insert and
delete of whatever is under the pointer — rows on a row header,
columns on a column header, either on a cell; hide; sort; clear. It is
a menu, so it inherits the top bar's keyboard model, and the
Menu key and Shift+F10 open it at the cursor for anybody without a
second button.

**Exit:** a spec for each gesture through the semantics tree, and a
header click on a sheet of a million rows that selects the column
without walking it: the selection is a range, and neither selecting
it nor the status bar's Sum over it may cost a million of anything. `pnpm proof` with the
context menu open over a scrolling sheet, under Phase 8's fifth
budget.

**Met.** `Pointer.spec.tsx`, 28 specs through the semantics tree and
the hit tester: a letter, a number and the corner select the column,
the row and the sheet; Shift extends from where the selection was, a
drag sweeps, and one that wanders down into the cells keeps sweeping
letters; a grip's click selects nothing and its drag sweeps nothing;
a double-click on a column's grip fits it, or every selected column
when it is one of theirs, on a row's grip gives the row back to its
contents, and on the fill handle fills down as far as the column
beside goes — left, then right, then nothing — in one step of undo.
The right button opens the menu for what is under it, selecting that
first when it is outside the selection; Shift+F10 and the Menu key
open it at the cursor. A click on column B of a sheet a million rows
tall selects it and the status bar adds it up in well under a frame's
worth of anything, because both are a range. And `pnpm proof` has a
sixth run, scrolling with the context menu open, held to the fifth
budget's allowance: 3.6ms median against 2.3ms idle.

Three things worth carrying forward.

**What was pressed is read from the node, not the point.** The first
version worked the header out from the pointer's coordinates, which is
what a drag needs and what a right-click is given. But a click a
screen reader sends has no point at all, and neither does a spec that
clicks a node; so a click walks up from the node it landed on to the
strip's own nodes, which already say what they are — the column in
`posInSet`, the row in the label, and the corner is a button called
*Select all*. Coordinates are the fallback, for the two gestures that
only have them.

**The strip is drawn over the rows.** The spec helper that clicked
"row 1" aimed at the row's own box, which on a sheet scrolled by half a
row starts under the sticky header — so it had been clicking the
column's letter all along, harmlessly, until a letter meant something.
The helper aims below the header now. The pointer was right.

**A column click does not scroll.** The active cell of a whole column
is its first row, which is where Excel puts it, and the one move of
the selection that must not bring the cursor into view: somebody who
clicked a letter half way down is looking at the half they clicked.

Done in two sessions at once, and said so because it will happen
again: one wrote this phase's specs and the engine's `DoubleClick`,
the other the implementation, and they met in the middle — the specs
ran against an implementation their author had not seen, and passed
once the two agreed which end of a Shift+click moves.

### Phase 19 — The keyboard Excel users already have — **done**

`SheetKeys` answers arrows, Home, End, PageUp and PageDown, and
nothing with Ctrl on it but Home and End. The missing half:

- **Ctrl+Arrow** to the edge of the data — the next non-empty cell
  after an empty one, or the last before one — and **Ctrl+Shift+Arrow**
  to select there. This is the key people use to find out how long a
  column is, and it needs the sheet model to answer "next occupied
  cell in this row or column" without scanning a million, which is a
  question for `src/sheet` and gets its own budget spec.
- **Ctrl+Space** selects the column, **Shift+Space** the row.
- **Ctrl+Enter** commits what was typed into every cell of the
  selection, relative references shifted as fill shifts them.
- **Alt+Enter** puts a line break in the cell being edited instead of
  committing it. Wrap is already there; without this, the only way to
  get a second line into a cell is to paste one.
- **Ctrl+;** types today's date, **Ctrl+Shift+;** the time — as
  values, not `TODAY()`, which is the difference that makes them
  useful in a log.
- **Enter inside a selection** walks the selection and wraps, rather
  than leaving it, and **Tab then Enter** comes back to the column the
  Tab started from — how everybody types a row of a table.

Each goes into the `NAVIGATION` table beside the switch that answers
it, so the shortcuts sheet lists it without anybody writing it down.

**Exit:** `Keyboard.spec.tsx` covers every key above, and the
shortcuts sheet's generated list contains every one of them — a spec
that fails when a key is answered but not advertised. And a budget
spec in `src/sheet`, in the style of `Sheet.budget.spec.ts`: the next
occupied cell down a column of a million, found without visiting
them.

**Met, all but one.** Every key above is in `NAVIGATION`, and
`SheetKeys.spec.ts` now asks both directions: every advertised key
does something, and — the new half — every *kind* of thing the table
can do, over every key and chord it answers, open or not, is reachable
from an advertised key or a command's accelerator. It failed on its
first run, on Enter in an open cell, whose entry said "commit" and was
only ever pressed with nothing open. `Keyboard.spec.tsx` drives each
key through the grid, and `Edge.budget.spec.ts` puts one cell at the
foot of a column of 1,048,576 rows, beside a thousand-row table, and
jumps to it: 0.6 ms against 45 ms for the walk it replaces, on the same
sheet in the same run, and the spec holds the ratio rather than the
milliseconds. Checked in Chrome, which takes none of these keys for
itself: Ctrl+Down from E4 of the seeded sheet stops at the total in
E14 and then at E10000, Ctrl+Space selects E1:E10000, Ctrl+; writes
`2026-09-26` and shows it as a date, and Alt+Enter wraps its cell and
grows the row.

**Ctrl+Arrow is asked of the other thread.** Where the data ends is
the sheet's to know, and the render thread holds a window of it. The
command carries the selection it starts from, because this side's
selection leads the worker's by a frame and a second Ctrl+Down pressed
inside it would otherwise jump from where the cursor used to be. The
search steps along a run of filled cells, which costs the run, and
across a gap steps sixty-four cells and then reads what the sheet
holds, once — so its cost is the sheet's contents, never the distance.
A cell an array spilled into is filled, because it looks filled.

**Ctrl+Enter writes from the cursor.** The cell being typed into is the
corner Shift moved, not the anchor as in Excel, because that is where
this sheet opens the editor — so a formula typed there is in that
cell's frame, and the copies are moved from it. One transaction, one
undo, and the selection stays put.

**A typed line break turns wrap on**, in the same step as the text, as
a typed date brings its format: a cell that does not wrap draws one
line, and the break somebody typed would be there and not be seen.

**Tab then Enter** remembers the column the run of Tabs began in, and
typing into the cells between them does not end the run; an arrow, a
click, a jump or a selection does.

**Not done: Enter walking the selection.** It needs an active cell that
is not a corner of the selection, and this sheet's selection is an
anchor and a cursor — the cursor *is* the active cell, and it is always
a corner. Excel's model is a range plus an active cell anywhere inside
it; that is a change to `SheetSelection` and to everything that reads
one, which is a phase's worth of its own and not a key.

**Found on the way, not fixed:** text typed faster than a frame into a
selected cell keeps only its first character. The first key opens the
cell and the rest arrive before the editor has the keyboard. A person
typing does not outrun a frame, and a browser automation's `type`
does; it predates this phase, and a spec that types through the real
input path in one burst is where it should be pinned down.

### Phase 20 — Paste that asks — **done**

Paste today brings everything: value, formula, format. What people
want at least as often is less.

- **Paste values only** — Ctrl+Shift+V, which is also what Chrome
  means by it, so the text arrives already plain. Formulas land as
  their answers.
- **Paste formats only**, and its pointer form, a **format painter**:
  a toolbar button that takes the selection's formats and puts them
  on the next thing clicked, and a double-click that keeps it lit
  until Escape.
- **Paste transposed**, from Edit ▸ Paste special, next to the other
  two.
- **The copied range is marked** with a moving dashed outline until
  it is pasted, Escape is pressed, or the sheet is edited — so it is
  possible to see what Ctrl+V is about to do. A cut that is pasted
  moves the cells and the references that point into them, the way
  Phase 10's structural edits already move references; a cut today is
  a copy and a clear.

**Exit:** specs for each paste against the clipboard harness, and one
for a cut-and-paste of a range that a formula elsewhere reads, which
reads the new place afterwards. Undo takes each back as one step.

**Met.** `PasteSpecial.spec.tsx`, 15 specs through the clipboard
harness: values only from the menu and from Ctrl+Shift+V, a formula
landing as its answer in the formats already there; formats only;
transposed, from a copy made here and from text from anywhere; a cut
that moves the cells so that `=SUM(A1:A2)` elsewhere reads
`=SUM(C5:C6)` afterwards, a range only partly cut left alone, a
formula cut to another sheet naming the sheet its other references
were on; Escape calling a cut off; the outline marked by a copy, kept
through a paste, taken away by an edit or Escape; and the painter, lit
for one click or held by a double-click until Escape, without touching
the clipboard. Each paste and each move is one step of undo.

What was found on the way is in the phase's own first sentence.
"Paste today brings everything" was not true: a paste brought what had
been *typed*, and nothing of the formats. A copy is a snapshot now — the
inputs, the values and the formats of the block as it was — and a paste
of the text it put on the clipboard is a paste of the snapshot, so
formats come with it and "values only" has values to paste.

A cut no longer empties anything until it is pasted, as in Excel,
where a cut never pasted changes nothing. Pasted, it is a move by what
a reference *names* rather than where the formula sits: whatever
pointed into the block points at the same cells in their new place,
and the block's own formulas keep their other references where they
were — which is not what a fill does, and is why it is `Move.ts` and not
`rewriteFormula`.

The outline is drawn in the rows' own paint pass, as a formula's
references are, so it costs no nodes; only the rows the window has are
given their dashes, so a whole column copied costs thirty rows of them;
and the timer that moves them runs only while there is an outline. It
was checked by specs and not by eye: the browser window was on
another workspace for this one, and a hidden tab draws no screenshots.

### Phase 21 — Typing less — **done**

- **AutoComplete from the column.** Typing the first letters of a
  value that already appears above in the same column offers the rest,
  selected, so Enter takes it and typing on replaces it — Excel's
  behaviour, and the thing that makes a column of categories
  bearable. It stops at the first blank row, as Excel's does, which is
  also what keeps it proportional.
- **Undo says what it undoes.** Edit ▸ Undo reads "Undo sort", "Undo
  paste", "Undo delete columns", and the toolbar's tooltip does the
  same. The history already knows; the menu should say.
- **The formula bar shows the spilling formula**, greyed, over a
  spilled cell — the item Part Two left undone, and the reason a
  spilled cell currently looks like a value that came from nowhere.

**Exit:** a budget spec for AutoComplete on a column of a hundred
thousand entries — a keystroke's offer costs what a column of ten
does — and a spec that the spilled cell's formula bar is the
formula's, and cannot be edited as if it were its own.

**Met.** `Words.budget.spec.ts` asks the same five prefixes of a column
of ten and a column of a hundred thousand: 0.37 µs a keystroke against
0.28, and a row typed at the foot of the long column plus the offer
after it costs 3 µs. `Keyboard.spec.tsx` types through the grid — the
first letter offers, more than one word waits, Backspace takes the
offer and nothing else, a number and a formula get nothing — and has
five specs on the spilled cell's bar, one of which fails with the bar's
`readOnly` taken away. Checked in Chrome: C under the regions offers
Central with *entral* selected, Enter writes it, Edit ▸ Undo then reads
*Undo typing in A15*, and the cell under `=SEQUENCE(3)` shows the
formula in the bar, in grey.

**The words are an index, not a walk.** The plan above said
AutoComplete stops at the first blank row, as Excel's does, and that
this is what keeps it proportional. It is not: a column of a hundred
thousand with no gaps is one block, and walking it is a hundred
thousand reads a keystroke. So the words are the whole column's, kept
in `src/sheet/Words.ts` sorted and case-folded, with a count per word so
that clearing one of a thousand Norths keeps North. The workbook builds
a column's words the first time it is asked, keeps them up to date on
every write after that, and drops them all when cells move wholesale.
The answer is a binary search, and "one word, or nothing" is the
neighbour after it: *Nor* offers nothing beside North and Northwest,
which is how AutoComplete avoids putting the wrong region on a row.
The difference from Excel is that a word below a gap is offered too,
which is what LibreOffice does.

**The question crosses, and a late answer is dropped.** The render
thread holds a window of the column and the words are in all of it, so
each letter sends one question and the answer names the cell, the
prefix and a serial. An answer about text that has since changed is
thrown away rather than written over what is being typed. The word
comes back as the column spells it, so *nor* becomes North and Enter
writes the word that is already there.

**The history did not know.** The plan said it did. A step was a list
of edits and nothing more. Now `transact` takes a name, and the
outermost name wins, so a sort that writes a thousand cells is a
sort. A step nobody named is described from what it holds: one cell's
text, with the format a typed date brings, is *typing in B4*; a
structural edit says *insert 2 rows*. The toolbar's tooltip needed an
engine change to follow the label: `tooltip` now takes a function for
its text, asked each time it opens.

**The spilled cell's bar is read-only.** It shows the formula the value
came from, greyed. An edit begun there would start from the formula
and write it into the spilled cell, which blocks the array. Typing
over the cell from the sheet still replaces it, as in Excel, and the
array reports `#SPILL!`.

### Phase 22 — Knowing where you are

- **Zoom** from View, Ctrl+Alt+= and Ctrl+Alt+-, per sheet, saved with
  the document. It scales the grid, not the chrome. This one is not
  free: the grid's row and column offsets are pixels, and a zoom that
  merely scales the paint leaves hit-testing and the fill handle in
  the wrong place. It wants a design note before it wants code.
- **Show formulas** — Ctrl+` — draws every cell's formula instead of
  its answer, columns widened to fit, for checking a sheet somebody
  else wrote.
- **The status bar's readout is chosen.** Sum, Average and Count stay
  the default; a click on it offers Min, Max and Count numbers, and
  a click on a figure copies it.
- **Notes on a cell.** A small corner mark, the text on hover and on
  the keyboard focus, and Shift+F2 to write one. `.xlsx` comments come
  in as notes, and `.gsheet` saves them. This is the one item here
  with a file format attached, which is why it is last.

**Exit:** a zoomed sheet whose click, drag, fill and resize land on
the cell under the pointer at 50%, 100% and 200%, and `pnpm proof`
run at 200%, which is four times the cells' area in paint and the
same number of cells mounted.

---

## Not in Part Three

Touch — its own phase, for the reasons recorded above. Flash fill,
which is a feature and not a habit. Drag-to-move a selection by its
border, which is worth having but needs Gesso's Drag and the grid's
Pan to agree about who owns a press on a cell edge, and that is an
engine question first.

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
