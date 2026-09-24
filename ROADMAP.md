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

**Status:** Phases 0 to 4 done, all five exit criteria met. Phase 0's
findings are in [`PHASE0.md`](PHASE0.md); the sheet model is in
`src/sheet`, the contract, the application worker, the grid and the
editor in `src/app`, and `pnpm test` is its 205 specs. `pnpm dev` is a
spreadsheet you can type into. Phase 5 has not started.

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

### Phase 5 — Clipboard and fill

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

### Phase 6 — Persistence

A repository writing through `FileSystemSyncAccessHandle` in the app
worker, mirroring `OpfsNotesRepository` in Gesso's notes example.

**Exit:** the notes example's own proof, borrowed — type a marker,
reload, it is still there.

### Phase 7 — The proof surface

This is what the application is *for*, so it is a phase and not a
nicety. Recalculate 200,000 dependent cells while scrolling, and show
that the scroll re-measured nothing. A "block the main thread for five
seconds" button. `engine.explain` wired to a cell inspector. The
re-measure heatmap visible.

**Exit:** each of those is a thing a stranger can do in a browser in
under a minute, and a frame budget in CI that fails when one of them
regresses.

---

## Not in v1

Charts, pivot tables, multiple sheets, conditional formatting, and
merged cells.

Merging is the tempting one and the one that fights `subgrid: 'columns'`
hardest: it wants a cell to span tracks its row does not own. If merged
cells turn out to be required, they are their own engine phase, not a
detail of Phase 3.

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
