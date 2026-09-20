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

**Status:** scaffolded. Phase 0 has not started.

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
overscan band, as plain arrays of already-formatted display strings.

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

## What Gesso does not have yet

Two gaps found by reading the engine, both of which this project will
have to close. Both are plausibly worth upstreaming rather than working
around, and that decision should be made once the shape is known, not
now.

**Two-axis virtualization.** `LazyGrid` in
`core/src/composition/UiLazyList.ts` virtualizes rows only; columns are
declared tracks shared with the header through `subgrid: 'columns'`.
Nothing windows columns. This is the main engine work, and the reason
Phase 0 is a spike rather than a start.

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

### Phase 0 — Burn down the two risks

Timeboxed, throwaway code. Spike two-axis windowing and the viewport
round trip together, with fake data and no formulas.

**Exit:** a 1,000,000-cell empty grid scrolls at 60fps on both axes with
values arriving from the app worker, and the overscan band that keeps it
clean is a known number. If this cannot be hit, the rest of this file
changes.

### Phase 1 — The sheet model, headless

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

### Phase 2 — Contract and application worker

`SheetContract.ts` with the windowed view and the commands
(`setViewport`, `setCell`, `setSelection`, `undo`, `redo`), and the
worker serving it.

**Exit:** a spec that drives commands and asserts patch counts per edit.
One keystroke in a cell with 50,000 dependents should emit patches
proportional to the visible window, not to the dependents. That single
assertion is the whole thesis.

### Phase 3 — The grid surface

Frozen row and column headers through `position: 'sticky'`, which is
already conformance-tested against Chrome. Selection rectangle, active
cell ring, column resize by drag, A/B/C and 1/2/3 headers.

**Exit:** specs querying the semantics tree as `grid` / `row` /
`columnheader` / `cell`, and `toHaveBox` assertions pinning the frozen
panes.

### Phase 4 — Editing

An in-cell editor over `EditableTextModel`, a formula bar bound to the
same buffer, and the commit semantics people have muscle memory for:
Enter commits and moves down, Tab commits and moves right, Esc reverts,
F2 enters edit mode, typing over a selected cell replaces it. IME
composition inside a cell.

**Exit:** a keyboard spec in the style of `gesso-components`'
`Keyboard.spec.ts` — navigate, type, commit, undo, entirely through the
semantics tree with no synthetic mouse.

### Phase 5 — Clipboard and fill

Copy a range to TSV. Paste a range, which needs the second gap closed.
Then the fill handle, with relative references adjusted as it extends.

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

Do **not** follow the generated `README.md`'s advice to re-run
`create-gesso-app --local --force` over this directory. It rewrites
`src/`, `README.md` and `package.json` from the template, which would
delete this project. That is what the script above is for.

To move to published packages once the engine work has landed upstream:
delete `vendor/`, delete `pnpm-workspace.yaml`, delete `overrides`, and
put version ranges back in `dependencies`.
