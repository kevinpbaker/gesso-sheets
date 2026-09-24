# Phase 0 — what the spike found

**The exit criterion is met.** A 1,000,000-cell sheet (10,000 × 100)
scrolls at 60fps on both axes with every visible value arriving from
the application worker, and the band that keeps it clean is **8 rows
and 2 columns** at a 9,000 px/s fling — with 16 × 4 as the setting that
was clean in every run.

| run | axis | px/s | frame gap | frame cost | frames with a blank cell |
| --- | ---- | ---- | --------- | ---------- | ------------------------ |
| `confirm-y-3000` | vertical | 3,000 | 16.62 ms | 4.17 ms | 0% |
| `confirm-y-9000` | vertical | 9,000 | 16.69 ms | 4.91 ms | 0% |
| `confirm-x-3000` | horizontal | 3,000 | 16.69 ms | 4.63 ms | 0% |
| `confirm-x-9000` | horizontal | 9,000 | 16.67 ms | 5.18 ms | 0% |

The numbers here are from a run on one Linux machine under software
rendering; what matters is the shape of the differences, which held
across four runs, not the third decimal.

**The harness is gone, and the numbers are not reproducible from the
current tree.** `pnpm phase0` built the spike, opened it in headless
Chrome with `?bench`, ran twenty-eight configurations and printed the
table; the spike it drove was retired in Phase 3, when the real grid
replaced it. Everything here reproduces from commit `e4ea6e1`, which
is the last one where `src/spike` and `scripts/phase0.ts` exist.

That leaves the project without a frame budget in CI, which Phase 7
asks for by name — and it should be rebuilt against the real grid
rather than the spike, driving *input* rather than writing scroll
offsets. Both times hand-scrolling found something the bench could
not, it was because the bench went in through a side door.

So the phases in `ROADMAP.md` stand. What follows is what changes
inside them.

---

## 1. The engine gap is real, and smaller than it looked

`LazyGrid` virtualizes rows only, and the roadmap called closing that
"the main engine work". It turned out to be about 400 lines, in
`gesso/packages/core/src/composition/UiVirtualSheet.ts`, plus ten lines
in `GessoRuntime` and one registered property.

It is a new class rather than a second axis on `UiVirtualWindow`,
because a sheet does not need most of what that class is. `LazyColumn`
learns an item's extent by mounting it and measuring it, correcting an
estimate as the user scrolls, which buys the correction table, the
scroll anchoring and a per-frame walk that collects extents. A sheet is
*told* its row height and column width, so every offset is exact, the
window is arithmetic on two scroll offsets, and the rows a scroll
reveals are placed correctly on the frame that reveals them rather than
corrected on the one after. Dropping measurement is what makes the
class small enough to read.

The runtime hook is the part that cannot be done in userland: the
window has to be advanced *before* layout, from the scroll offsets the
frame is about to draw with. There is no element event that reports a
scroll, and a component reacting to one would always be a frame behind
— which is the blanking this phase exists to measure, added on purpose.

## 2. `subgrid: 'columns'` is the wrong tool for a sheet

`LazyGrid`'s rows share column tracks with a header, and `DataTable`
needs that because its column widths come from cell content.
`LayoutEngine.collectSubgridColumns` pays for it by measuring **every
cell of every mounted row, at unbounded constraints, on every layout**
(`LayoutEngine.ts:1536`) — the tracks cannot be sized until the cells
that decide them have been.

A sheet is told its column widths, so a row is an ordinary `Row` of
fixed-width cells and the tracks are never sized at all. The layout
phase costs 0.6–1.5 ms with ~100–250 nodes measured per frame, against
the ~1,500 cells that are mounted.

This is a code-reading finding with a measurement beside it, not a
head-to-head: the subgrid version was never built. It is why the sheet
element does not use `LazyGrid`, and it is the reason merged cells stay
out of v1 — they want a cell to span tracks its row does not own, which
is the one thing this design gives up.

## 3. The band belongs on the fetch side, not the mount side

This is the most useful thing the spike found, and it was not obvious.

There are two ways to keep values ahead of the eye. The **mount band**
is `LazySheet`'s overscan: mount more rows than fit, so the range asked
for is wider. The **fetch band** is the application worker publishing
beyond what it was asked for. Both widen coverage. They do not cost the
same.

At 9,000 px/s down the sheet, with the other band at zero:

| band | frames with a blank cell | frame gap | peak nodes |
| ---- | ------------------------ | --------- | ---------- |
| fetch 4 / 1 | 99.3% | 16.70 ms | 166 |
| fetch 8 / 2 | 0% | 16.68 ms | 138 |
| fetch 16 / 4 | 1.4% | 16.90 ms | 167 |
| fetch 24 / 6 | 0% | 16.67 ms | 140 |
| mount 8 / 2 | 12.3% | 17.15 ms | 566 |
| mount 16 / 4 | 47.2% | 23.77 ms | 688 |
| mount 24 / 6 | 31.0% | 35.45 ms | 2097 |

The mount band is a trap. It buys the same lookahead and pays for it in
nodes on every frame — and past about eight rows it makes the frame
slower, which makes the sheet move further between frames, which loses
more coverage than the wider band gained. `mount 24/6` spends 17.85 ms
a frame in the windowing phase alone and is still blank a third of the
time. The fetch band costs cells on the wire and nothing else: from
4/1 to 24/6 the mounted node count does not move.

**So: keep the mount band small (2 rows, 1 column — enough to cover the
partial rows at the edges) and put the lookahead in the contract.** The
roadmap already says the application worker publishes "that window plus
an overscan band"; this is the measurement that says the band should
live there and only there.

Coverage collapses sharply between a band of 4 and a band of 8: 99% of
frames blank at 4, 0–2% at 8 and above. The knee is not gradual because
it is a race — either the band covers a frame's movement plus the round
trip, or it does not.

**A band covers a scroll, not a jump.** Found by scrolling the spike by
hand after the matrix was already written, which is the second thing
hand-scrolling caught that the bench did not. The bench sweeps at a
constant velocity, so it only ever asks the band to cover one frame's
movement. Real input is not always like that: a wheel notch, a trackpad
fling, a scrollbar track click and a thumb drag all deliver a single
large delta, and a 2,500 px one moves the viewport 104 rows between two
frames. No lookahead of 16 rows covers that, and none of 104 would be
worth having.

With the fetch band at 16/4, continuous scrolling is clean well past
the speed the matrix stresses, and a jump is not:

| input | speed | frames with a blank cell |
| ----- | ----- | ------------------------ |
| 100 px per frame | 6,000 px/s | 0% |
| 200 px per frame | 12,000 px/s | 0% |
| 2,500 px per event | a fling | 43%, reaching 100% |

So the band is sized for velocity and there is no size that answers a
teleport. What answers a teleport is the other mitigation this file's
parent roadmap lists — a pending treatment that does not read as an
empty sheet — and it is worth saying that the *first* of those two,
keeping last-known values, is exactly wrong here: after a jump the
values previously held belong to entirely different rows, so showing
them is wrong data rather than stale data. A cell with nothing in it
yet should look like a cell waiting, which makes this a question for
Phase 3's cell treatment rather than for the contract.

**The spike now defaults to 16/4** rather than to 0/0. Zero was the
pessimal setting the knee was measured from, and leaving it as the
default meant the demonstration shipped in the configuration the
measurement had just ruled out.

## 4. The wire shape the roadmap assumed costs 80× the bytes

The roadmap says the application worker publishes "plain arrays of
already-formatted display strings". Measured, at 3,000 px/s with the
window fully covered:

| shape | patches per publish | bytes per publish |
| ----- | ------------------- | ----------------- |
| `keyed` — absolute row, then column | 6.2 | 1.1 KiB |
| `rows` — row-major arrays | 1,199 | 90 KiB |

The reason is in `diffArray`. A scroll shifts every element, so no
common prefix or suffix matches; the two windows are the *same length*,
so it does not splice, it recurses element-wise, and every row differs,
so it recurses again — one `set` patch per cell. The array shape turns a
one-row scroll into a full re-send of the window, cell by cell. Keyed by
absolute coordinates, the same scroll leaves every other key
structurally equal and the patch is the row that entered and the row
that left.

**Phase 2's contract should publish a map keyed by absolute
coordinates, not arrays.** Note this also makes the block
self-describing: a cell that is absent is a cell that has not arrived,
which is what makes "missed" countable at all.

One caveat, from a third shape added to test it: keying is not
symmetric. Row-major costs 184 patches on a horizontal fling (two per
mounted row); column-major costs 303 on a vertical one. Both still held
60fps here, so the volumes are survivable either way — but a sheet is
scrolled vertically far more than horizontally, so row-major is the
right default, and this is worth re-checking when real formatting makes
cells bigger than eight characters.

## 5. The cost is allocating bindings, not layout and not the wire

The first three runs had the horizontal axis at 21–24 ms a frame while
the vertical held 16.7 ms — with a *lower* frame cost and, in one
configuration, a *twentieth* of the wire traffic. Both obvious
explanations were wrong, and the phase breakdown found it: the
`virtualize` phase cost 9–11 ms horizontally against 1.7–3.4 ms
vertically.

A column window that moves invalidates every mounted row, because its
cells changed. Rebuilding a row rebuilt its cells, and building a cell
allocates two RxJS pipelines — about 1,100 of them, one or two times a
frame. Memoizing the cell elements by the cell they hold (`App.tsx`)
dropped the horizontal `virtualize` phase to 1.7–2.6 ms and the frame
gap to 16.6–16.8 ms, which is what closed the exit criterion.

The general shape: **on this design the per-frame cost is dominated by
constructing cell elements and their bindings, not by layout, paint or
the barrier.** Layout is 0.6–1.5 ms, paint 2.3–2.7 ms, patch
application under 1 ms in every configuration that matters. Phase 3
should treat cell-element identity as a performance contract, not an
implementation detail.

## 6. A third engine gap: input is one-axis

Found by scrolling the spike by hand, which is the part the bench never
did — it drives the sheet by writing `scrollX` / `scrollY` directly, so
every number above was measured with the input path bypassed
completely. Two-axis *windowing* was the risk this phase set out to
burn down; two-axis *input* was not on the list, and it should have
been.

`UiWheelController` asked each scroll container for one axis —
`ScrollContainerState.horizontal`, derived in `GessoRuntime` from
whether the container lays its children out in a row. A sheet's
viewport is one ScrollView whose content overflows both ways, so it was
classified as vertical, every horizontal delta was dropped, and the
chain then found no horizontal container to hand it to either. A
trackpad could not scroll the sheet sideways at all. The sink's
`scrollBy` always took both deltas; only the routing was one-axis.

Fixed by asking each axis whether it has room and giving it the delta
it can use, which leaves a single-axis container exactly as it was —
the axis it does not scroll has no room, so it is not taken. A page
delta is now converted in each axis's own extent rather than one
borrowing the other's.

Two smaller things came out of the same session, both in the scrollbar:

- **The thumb's grab target was the six pixels it paints.** Missing it
  is not a near miss — the press lands on the track and pages a whole
  screenful, which reads as the bar refusing to be dragged. The grab
  box is now the thumb's length by the full band that reveals it, which
  is what an overlay scrollbar is expected to be. What is drawn has not
  changed.
- **The corner where the two bars meet belonged to the vertical one.**
  `scrollbarZoneAt` has to answer with one axis and checked vertical
  first, so a press aimed at the horizontal thumb down in the corner
  paged the vertical bar. A press that lands on a thumb now belongs to
  that thumb's axis whatever the precedence says.

`UiTouchScroller` is still single-axis: it reads the same `horizontal`
flag to pick the axis a drag and its inertia run along, so a two-axis
container is panned only along whichever axis that names. Not fixed
here — the inertia model is per-axis and making it two-axis is a real
change, not a routing one — and not exercised by this phase, which has
no touch input in it. It is the next thing to go wrong on a tablet.

**What this says about the spike:** a harness that drives the thing
under test through a side door measures the thing and not the product.
Phase 7's proof surface should drive input, not offsets.

## 7. A recalc that blocks the publish blanks the sheet

Burning 8 ms on the application thread before each publish costs
nothing measurable. Burning 30 ms — one recalc of a moderate dependency
graph — leaves 88.7% of frames with a blank cell and reaches 100%
uncovered, at an ordinary 3,000 px/s scroll, with the band that is
otherwise clean.

The render thread stays at 60fps throughout, which is the thread model
working exactly as advertised: the sheet keeps scrolling smoothly and
simply has nothing to put in the cells. So the scroll is never the
casualty, but "the formula engine is on another thread" is not on its
own enough to keep the sheet filled.

**This is a constraint on Phases 1 and 2, not a problem with Phase 0.**
Serving the viewport has to be able to overtake a recalc in flight —
either by yielding between dependency levels, or by publishing the
window from the last settled values and letting the recalculated ones
follow. Phase 2's patch-count spec should be joined by a spec that
edits a large graph *while* the viewport moves.

---

## What is throwaway and what is not

Throwaway: everything in `src/`. `SheetContract.ts` has no edits, no
selection and no formulas; `SheetSource.ts` computes cell text from a
hash; `App.tsx` is a readout with a sheet under it. Phase 1 starts from
an empty `src/`.

Worth keeping:

- **`UiVirtualSheet` and `LazySheet` in Gesso**, with their spec. This
  is the engine work the roadmap deferred, and the shape is now known,
  so the upstreaming decision it wanted to defer can be made. It is
  written as a peer of `UiLazyList` and has no sheet-specific
  vocabulary in it.
- **`scripts/phase0.ts`**, which was Phase 7's proof surface in
  miniature. Retired in Phase 3 with the spike it drove; the shape of
  it is what Phase 7 should rebuild, against the real grid and through
  the input path.
- **The fix to `scripts/vendor-gesso.sh`.** pnpm resolves a `file:`
  tarball by path and version and then trusts its store, so an engine
  change that leaves the version alone — every change during a phase —
  was re-packed under the same name and never unpacked: `pnpm install`
  said "Already up to date" and the app kept running the previous
  build. The tarballs now carry a hash of their own bytes. This cost an
  hour of believing a fix had not worked.

## What is still unknown

- **Frozen panes.** Phase 3 wants `position: 'sticky'` headers inside
  this scroll container, and nothing here has a header in it. The
  sticky path is conformance-tested against Chrome, but not against a
  two-axis window.
- **Touch.** `UiTouchScroller` pans a two-axis container along one axis
  only (§6). Nothing here has a finger in it.
- **Non-uniform geometry.** `UiVirtualSheet` takes one row height and
  one column width. Column resize is Phase 3, and it turns every offset
  from a multiplication into a prefix sum — which is cheap, but changes
  `rowAt`/`columnAt` from arithmetic into a search.
- **Real formatting.** Cells here are at most eight characters of
  ASCII. Wider strings move the wire numbers and the paint cost, and
  nothing else.
- **Whether the frame gap is trustworthy.** Headless Chrome schedules
  frames however it likes, so `cost` — the work the render worker did —
  is the number that carries the argument, and it never exceeded 20.6
  ms even in the configurations that were deliberately bad. The gap
  agreed with it everywhere, which is why both are reported.
