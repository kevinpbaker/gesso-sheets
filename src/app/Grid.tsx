import { BehaviorSubject, combineLatest, type Observable } from 'rxjs';
import { map } from 'rxjs/operators';

import {
  Box,
  decorated,
  EditableText,
  editorFor,
  LazySheet,
  percent,
  Row,
  Text,
  type DecorationShape,
  type UiElement,
  type UiKeyboardEvent,
  type UiNode,
  type UiPasteEvent,
  type UiModifier,
  type UiPointerEvent,
  type UiTextChangeEvent,
  type SheetRange,
  type UiVirtualSheet
} from 'gesso-core';
import {
  fanOut,
  FocusService,
  internalState,
  ShellService,
  TextService,
  type ComponentContext,
  type FanCell,
  type Inputs
} from 'gesso-framework';

import { columnName, relativeRef } from '../sheet/A1';
import type { Span } from '../sheet/Tokenizer';

import {
  CELL_FONT_SIZE,
  CELL_PADDING,
  COLUMN_COUNT,
  COLUMN_WIDTH,
  MAX_COLUMN_WIDTH,
  GUTTER_WIDTH,
  HEADER_HEIGHT,
  MIN_COLUMN_WIDTH,
  ROW_COUNT,
  ROW_HEIGHT
} from './dimensions';
import type { CellEdge, CellPaint } from '../sheet/Format';
import { cellIn, PLAIN_PAINT, Sheet, type SheetMerge, type SheetSelection, type SheetWindow } from './SheetContract';
import { colouredReferences, formulaSpans } from './FormulaColours';
import { pick, repick } from './FormulaEditing';
import { commandFor } from './SheetCommands';
import { keyAction } from './SheetKeys';
import type { SheetEditing } from './SheetEditing';

/**
 * The sheet, on screen.
 *
 * Everything a person sees is built here and painted in the render
 * worker; the cells come over the channel as display strings and this
 * file has never heard of a formula. What it owns is the window, the
 * two frozen strips, the selection, and the drag that resizes a
 * column.
 *
 * Three rules carried in from earlier phases, all of them measured:
 *
 *   - **Cell elements are memoised by the cell they hold.** Phase 0's
 *     entire frame budget went on allocating RxJS pipelines until they
 *     were, and a column window that moves invalidates every row. This
 *     is a performance contract, not an implementation detail.
 *   - **The band lives on the fetch side.** The mount band is two rows
 *     and one column, which covers the partial rows at the edges and
 *     nothing more; the lookahead is the application worker's, where
 *     it costs cells on the wire instead of nodes on every frame.
 *   - **A cell with no value yet looks like a cell waiting.** A band
 *     covers a scroll and cannot cover a jump, so this still happens
 *     on a fling; what it must not look like is an empty sheet.
 */

/** Where a cell stands in the selection: outside it, in it, or the one. */
type Standing = 0 | 1 | 2;

/**
 * A cell on screen: the element, and the two subjects that feed it.
 *
 * The row and column are kept beside them because the feeds walk this
 * map every time the window or the selection moves, and parsing them
 * back out of the key was the only reason the key had a shape.
 */
interface MountedCell {
  readonly row: number;
  readonly column: number;
  readonly element: UiElement;
  readonly value: FanCell<string | null, { row: number; column: number }>;
  readonly standing: FanCell<Standing, { row: number; column: number }>;
  /**
   * How the cell is painted, pushed in like the other two.
   *
   * A third subject rather than a pipe off the channel, on the rule
   * this file has followed since Phase 0: a cell subscribes to its
   * own subjects and one subscriber per source fills them. A cell
   * that piped its own paint off the palette would put every mounted
   * cell on the palette's observer list, and RxJS removes an observer
   * by scanning that list.
   */
  readonly paint: BehaviorSubject<CellPaint>;
  /**
   * The cell's border rectangles, pushed in like everything else.
   *
   * Piped instead — `combineLatest([paint, width])` per cell — this
   * cost 0.2ms of median frame and five milliseconds of input
   * latency, measured. It is the same lesson the value and the
   * standing learned in Phase 0 and Phase 3: a cell subscribes to its
   * own subjects, and one writer fills them.
   */
  readonly shapes: BehaviorSubject<readonly DecorationShape[]>;
}

const SELECTED_WASH = 'selectionBackground';
const GRID_LINE = 'border';

/**
 * The editing handle for the screen around the grid.
 *
 * The formula bar has to write the *same* buffer the cell does — not a
 * copy that syncs, which is two sources of truth and a race over which
 * of them Escape puts back. There is one `SheetEditing` per screen and
 * both views are bound to it, so this is where it is made and
 * `SheetApp` is where it is shared.
 */
export function Grid(_inputs: Inputs<{ editing: SheetEditing }>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const shell = ctx.inject(ShellService);
  const measure = ctx.inject(TextService);
  const window$ = sheet.view.window;
  const edit = _inputs.editing.value;
  const selection$ = edit.selection;

  /**
   * The cells now mounted, by the cell they hold.
   *
   * An element is immutable, so a cell still in the window can simply
   * be handed back. Without this a column step rebuilds every row and
   * every row rebuilds its cells, and building a cell allocates its
   * bindings — about eleven hundred of them, once or twice a frame.
   */
  const cells = new Map<string, MountedCell>();

  /**
   * The window and the selection, pushed into the cells rather than
   * piped into them.
   *
   * This is the sheet's second performance contract, and it was bought
   * with a profile. A cell used to pipe its own value off `window$` and
   * its own standing off the selection, which reads well and puts every
   * mounted cell on the subscriber list of two subjects — about four
   * hundred and sixty of each, three times over, because each property
   * that reads a pipe subscribes to it again. RxJS removes a subscriber
   * by scanning the observer array for it, so tearing a window down
   * costs the *square* of what the window holds.
   *
   * That is invisible at a wheel's pace, which retires four rows a
   * frame. It is not invisible on the scrollbar: ten thousand rows in a
   * seven-hundred-pixel viewport pin the thumb at its twenty-four pixel
   * minimum, so one pixel of thumb travel is fourteen rows and a drag
   * replaces the whole window every frame. Measured on this machine,
   * that spent 38ms a frame inside `arrRemove` — 42% of the render
   * worker — and drew at 22fps. The same drag draws at over a hundred
   * once each cell subscribes only to subjects of its own.
   *
   * **That shape is `fanOut` now, and the lesson lives in the framework
   * rather than in this comment.** One subscription per source however
   * many cells are live, a stable cell per key, and a release that is a
   * map delete. What is kept here is why it matters, because the
   * profile that bought it is this application's and the primitive
   * carries only the rule.
   *
   * The row and column travel beside the key rather than inside it.
   * `fanOut` keys are strings because a `Map` wants one, and a reader
   * that took `${row}:${column}` apart would do it for every live cell
   * on every emission — which is exactly what the first draft of this
   * migration did, and the reason `fanOut` grew a datum. Measured three
   * times against the hand-rolled version it replaced: the same
   * numbers, to the tenth of a millisecond.
   */
  let latestWindow: SheetWindow | null = null;
  let latestSelection: SheetSelection | null = null;

  /** Where a cell is, carried beside its key so no read parses one. */
  interface At {
    readonly row: number;
    readonly column: number;
  }

  const values = fanOut<SheetWindow, string | null, At>(
    window$,
    (current, _key, at) => cellIn(current, at.row, at.column),
    { initial: null }
  );

  const standings = fanOut<SheetSelection, Standing, At>(
    selection$,
    (selection, _key, at) => standingOf(selection, at.row, at.column),
    { initial: 0 }
  );

  ctx.effect(window$, current => {
    latestWindow = current;
  });

  ctx.effect(selection$, selection => {
    latestSelection = selection;
  });

  /**
   * The formats, resolved through the palette on the way in.
   *
   * The window carries an index and the palette carries the entry,
   * and a cell wants neither — it wants the paint. Resolving here
   * means the lookup happens once per changed cell per publish rather
   * than once per bound property per frame, and it is the same shape
   * as the two effects above for the same reason.
   *
   * `paintOf` returns the *same object* for every unformatted cell,
   * so the `!==` below is a pointer comparison that is false for the
   * whole window on a sheet nobody has formatted, and nothing is
   * pushed at all.
   */
  let latestFormats: { cells: Readonly<Record<string, Readonly<Record<string, number>>>> } | null = null;
  let latestPalette: readonly CellPaint[] = [PLAIN_PAINT];

  const paintOf = (row: number, column: number): CellPaint => {
    const id = latestFormats?.cells[row]?.[column] ?? 0;
    return id === 0 ? PLAIN_PAINT : (latestPalette[id] ?? PLAIN_PAINT);
  };

  const repaint = (): void => {
    for (const mounted of cells.values()) {
      const next = paintOf(mounted.row, mounted.column);
      if (next !== mounted.paint.value) {
        mounted.paint.next(next);
      }
      refreshShapes(mounted);
    }
  };

  /**
   * A cell's borders, recomputed only when they could have changed.
   *
   * `bordersOf` hands back the *same* empty array for a cell with no
   * borders, so the comparison below is a pointer check that is true
   * for almost every cell on the screen and pushes nothing at all.
   */
  const refreshShapes = (mounted: MountedCell): void => {
    const width = columnWidths.get(mounted.column)?.value ?? COLUMN_WIDTH;
    const next = bordersOf(mounted.paint.value, width);
    if (next === NO_SHAPES && mounted.shapes.value === NO_SHAPES) {
      return;
    }
    mounted.shapes.next(next);
  };

  ctx.effect(sheet.view.formats, current => {
    latestFormats = current;
    repaint();
  });

  ctx.effect(sheet.view.palette, current => {
    latestPalette = current.entries;
    repaint();
  });

  /**
   * The column widths, which a drag on a header's edge changes.
   *
   * Held here rather than on the channel. A width is not the
   * application's business — the sheet's values do not depend on how
   * wide anything is drawn — and putting it on the wire would make
   * every frame of a drag a round trip through another thread to
   * decide something this one already knows. Phase 6 will persist them,
   * and that is when they become the document's.
   */
  const widths = internalState<number[]>(Array.from({ length: COLUMN_COUNT }, () => COLUMN_WIDTH));

  /**
   * The rows somebody has hidden, as heights the window understands.
   *
   * Sparse, because a sheet is ten thousand rows tall and all but a
   * handful are the default — which is the shape
   * `UiVirtualSheetOptions.rowHeights` takes, and the reason it is a
   * map rather than the array `columnWidth` is.
   */
  /**
   * How many rows and columns stay put while the rest scrolls.
   *
   * Read rather than owned: the pane is the document's, so it
   * survives a reload. Held here as well because a cell has to know
   * whether it is inside the pane to draw itself stuck, and asking
   * across the barrier per cell is not a question worth sending.
   */
  const frozen = internalState<{ rows: number; columns: number }>({ rows: 0, columns: 0 });

  /**
   * The merged rectangles, all of them.
   *
   * Held whole rather than windowed, because a merge anchored above
   * or to the left of the window still has to paint into it — which
   * is what `extendRange` below reaches back for.
   */
  const merges = internalState<readonly SheetMerge[]>([]);
  const mergeAt = (row: number, column: number): SheetMerge | null => {
    for (const rect of merges.value) {
      if (row >= rect.firstRow && row <= rect.lastRow && column >= rect.firstColumn && column <= rect.lastColumn) {
        return rect;
      }
    }
    return null;
  };

  const hidden = internalState<ReadonlySet<number>>(new Set<number>());
  const heightsOf = (rows: ReadonlySet<number>): Map<number, number> =>
    new Map([...rows].map(row => [row, 0] as const));

  /**
   * One height per row that anyone is looking at.
   *
   * The same argument as `columnWidths`, at a smaller scale: a row
   * that is hidden changes the cells in that row and nothing else, so
   * a subject per row is a list a cell can be taken off in a glance.
   */
  /**
   * Where a column starts, past the gutter.
   *
   * Summed here rather than asked of the window, because a cell is
   * built *during* the window's construction — the first `renderRow`
   * runs inside `LazySheet(...)`, before the handle it returns has
   * been assigned — so reaching for the window from a cell is a
   * reference to something that does not exist yet. Only the frozen
   * columns ask, and there are a handful of those.
   */
  const columnLeft = (column: number): number => {
    let left = 0;
    for (let at = 0; at < column; at++) {
      left += widths.value[at] ?? COLUMN_WIDTH;
    }
    return left;
  };

  const rowHeights = new Map<number, BehaviorSubject<number>>();
  const heightOf = (row: number): Observable<number> => {
    let height = rowHeights.get(row);
    if (height === undefined) {
      height = new BehaviorSubject(hidden.value.has(row) ? 0 : ROW_HEIGHT);
      rowHeights.set(row, height);
    }
    return height;
  };
  ctx.effect(hidden, rows => {
    for (const [row, height] of rowHeights) {
      const next = rows.has(row) ? 0 : ROW_HEIGHT;
      if (next !== height.value) {
        height.next(next);
      }
    }
  });

  /**
   * One width per column, rather than one array every cell reads.
   *
   * The same argument as the window and the selection, at a smaller
   * scale: a column's width changes for the thirty-odd cells in that
   * column and for nothing else, so a subject per column is a list a
   * cell can be taken off in a glance. It also makes a resize drag
   * touch one column's cells instead of the window's.
   */
  const columnWidths = new Map<number, BehaviorSubject<number>>();
  const widthOf = (column: number): Observable<number> => {
    let width = columnWidths.get(column);
    if (width === undefined) {
      width = new BehaviorSubject(widths.value[column] ?? COLUMN_WIDTH);
      columnWidths.set(column, width);
    }
    return width;
  };
  ctx.effect(widths, all => {
    const moved = new Set<number>();
    for (const [column, width] of columnWidths) {
      const next = all[column] ?? COLUMN_WIDTH;
      if (next !== width.value) {
        width.next(next);
        moved.add(column);
      }
    }
    if (moved.size === 0) {
      return;
    }
    // A right-hand border is drawn at the far side of the cell, so a
    // column that changed width has to redraw the borders in it —
    // and only in it.
    for (const mounted of cells.values()) {
      if (moved.has(mounted.column)) {
        refreshShapes(mounted);
      }
    }
  });

  /**
   * A cell's borders, as rectangles.
   *
   * **Four thin rects and not four child boxes.** A border in the
   * engine is one `borderWidth` for all four sides, so a cell cannot
   * have a heavy rule above it and a hairline below by that route.
   * What it *can* have is decorations: the `decorated` modifier takes
   * an Observable of arbitrary coloured rectangles, drawn in the
   * node's own paint pass, clipped by its ancestors and transformed
   * with it — with nothing to lay out and nothing to hit test. So the
   * cost of a bordered cell is four draw instances and no extra
   * nodes, which is the number that matters on this surface.
   *
   * The rects are drawn *inside* the cell, as the engine's own border
   * is and as CSS draws one, so a border never encroaches on the
   * neighbour and two adjacent cells can each have their own.
   */
  const bordersOf = (paint: CellPaint, width: number): readonly DecorationShape[] => {
    const edges = paint.borders;
    if (edges.top.width === 0 && edges.right.width === 0 && edges.bottom.width === 0 && edges.left.width === 0) {
      // The common case by a very long way, and it allocates nothing.
      return NO_SHAPES;
    }
    const shapes: DecorationShape[] = [];
    const edge = (e: CellEdge, box: { x: number; y: number; width: number; height: number }): void => {
      if (e.width > 0) {
        shapes.push({ kind: 'fill', ...box, radius: 0, color: e.color === '' ? 'text' : e.color, after: 'children' });
      }
    };
    edge(edges.top, { x: 0, y: 0, width, height: edges.top.width });
    edge(edges.bottom, { x: 0, y: ROW_HEIGHT - edges.bottom.width, width, height: edges.bottom.width });
    edge(edges.left, { x: 0, y: 0, width: edges.left.width, height: ROW_HEIGHT });
    edge(edges.right, { x: width - edges.right.width, y: 0, width: edges.right.width, height: ROW_HEIGHT });
    return shapes;
  };

  /**
   * The boxes drawn round the cells a formula being typed refers to.
   *
   * Kept per row and *pushed*, on the rule this file has followed
   * since Phase 0: a row subscribes to a subject of its own, and one
   * writer fills it. Piping the draft into every mounted row would
   * put every row on the draft's observer list and pay for it on
   * every keystroke, which is the shape the profile in this file's
   * header was written about.
   *
   * Drawn with `decorated`, so an outlined range costs **no extra
   * nodes at all** — four coloured rectangles in the row's own paint
   * pass, the same mechanism per-edge cell borders use.
   */
  interface RowOutline {
    readonly shapes: BehaviorSubject<readonly DecorationShape[]>;
    /**
     * The modifier list, built once per row and handed back every
     * render.
     *
     * **Hoisted, and it has to be.** A modifier's argument is compared
     * by identity and the list is static per element, so a fresh
     * `[decorated(...)]` in the render body is a fresh modifier that
     * detaches and re-attaches. Re-attaching subscribes again, a
     * `BehaviorSubject` replays to a new subscriber, replaying
     * decorates the node, decorating dirties it, and dirtying it
     * renders the row again: a loop with nothing in it that looks like
     * a loop. It froze the render worker on the first formula typed in
     * a browser, with every spec passing — the specs settle a finite
     * number of frames and never ask whether the frames stop.
     */
    readonly modifiers: readonly UiModifier<unknown>[];
  }

  const outlines = new Map<number, RowOutline>();

  const outlineFor = (row: number): RowOutline => {
    let entry = outlines.get(row);
    if (entry === undefined) {
      const shapes = new BehaviorSubject<readonly DecorationShape[]>(NO_SHAPES);
      entry = { shapes, modifiers: [decorated(shapes)] as readonly UiModifier<unknown>[] };
      outlines.set(row, entry);
    }
    return entry;
  };

  /**
   * The segments of one reference's box that fall on one row.
   *
   * A range crossing five rows is drawn by five rows, each
   * contributing the pieces that cross it: the two sides always, the
   * top only on the first row and the bottom only on the last. Done
   * this way rather than as one tall box hanging off the first row so
   * that a range reaching into the viewport from above is still
   * outlined — the same problem merges solved with `extendRange`, and
   * a cheaper answer to it.
   */
  const outlineSegments = (
    range: { start: { row: number; column: number }; end: { row: number; column: number } },
    color: string,
    row: number,
    into: DecorationShape[]
  ): void => {
    const firstRow = Math.min(range.start.row, range.end.row);
    const lastRow = Math.max(range.start.row, range.end.row);
    if (row < firstRow || row > lastRow) {
      return;
    }
    const firstColumn = Math.min(range.start.column, range.end.column);
    const lastColumn = Math.max(range.start.column, range.end.column);
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(firstColumn);
    const right = GUTTER_WIDTH + sheetWindow.offsetOf(lastColumn) + sheetWindow.widthOf(lastColumn);
    const width = Math.max(0, right - left);
    if (width === 0) {
      return;
    }
    const box = (x: number, y: number, w: number, h: number): void => {
      into.push({ kind: 'fill', x, y, width: w, height: h, radius: 0, color, after: 'children' });
    };
    box(left, 0, OUTLINE, ROW_HEIGHT);
    box(right - OUTLINE, 0, OUTLINE, ROW_HEIGHT);
    if (row === firstRow) {
      box(left, 0, width, OUTLINE);
    }
    if (row === lastRow) {
      box(left, ROW_HEIGHT - OUTLINE, width, OUTLINE);
    }
  };

  /** Which rows currently carry an outline, so they can be cleared. */
  let outlinedRows: readonly number[] = [];

  /**
   * Recomputes the outlines from the draft.
   *
   * Only the rows that have them or had them are touched, so a
   * keystroke in a formula naming two cells writes to two subjects
   * and not to every row on the screen.
   */
  const paintOutlines = (draft: string | null): void => {
    const references = draft === null ? [] : colouredReferences(draft);
    const rows = new Map<number, DecorationShape[]>();
    for (const reference of references) {
      const firstRow = Math.min(reference.range.start.row, reference.range.end.row);
      const lastRow = Math.max(reference.range.start.row, reference.range.end.row);
      for (let row = firstRow; row <= lastRow; row++) {
        let shapes = rows.get(row);
        if (shapes === undefined) {
          shapes = [];
          rows.set(row, shapes);
        }
        outlineSegments(reference.range, reference.color, row, shapes);
      }
    }
    for (const row of outlinedRows) {
      if (!rows.has(row)) {
        outlineFor(row).shapes.next(NO_SHAPES);
      }
    }
    for (const [row, shapes] of rows) {
      outlineFor(row).shapes.next(shapes);
    }
    outlinedRows = [...rows.keys()];
  };

  ctx.effect(edit.draft, paintOutlines);

  const buildCell = (
    row: number,
    column: number,
    value: Observable<string | null>,
    state: Observable<Standing>,
    paint: Observable<CellPaint>,
    shapes: Observable<readonly DecorationShape[]>
  ): UiElement => {
    /**
     * A frozen column's cells are stuck at their own offset.
     *
     * The same `position: 'sticky'` the gutter has always used, one
     * column further along. Read once, at build time, because the
     * cells are rebuilt when the pane moves — see the effect that
     * empties `cells` on a freeze.
     */
    const stuck = column < frozen.value.columns;
    /**
     * A merged cell is drawn by its anchor and by nothing else.
     *
     * The anchor is as wide as the columns it covers and as tall as
     * the rows, and it overflows its own row downwards to reach them
     * — rows are not merged, only cells are, so there is nothing else
     * for a vertical merge to be. The cells it covers are given no
     * width and no height at all, which keeps every other cell in the
     * row at the offset the window put it and is the only version
     * that does not need the window to know about merges.
     */
    const merge = mergeAt(row, column);
    const anchors = merge !== null && merge.firstRow === row && merge.firstColumn === column;
    const covered = merge !== null && !anchors;
    const spanWidth = merge === null ? null : columnLeft(merge.lastColumn + 1) - columnLeft(merge.firstColumn);
    const spanRows = merge === null ? 1 : merge.lastRow - merge.firstRow + 1;

    return Text({
      key: column,
      position: stuck ? 'sticky' : undefined,
      left: stuck ? GUTTER_WIDTH + columnLeft(column) : undefined,
      zIndex: stuck ? 1 : anchors && spanRows > 1 ? 1 : undefined,
      /**
       * Borders follow the paint *and* the column's width, because a
       * right edge is drawn at the far side of a cell and a drag
       * moves it. Both are the cell's own subjects, so this is one
       * more subscriber on each and not one on anything shared.
       */
      modifiers: [decorated(shapes)],
      text: value.pipe(map(text => text ?? '')),
      // A cell the application worker has not sent yet is drawn as a
      // rule rather than left blank, so a gap on a fling reads as
      // "not here yet" instead of as the end of the sheet.
      color: combineLatest([value, paint]).pipe(
        map(([text, how]) => (text === null ? 'placeholder' : how.color === '' ? 'text' : how.color))
      ),
      /**
       * A fill wins everywhere except inside a selected range.
       *
       * The active cell keeps its fill — the ring is what marks it,
       * and washing it out would hide the colour somebody is in the
       * middle of choosing. The rest of a multi-cell selection takes
       * the wash, because a rectangle you cannot see is not a
       * selection.
       */
      backgroundColor: combineLatest([state, paint]).pipe(
        map(([where, how]) => (where === 1 ? SELECTED_WASH : how.fill === '' ? 'background' : how.fill))
      ),
      borderColor: state.pipe(map(where => (where === 2 ? 'primary' : GRID_LINE))),
      borderWidth: state.pipe(map(where => (where === 2 ? 2 : 1))),
      /**
       * A covered cell gives up its width only to the anchor's own row.
       *
       * The anchor takes the whole span across the row it is in, so the
       * cells beside it there must come to nothing or the row is twice
       * as wide as it should be. A row *below* the anchor has no anchor
       * in it — the anchor reaches down into it by overflowing — so the
       * cells it covers there still have to hold their columns open.
       * Zeroed, they dragged every column after them one place left,
       * which is invisible until there is something in those columns
       * to see move. Found by reading the boxes in a browser.
       */
      width: covered
        ? merge !== null && row === merge.firstRow
          ? 0
          : widthOf(column)
        : anchors && spanWidth !== null
          ? spanWidth
          : widthOf(column),
      height: covered
        ? 0
        : anchors && spanRows > 1
          ? heightOf(row).pipe(map(height => (height === 0 ? 0 : height * spanRows)))
          : heightOf(row),
      flexShrink: 0,
      // A covered cell has no room for padding either, or a row of
      // them adds twelve pixels each to the width of the row. One
      // below the anchor keeps its width but still draws nothing, so
      // padding it would only push against a zero height.
      paddingLeft: covered ? 0 : CELL_PADDING,
      paddingRight: covered ? 0 : CELL_PADDING,
      fontSize: paint.pipe(map(how => (how.fontSize === 0 ? CELL_FONT_SIZE : how.fontSize))),
      /**
       * Wired, and not yet visible.
       *
       * `LazySheet` takes one `rowHeight` for every row — a number
       * and not an array, where `columnWidth` is already either — so
       * a wrapped cell has nowhere to put its second line and shows
       * the first, which is what clipping shows too. The property is
       * correct here so that the day row heights vary, this does not
       * have to be found and fixed; the control that set it is out of
       * the toolbar until then, on the rule that a control which
       * silently does nothing is worse than one that is missing.
       */
      textWrap: paint.pipe(map(how => (how.wrap ? 'word' : 'none'))),
      fontWeight: paint.pipe(map(how => (how.bold ? 'bold' : 'normal'))),
      fontStyle: paint.pipe(map(how => (how.italic ? 'italic' : 'normal'))),
      textDecoration: paint.pipe(map(how => (how.underline ? 'underline' : 'none'))),
      textOverflow: 'clip',
      verticalAlign: 'middle',
      /**
       * `auto` is the spreadsheet rule — numbers right, text left —
       * and it is a real alignment rather than an absent one: a
       * column of numbers nobody has touched still has to line up.
       */
      textAlign: combineLatest([value, paint]).pipe(
        map(([text, how]) =>
          how.align === 'auto' ? (isNumeric(text) ? 'right' : 'start') : how.align === 'center' ? 'center' : how.align
        )
      ),
      // A sweep drags a text selection through anything selectable,
      // so the numbers and the row labels would highlight as prose
      // does while a rectangle is being picked out.
      selectable: false,
      role: 'cell',
      onClick: (event: UiPointerEvent) => selectByPointer(row, column, event.modifiers.shift)
    });
  };

  /** The cell the fill handle hangs off: the selection's far corner. */
  let cornerAt: { row: number; column: number } | null = null;

  /**
   * A click on a cell.
   *
   * Three things, and the third is the one that was missing: the
   * keyboard has to end up somewhere. The grid is what holds focus —
   * a spreadsheet does not put focus on a cell, it puts focus on the
   * sheet and moves a selection inside it — so a click that moved the
   * selection and left focus where it was gave you a selected cell
   * that no key did anything to.
   */
  /**
   * A cell as the selection should name it.
   *
   * A merge is its anchor: the covered cells hold nothing and are not
   * drawn, so a selection sitting on one would be a ring around
   * nothing and a formula bar showing a cell nobody can see. Clicks
   * land on the anchor already, because the anchor is the node that
   * spans those pixels — it is the *arithmetic* paths that need this,
   * `cellAt` being a division over the offsets that has never heard
   * of a merge.
   */
  const anchorOf = (row: number, column: number): { row: number; column: number } => {
    const merge = mergeAt(row, column);
    return merge === null ? { row, column } : { row: merge.firstRow, column: merge.firstColumn };
  };

  /**
   * Where a reference being picked was written, so a drag can rewrite
   * it in place rather than adding a corner per pointer move.
   */
  let picking: Span | null = null;
  /**
   * The text the last pick wrote, so typing can cancel the pick.
   *
   * Clicking a second cell straight after picking one *replaces* it —
   * you are still choosing which cell you meant. But typing anything
   * in between ends that: `=A1` then `+` then a click on B2 has to
   * give `=A1+B2`, and without this it gives `=B2+`, which is the
   * first reference silently eaten.
   */
  let pickedText: string | null = null;
  /** The corner a range is being dragged from, while one is. */
  let pickingFrom: { row: number; column: number } | null = null;

  ctx.effect(edit.draft, draft => {
    if (draft !== pickedText) {
      picking = null;
      pickedText = null;
    }
  });

  /**
   * A click while a formula is open, when it means "this cell".
   *
   * **This is the mode the roadmap called the hard part**, and the
   * decision is not made here: `pickDecision` is a function of the
   * text and the caret with a table of cases beside it, and this asks
   * it and does what it says. A pointer handler that decided for
   * itself is how a click during an edit ends up throwing the formula
   * away — or how the selection freezes because every click is read
   * as picking.
   *
   * Returns whether the click was a pick. False means it was a click.
   */
  const pickByPointer = (row: number, column: number, to?: { row: number; column: number }): boolean => {
    if (!edit.openNow() || editorNode === null) {
      return false;
    }
    const draft = edit.draftNow();
    if (draft === null) {
      return false;
    }
    const model = editorFor(editorNode);
    const range = {
      start: relativeRef(row, column),
      end: relativeRef(to?.row ?? row, to?.column ?? column)
    };
    // A drag rewrites the address it already wrote; the first press
    // asks the caret where it is.
    const picked = picking === null ? pick(draft, model.focus, range) : repick(draft, picking, range);
    if (picked === null) {
      return false;
    }
    picking = picked.span;
    // Before the write, because the write feeds the effect above
    // synchronously and it has to recognise its own text.
    pickedText = picked.text;
    edit.write(picked.text);
    // The model is the text's owner while the cell is open, so the
    // caret has to be put back by hand: `write` changes the draft and
    // the field follows it, and neither of them knows where somebody
    // was typing.
    model.replaceText(picked.text);
    model.select(picked.caret);
    focus.focus(editorNode);
    return true;
  };

  const selectByPointer = (row: number, column: number, extend = false): void => {
    // A click ends any sweep, whatever the gesture recogniser thinks.
    //
    // Without this a click after a sweep extended the rectangle
    // instead of putting it down: the sweep's flag was still set, the
    // next pointer movement read as more of the same drag, and the
    // anchor stayed where the sweep had started. The click is the
    // later and more definite statement of what the person wants, so
    // it wins.
    sweeping = false;
    // A click that is picking a reference into a formula is not a
    // click on a cell, and must not commit what is open.
    if (!extend && pickByPointer(row, column)) {
      return;
    }
    // A click elsewhere commits what is open, as it does everywhere.
    if (edit.openNow()) {
      edit.commit(0, 0);
    }
    const at = anchorOf(row, column);
    if (extend) {
      edit.extendTo(at.row, at.column);
    } else {
      edit.moveTo(at.row, at.column);
    }
    if (gridNode !== null) {
      focus.focus(gridNode);
    }
  };

  const cell = (row: number, column: number): UiElement => {
    if (openAt !== null && openAt.row === row && openAt.column === column) {
      return editorCell(row, column);
    }
    const key = `${row}:${column}`;
    const mounted = cells.get(key);
    if (mounted !== undefined) {
      return mounted.element;
    }
    // Seeded from what the window and the selection say *now*, because
    // a cell is built during the frame that reveals it and the feeds
    // above have already run for this one.
    const value = values.for(key, { row, column });
    const standing = standings.for(key, { row, column });
    const paint = new BehaviorSubject<CellPaint>(paintOf(row, column));
    const shapes = new BehaviorSubject<readonly DecorationShape[]>(
      bordersOf(paint.value, columnWidths.get(column)?.value ?? widths.value[column] ?? COLUMN_WIDTH)
    );
    const element = buildCell(row, column, value, standing, paint, shapes);
    cells.set(key, { row, column, element, value, standing, paint, shapes });
    return element;
  };

  /**
   * The cell being typed into.
   *
   * An `EditableText` in the cell's own place rather than a box
   * floating over it: the caret, the selection, IME composition and
   * the clipboard are all `EditableText`'s already, and a cell that
   * *is* the editor cannot drift away from the cell it is editing when
   * the sheet scrolls or a column is dragged.
   *
   * Not memoised. There is one of these at a time and its life is the
   * edit.
   */
  const editorCell = (row: number, column: number): UiElement => {
    // The open cell is always a merge's anchor — `anchorOf` moves the
    // selection there before anything can open it — so this needs the
    // anchor's arithmetic and not the covered case.
    const merge = mergeAt(row, column);
    const spanWidth = merge === null ? null : columnLeft(merge.lastColumn + 1) - columnLeft(merge.firstColumn);
    const spanRows = merge === null ? 1 : merge.lastRow - merge.firstRow + 1;

    return EditableText({
      key: column,
      ref: node => {
        editorNode = node;
        if (node === null) {
          // The open cell scrolled out of the window and its node went
          // with it. The draft survives — it was never in the node —
          // but focus would be on nothing, and the next key would go
          // nowhere at all. Handing it to the grid keeps Enter and
          // Escape working, and the formula bar goes on showing what
          // is being typed because it is the same buffer.
          if (edit.openNow() && gridNode !== null) {
            focus.focus(gridNode);
          }
          return;
        }
        {
          // The caret goes to the end: where it is after typing the
          // character that opened the cell, and where a spreadsheet
          // leaves it on F2. `replaceText` deliberately keeps the
          // caret where it still fits — there is a spec for that — so
          // placing it is the caller's job, and the model hangs on the
          // node for exactly this.
          //
          // Seeded from the draft rather than from the node's `value`,
          // because a ref fires as the node is created and the
          // property may not have been written yet: read the other way
          // round this put the caret at the end of an empty string and
          // every character typed afterwards landed in front of what
          // was already there.
          const model = editorFor(node);
          const text = edit.draftNow() ?? '';
          if (model.text !== text) {
            model.replaceText(text);
          }
          model.select(text.length);
          // Focus follows the edit into the cell — unless the person
          // put the caret in the formula bar, in which case taking it
          // away would bounce them into the cell they chose not to
          // type in. Done here rather than from the draft, because the
          // node exists at exactly this moment and not before it.
          if (edit.startedIn() === 'grid') {
            focus.focus(node);
          }
        }
      },
      value: edit.draft.pipe(map(text => text ?? '')),
      /**
       * The references, in the colours their boxes are drawn in.
       *
       * Derived from the same draft the value is, one `map` further
       * along, so the runs and the text can never be a frame apart —
       * which matters because the engine checks they describe the same
       * string and draws plainly when they do not. A cell holding
       * anything that is not a formula gets `undefined` and costs
       * nothing.
       */
      spans: edit.draft.pipe(map(text => formulaSpans(text ?? ''))),
      /**
       * The editor is the merged cell, not the cell under its corner.
       *
       * This is a cell in the row like any other, so its width is what
       * holds the columns after it in place. Built from `widthOf`
       * alone, opening a merge collapsed the anchor to one column and
       * slid the rest of the row left by the difference — which on
       * screen looks exactly like the merge coming apart to show the
       * original cells underneath it. Same arithmetic as `buildCell`
       * above, and for the same reason.
       */
      width: spanWidth ?? widthOf(column),
      // A tall editor reaches down out of its own row, as the anchor
      // does: rows are not merged, only cells are. It already carries
      // the `zIndex: 1` below that lets it, being the open cell.
      height: spanRows > 1 ? heightOf(row).pipe(map(height => (height === 0 ? 0 : height * spanRows))) : heightOf(row),
      flexShrink: 0,
      paddingLeft: 6,
      paddingRight: 6,
      fontSize: 12,
      textWrap: 'none',
      verticalAlign: 'middle',
      backgroundColor: 'background',
      color: 'text',
      borderColor: 'primary',
      borderWidth: 2,
      zIndex: 1,
      role: 'textbox',
      label: 'Cell',
      onInput: (event: UiTextChangeEvent) => edit.write(event.value),
      onKeyDown: onKey
    });
  };

  /** The node holding the open cell's editor, so focus can be put in it. */
  let editorNode: UiNode | null = null;
  /** The cell the renderer should build as an editor, read while building. */
  let openAt: { row: number; column: number } | null = null;

  /** The frozen strip at the start of a row: the row's number. */
  const rowHeader = (row: number): UiElement =>
    Text({
      key: 'gutter',
      text: String(row + 1),
      width: GUTTER_WIDTH,
      height: heightOf(row),
      flexShrink: 0,
      // Held at the left edge while the sheet scrolls sideways. Its
      // vertical travel is its row's, which it gets for free by being
      // inside it.
      position: 'sticky',
      left: 0,
      zIndex: 1,
      backgroundColor: 'surface',
      borderColor: GRID_LINE,
      borderWidth: 1,
      color: 'textMuted',
      fontSize: 11,
      textAlign: 'center',
      verticalAlign: 'middle',
      selectable: false,
      role: 'rowheader'
    });

  /**
   * Every key, whether it arrived at the grid or at the open cell.
   *
   * One handler for both because the table in `SheetKeys` is one
   * table: Enter means something different inside a cell and the
   * difference is a parameter, not a second code path. A key the table
   * has no meaning for is left alone — that is what lets an arrow key
   * move the caret rather than the selection while a cell is open.
   */
  const onKey = (event: UiKeyboardEvent): void => {
    /**
     * Accelerators first, and only with no cell open.
     *
     * First, because `commandFor` and `keyAction` are two tables and
     * the one that answers has to be decided somewhere rather than by
     * which `if` was written above the other — `SheetCommands.spec`
     * is what guarantees they never both answer the same key.
     *
     * And only with no cell open, because a cell being typed into is
     * a text field: every key in it belongs to the text, which is the
     * same rule the `editing` argument below encodes.
     */
    if (!edit.openNow()) {
      // Escape closes whatever the chrome has open, before anything
      // else looks at it. With a cell open it belongs to the cell —
      // it is what puts back what was there — which is why this is
      // inside the same guard the accelerators are.
      if (event.key === 'Escape' && edit.dismiss()) {
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      const command = commandFor(event.key, event.modifiers);
      if (command !== null) {
        event.preventDefault();
        event.stopPropagation();
        edit.runCommand(command);
        return;
      }
    }
    const action = keyAction(event.key, event.modifiers, edit.openNow());
    if (edit.apply(action)) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  /**
   * The fill handle: the small square on the selection's bottom-right
   * corner that extends it.
   *
   * A child of the row it is on rather than a thing floating over the
   * grid, so it travels with the selection, survives a scroll and
   * needs nothing kept in step. The row is `position: 'relative'`
   * because an absolute child is placed against its nearest positioned
   * ancestor, and without one it would walk up to the layout root and
   * be drawn in the corner of the screen.
   */
  const fillHandle = (column: number): UiElement =>
    Box({
      key: 'fill',
      width: 8,
      height: 8,
      position: 'absolute',
      // Placed from the window's own offsets rather than by hanging
      // off the cell: a `Text` takes no children, and giving the one
      // selected cell a different element type would cost it its place
      // in the cache and its role in the semantics tree.
      left: GUTTER_WIDTH + sheetWindow.offsetOf(column) + sheetWindow.widthOf(column) - 5,
      top: ROW_HEIGHT - 5,
      zIndex: 3,
      backgroundColor: 'primary',
      borderColor: 'background',
      borderWidth: 1,
      cursor: 'crosshair',
      role: 'button',
      label: 'Fill',
      onPanStart: (event: UiPointerEvent) => {
        filling = true;
        // The grid listens for pans as well, to sweep a selection.
        // Without this a drag on the handle would do both.
        event.stopPropagation();
      },
      onPanMove: (event: UiPointerEvent) => {
        if (!filling) {
          return;
        }
        event.stopPropagation();
        // Where the pointer is, in cells. The window owns the offsets,
        // so this is arithmetic rather than a hit test — which matters
        // because the cell being dragged towards is usually one that
        // has not been mounted yet.
        const box = viewport.value;
        fillTo = sheetWindow.cellAt(event.x - box.x, event.y - box.y);
      },
      onPanEnd: (event: UiPointerEvent) => {
        filling = false;
        event.stopPropagation();
        if (fillTo !== null) {
          sheet.send.fill(fillTo.row, fillTo.column);
          fillTo = null;
        }
      }
    });

  let filling = false;
  let fillTo: { row: number; column: number } | null = null;
  let sweeping = false;

  /**
   * The cell under a pointer event, or null before the first layout.
   *
   * The frozen strips are content the window already accounts for, so
   * a sweep that wanders over the row numbers reads as the first
   * column rather than as nothing — which is what dragging into the
   * gutter should do.
   */
  const cellUnder = (event: UiPointerEvent): { row: number; column: number } | null => {
    const box = viewport.value;
    if (box.width === 0) {
      return null;
    }
    const at = sheetWindow.cellAt(event.x - box.x, event.y - box.y);
    // `cellAt` is arithmetic on the offsets and knows nothing about
    // merges, so a sweep across one reports the cells underneath it —
    // which are not drawn and hold nothing. The merge is what is
    // there, and the merge is its anchor.
    return anchorOf(at.row, at.column);
  };

  const renderRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [rowHeader(row)];
    // The frozen columns first, which is the order the window's own
    // leading spacer is placed against: everything that stays put,
    // then the gap, then the columns that scrolled into view.
    for (let column = 0; column < frozen.value.columns; column++) {
      line.push(cell(row, column));
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(cell(row, column));
    }
    const corner = cornerAt !== null && cornerAt.row === row;
    if (corner && cornerAt !== null) {
      line.push(fillHandle(cornerAt.column));
    }
    /**
     * The row clips its own cells, which is what makes a hidden row
     * disappear rather than merely collapse.
     *
     * A hidden row is one of height zero, and a zero-height row whose
     * cells are also zero-height still *paints* them — nothing in the
     * engine clips a node to its box unless it is asked to, so the
     * text of the hidden row went on drawing over its neighbours.
     * Found by hiding a row in a browser; every spec passed, because
     * they asked what the cell's height property was and a height of
     * zero was exactly what they got.
     *
     * Only the hidden rows are clipped. Clipping every row would cut
     * off the fill handle, which deliberately hangs outside the
     * corner cell.
     */
    /**
     * A frozen row is stuck under the header, at its own offset.
     *
     * The engine mounts it; making it *stay* is this line, which is
     * the same `position: 'sticky'` the header row and the gutter
     * have always used. `zIndex` puts it over the rows it covers and
     * under the header that covers it.
     */
    const stuck = row < frozen.value.rows;
    /**
     * A row holding the anchor of a vertical merge is lifted.
     *
     * The anchor is taller than its row and reaches down over the
     * rows below, which are painted after it — so without this the
     * merge is drawn and then covered up by the very cells it is
     * supposed to be hiding.
     */
    const spans = merges.value.some(rect => rect.firstRow === row && rect.lastRow > row);
    return Row(
      {
        role: 'row',
        posInSet: row + 1,
        // Four coloured rectangles in the row's own paint pass when a
        // formula being typed names something on this row, and an
        // empty array the rest of the time.
        modifiers: outlineFor(row).modifiers,
        position: stuck ? 'sticky' : corner || spans ? 'relative' : undefined,
        top: stuck ? HEADER_HEIGHT + row * ROW_HEIGHT : undefined,
        zIndex: stuck || spans ? 1 : undefined,
        backgroundColor: stuck ? 'background' : undefined,
        overflow: heightOf(row).pipe(map(height => (height === 0 ? 'hidden' : undefined)))
      },
      ...line
    );
  };

  const renderHeader = (firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [
      // The corner is sticky on both axes: its own `left` holds it
      // against horizontal scroll, and it inherits the header row's
      // `top` by being inside it.
      Text({
        key: 'corner',
        text: '',
        width: GUTTER_WIDTH,
        height: HEADER_HEIGHT,
        flexShrink: 0,
        position: 'sticky',
        left: 0,
        zIndex: 3,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1
      })
    ];
    // The frozen columns first, in the same order a row puts them,
    // so a column label sits over the column it labels whichever of
    // the two is stuck.
    for (let column = 0; column < frozen.value.columns; column++) {
      line.push(columnHeader(column));
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(columnHeader(column));
    }
    // Held at the top while the rows scroll under it, and above them
    // in paint order — it is the first child, so without a zIndex the
    // rows would be drawn over it.
    return Row({ role: 'row', position: 'sticky', top: 0, zIndex: 2 }, ...line);
  };

  /**
   * A column's label, with the grip that resizes it.
   *
   * The grip is a child of the header rather than a strip of its own,
   * so it travels with the column and there is nothing to keep in
   * step. It is eight pixels wide against a one-pixel rule, because a
   * one-pixel target is the scrollbar mistake again — Phase 3 of the
   * engine widened a six-pixel thumb for exactly this reason.
   */
  const columnHeader = (column: number): UiElement => {
    const stuck = column < frozen.value.columns;
    return Box(
      {
        key: column,
        width: widthOf(column),
        height: HEADER_HEIGHT,
        flexShrink: 0,
        x: 'center',
        y: 'center',
        // The grip is absolute, and an absolute child is placed
        // against its nearest *positioned* ancestor — without this it
        // would walk past the header to the layout root and be drawn
        // in the corner of the screen. A frozen column's label is
        // stuck instead, which positions it just the same.
        position: stuck ? 'sticky' : 'relative',
        left: stuck ? GUTTER_WIDTH + columnLeft(column) : undefined,
        zIndex: stuck ? 1 : undefined,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1,
        role: 'columnheader',
        label: columnName(column),
        posInSet: column + 1
      },
      Text({
        text: columnName(column),
        color: 'textMuted',
        fontSize: 11,
        fontWeight: 600,
        verticalAlign: 'middle',
        selectable: false
      }),
      Box({
        key: 'grip',
        width: 8,
        height: HEADER_HEIGHT,
        position: 'absolute',
        right: -4,
        top: 0,
        zIndex: 4,
        cursor: 'col-resize',
        // Pan, not Drag. Gesso's Drag is the long-press-to-pick-up
        // one, which is right for moving a thing and wrong for
        // dragging an edge: nobody holds still on a resize handle
        // before pulling it. Pan claims the press as soon as the
        // pointer passes the slop, which is what a grip wants.
        onPanStart: event => {
          resizing = { column, width: widths.value[column] ?? COLUMN_WIDTH, from: event.x };
        },
        onPanMove: event => {
          if (resizing !== null) {
            resizeTo(resizing.column, resizing.width + (event.x - resizing.from));
          }
        },
        onPanEnd: () => {
          if (resizing !== null) {
            // The end of the drag, and the only moment the other
            // thread hears about it. Sending each frame would be a
            // round trip per pixel to agree on something this side has
            // already drawn; sending the result is what makes the
            // width survive a reload.
            sheet.send.setColumnWidth(resizing.column, widths.value[resizing.column] ?? COLUMN_WIDTH);
          }
          resizing = null;
        }
      })
    );
  };

  let resizing: { column: number; width: number; from: number } | null = null;

  /** A column may be dragged narrow, but not to nothing. */
  const resizeTo = (column: number, width: number): void => {
    const next = widths.value.slice();
    next[column] = Math.max(MIN_COLUMN_WIDTH, Math.round(width));
    widths.value = next;
  };

  /** Written only when the selection needs bringing into view. */
  const scrollX = internalState(0);
  const scrollY = internalState(0);
  let gridNode: UiNode | null = null;
  // The chrome — menus, the find bar, the name box — has to hand the
  // keyboard back when it is done, and the grid's node is the grid's.
  edit.provideFocus(() => {
    if (gridNode !== null) {
      focus.focus(gridNode);
    }
  });
  // How big the viewport is, which is what "already visible" is
  // measured against. Taken from the node rather than tracked here:
  // the window is a flex child and nothing on this side knows its
  // height until the frame that laid it out.
  const viewport = ctx.bounds('sheet');

  let window: UiVirtualSheet | undefined;
  const grid = LazySheet(
    {
      flex: 1,
      minHeight: 0,
      width: percent(100),
      backgroundColor: 'background',
      // The extent is the application worker's, published on
      // `geometry`; the window follows it rather than agreeing with it.
      rowCount: sheet.view.geometry.pipe(map(g => g.rowCount)),
      columnCount: sheet.view.geometry.pipe(map(g => g.columnCount)),
      rowHeight: ROW_HEIGHT,
      rowHeights: heightsOf(hidden.value),
      columnWidth: widths.value,
      frozenRows: frozen.value.rows,
      frozenColumns: frozen.value.columns,
      /**
       * The window, reaching back to the anchors it is cutting
       * through.
       *
       * A merge is drawn by its top-left cell, so a window that
       * starts past that cell has nothing to draw and the merge
       * disappears at the edge of the screen. This is the one place a
       * sheet's geometry depends on its contents, and the engine
       * takes it as a hook for exactly this reason.
       */
      extendRange: (range: SheetRange) => {
        let firstRow = range.firstRow;
        let firstColumn = range.firstColumn;
        for (const rect of merges.value) {
          const inRows = rect.lastRow >= range.firstRow && rect.firstRow <= range.lastRow;
          const inColumns = rect.lastColumn >= range.firstColumn && rect.firstColumn <= range.lastColumn;
          if (inRows && inColumns) {
            firstRow = Math.min(firstRow, rect.firstRow);
            firstColumn = Math.min(firstColumn, rect.firstColumn);
          }
        }
        return { ...range, firstRow, firstColumn };
      },
      gutterWidth: GUTTER_WIDTH,
      headerHeight: HEADER_HEIGHT,
      // Two rows and one column: the partial cells at the edges. The
      // lookahead is the application worker's; see PHASE0.md §3.
      rowOverscan: 2,
      columnOverscan: 1,
      role: 'grid',
      label: 'Sheet',
      focusable: true,
      ref: node => (gridNode = node),
      modifiers: [viewport.modifier],
      onKeyDown: onKey,
      // Sweeping a selection out with the pointer.
      //
      // On the grid rather than on every cell: three listeners instead
      // of three per cell, and the cell being swept over is found from
      // the window's offsets, which works for the ones that are not
      // mounted as well as the ones that are. A cell would have to be
      // under the pointer to hear about it, and past the edge of the
      // viewport none is.
      onPanStart: (event: UiPointerEvent) => {
        const at = cellUnder(event);
        if (at === null) {
          return;
        }
        // A drag that starts where a reference may go is dragging a
        // range *into the formula*, not sweeping a selection. The
        // corner is remembered so every move can rewrite the address
        // rather than add another one.
        if (pickByPointer(at.row, at.column)) {
          pickingFrom = at;
          return;
        }
        sweeping = true;
        if (edit.openNow()) {
          edit.commit(0, 0);
        }
        edit.moveTo(at.row, at.column);
        if (gridNode !== null) {
          focus.focus(gridNode);
        }
      },
      onPanMove: (event: UiPointerEvent) => {
        if (pickingFrom !== null) {
          if (event.buttons === 0) {
            return;
          }
          const at = cellUnder(event);
          if (at !== null) {
            pickByPointer(pickingFrom.row, pickingFrom.column, at);
          }
          return;
        }
        // The button has to still be down. A recogniser that reported
        // a move after the release would otherwise go on stretching
        // the selection under a pointer that is merely passing over.
        if (!sweeping || event.buttons === 0) {
          return;
        }
        const at = cellUnder(event);
        if (at !== null) {
          edit.extendTo(at.row, at.column);
        }
      },
      onPanEnd: () => {
        sweeping = false;
        pickingFrom = null;
      },
      // Text from the clipboard with no caret anywhere. Before the
      // engine offered this the paste was dropped: `paste` had nothing
      // editable to insert into and returned false.
      onPaste: (event: UiPasteEvent) => {
        edit.pasteText(event.text);
        event.preventDefault();
      },
      scrollX,
      scrollY,
      header: renderHeader,
      sheetRef: found => (window = found)
    },
    renderRow
  );
  if (window === undefined) {
    throw new Error('LazySheet did not hand back its window.');
  }
  const sheetWindow = window;

  // The window owns the offsets, so a new set of widths has to reach
  // it: everything past the column that moved sits somewhere else, and
  // the prefix sum is what says where.
  ctx.effect(widths, all => sheetWindow.setColumnWidths(all));

  // And the same for the rows, which the window needs before it can
  // place anything below a hidden one.
  ctx.effect(hidden, rows => sheetWindow.setRowHeights(heightsOf(rows)));

  /**
   * A pane that moved, reaching the window and the cells.
   *
   * Every mounted cell is dropped, because whether a cell is stuck is
   * decided when it is built and a cell that used to be frozen is now
   * an ordinary one. Freezing is rare and a full rebuild is the
   * honest price; the alternative is a binding per cell for something
   * that changes once a session.
   */
  ctx.effect(frozen, pane => {
    cells.clear();
    values.releaseAll();
    standings.releaseAll();
    sheetWindow.setFrozen(pane.rows, pane.columns);
  });

  /**
   * The geometry the application worker publishes, adopted.
   *
   * The widths *lead* on this side while a drag is happening — that
   * is Phase 3's decision and it is why a resize costs no round trip
   * — but they are the document's once they are written down, and a
   * sheet reopened had been showing default widths while the worker
   * held the saved ones. Skipped mid-drag, or the echo of what is
   * being dragged would fight the drag.
   */
  ctx.effect(sheet.view.geometry, geometry => {
    if (resizing === null && !sameWidths(geometry.columnWidths, widths.value)) {
      widths.value = [...geometry.columnWidths];
    }
    const rows = geometry.hiddenRows;
    if (rows.length !== hidden.value.size || rows.some(row => !hidden.value.has(row))) {
      hidden.value = new Set(rows);
    }
    if (geometry.frozenRows !== frozen.value.rows || geometry.frozenColumns !== frozen.value.columns) {
      frozen.value = { rows: geometry.frozenRows, columns: geometry.frozenColumns };
    }
    if (geometry.merges !== merges.value) {
      merges.value = geometry.merges;
      // A merge changes which cells are drawn and how wide, and that
      // is decided when a cell is built.
      cells.clear();
    values.releaseAll();
    standings.releaseAll();
      sheetWindow.invalidate();
    }
  });

  /**
   * Autofit: the answer to a question this thread asked.
   *
   * The application worker sent the longest strings in each column
   * and knows nothing about fonts; this thread knows the font and
   * holds thirty rows. Measuring the candidates here is the only
   * place both halves exist, and `TextService` measures through the
   * *same* measurer the layout engine uses — so the width is the
   * width the cells will actually be laid out at rather than a second
   * opinion about it.
   */
  let fitted = 0;
  ctx.effect(sheet.view.autofit, autofit => {
    if (autofit.serial <= fitted || !measure.ready) {
      return;
    }
    fitted = autofit.serial;
    const next = widths.value.slice();
    let moved = false;
    for (const entry of autofit.columns) {
      let widest = 0;
      entry.samples.forEach((text, index) => {
        widest = Math.max(
          widest,
          measure.widthOf(text, { fontSize: CELL_FONT_SIZE, fontWeight: entry.bold[index] ? 'bold' : 'normal' })
        );
      });
      // The padding a cell draws with, plus a hair so the widest
      // string is not flush against the gridline.
      const wanted = widest === 0 ? COLUMN_WIDTH : Math.ceil(widest) + CELL_PADDING * 2 + 2;
      const clamped = Math.min(Math.max(wanted, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH);
      if (next[entry.column] !== clamped) {
        next[entry.column] = clamped;
        moved = true;
      }
    }
    if (!moved) {
      return;
    }
    widths.value = next;
    // The document is what keeps a width, so it has to hear about
    // this one the same way a drag tells it.
    for (const entry of autofit.columns) {
      sheet.send.setColumnWidth(entry.column, next[entry.column] ?? COLUMN_WIDTH);
    }
  });

  // A copy asked for on this thread is answered on the other and comes
  // back as a patch, because a command has no return value — and the
  // shell is the only thing with a clipboard to put it on.
  let copied = 0;
  ctx.effect(sheet.view.clipboard, clipboard => {
    if (clipboard.serial > copied) {
      copied = clipboard.serial;
      shell.copyText(clipboard.text);
    }
  });

  // The round trip: the range the window settled on is what the
  // application worker is asked for. `range$` emits only when the
  // range changes, so this is not a command per frame.
  ctx.effect(sheetWindow.range$, range =>
    sheet.send.setViewport(range.firstRow, range.lastRow, range.firstColumn, range.lastColumn)
  );

  /**
   * Opening and closing a cell, which is the one thing that changes
   * what a cell *is* rather than what it says.
   *
   * A `Text` cannot become an `EditableText` by having a property
   * written, so the two cells involved are dropped from the cache and
   * the window is asked to rebuild its rows from what is left. That
   * costs a rebuild of the row elements — a few hundred plain objects,
   * with every cell but these two still cached — and it happens twice
   * per cell edited rather than per frame.
   */
  /**
   * The fill handle moves with the selection, and a cell cannot grow
   * one by having a property written, so the two cells involved are
   * dropped from the cache and the rows rebuilt from what is left —
   * the same trick opening a cell uses, at the same cost, and at the
   * pace a person moves a selection rather than per frame.
   */
  let cornerBefore: { row: number; column: number } | null = null;
  ctx.effect(edit.selection, at => {
    const next = { row: Math.max(at.row, at.anchorRow), column: Math.max(at.column, at.anchorColumn) };
    if (sameCell(next, cornerBefore)) {
      return;
    }
    for (const which of [cornerBefore, next]) {
      if (which !== null) {
        cells.delete(`${which.row}:${which.column}`);
        values.release(`${which.row}:${which.column}`);
        standings.release(`${which.row}:${which.column}`);
      }
    }
    cornerBefore = next;
    cornerAt = next;
    sheetWindow.invalidate();
  });

  let openBefore: { row: number; column: number } | null = null;
  ctx.effect(combineLatest([edit.selection, edit.open]), ([at, isOpen]) => {
    const next = isOpen ? { row: at.row, column: at.column } : null;
    if (sameCell(next, openBefore)) {
      return;
    }
    for (const which of [openBefore, next]) {
      if (which !== null) {
        cells.delete(`${which.row}:${which.column}`);
        values.release(`${which.row}:${which.column}`);
        standings.release(`${which.row}:${which.column}`);
      }
    }
    openBefore = next;
    openAt = next;
    sheetWindow.invalidate();
  });

  // Closing an edit hands the keyboard back to the grid, wherever the
  // edit was being typed; opening one is the editor's own ref, which
  // is the only moment its node is known to exist.
  ctx.effect(edit.open, isOpen => {
    if (!isOpen && gridNode !== null) {
      focus.focus(gridNode);
    }
  });

  /**
   * Brings the selection into view.
   *
   * The window knows where every row and column is — that is what the
   * offsets are for — so this is arithmetic rather than a search for a
   * node, which matters because the cell being scrolled to is usually
   * one that is not mounted yet.
   */
  ctx.effect(edit.selection, at => {
    const view = viewport.value;
    if (view.width === 0 || view.height === 0) {
      return;
    }
    const top = HEADER_HEIGHT + at.row * ROW_HEIGHT;
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(at.column);
    const width = sheetWindow.widthOf(at.column);
    // The frozen strips cover the near edges, so a cell is only really
    // visible once it is past them.
    scrollY.value = bring(scrollY.value, top, ROW_HEIGHT, view.height, HEADER_HEIGHT);
    scrollX.value = bring(scrollX.value, left, width, view.width, GUTTER_WIDTH);
  });

  // Cells that scrolled away, so their bindings go with them.
  ctx.effect(sheetWindow.range$, range => {
    for (const [key, mounted] of cells) {
      if (
        mounted.row < range.firstRow ||
        mounted.row > range.lastRow ||
        mounted.column < range.firstColumn ||
        mounted.column > range.lastColumn
      ) {
        cells.delete(key);
        values.release(key);
        standings.release(key);
      }
    }
  });

  return grid;
}

function sameCell(a: { row: number; column: number } | null, b: { row: number; column: number } | null): boolean {
  if (a === null || b === null) {
    return a === b;
  }
  return a.row === b.row && a.column === b.column;
}

/**
 * A scroll offset that brings `[start, start + extent)` into view,
 * without moving when it already is.
 *
 * `lead` is the frozen strip covering the near edge: a row under the
 * header is on screen and not visible, which is a distinction only
 * this function has to make.
 */
function bring(offset: number, start: number, extent: number, viewport: number, lead: number): number {
  if (start < offset + lead) {
    return Math.max(0, start - lead);
  }
  if (start + extent > offset + viewport) {
    return start + extent - viewport;
  }
  return offset;
}

/**
 * Where a cell stands in the selection, as one cheap number.
 *
 * A function of the selection rather than a pipe off it, so the feed
 * that walks the mounted cells can ask about each one without a
 * subscription in between.
 */
function standingOf(selection: SheetSelection, row: number, column: number): Standing {
  if (selection.row === row && selection.column === column) {
    return 2;
  }
  return inRange(selection, row, column) ? 1 : 0;
}

/** Whether the selection rectangle covers a cell. */
function inRange(selection: SheetSelection, row: number, column: number): boolean {
  const firstRow = Math.min(selection.row, selection.anchorRow);
  const lastRow = Math.max(selection.row, selection.anchorRow);
  const firstColumn = Math.min(selection.column, selection.anchorColumn);
  const lastColumn = Math.max(selection.column, selection.anchorColumn);
  return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
}

/**
 * Whether a display string should sit to the right.
 *
 * Decided from the text rather than sent as a flag: the wire carries
 * display strings, and adding a per-cell alignment would double what
 * a window costs to say something the string already says.
 */
function isNumeric(text: string | null): boolean {
  if (text === null || text === '') {
    return false;
  }
  return !Number.isNaN(Number(text.replace(/,/g, '')));
}

export type { SheetWindow };

/** Shared, because the overwhelming majority of cells have no border. */
const NO_SHAPES: DecorationShape[] = [];

/**
 * How thick a reference's outline is.
 *
 * Two pixels rather than one: it has to read as a deliberate mark
 * against the grid's own one-pixel rules, which are the same colour
 * family and everywhere.
 */
const OUTLINE = 2;

function sameWidths(a: readonly number[], b: readonly number[]): boolean {
  return a.length === b.length && a.every((width, index) => width === b[index]);
}
