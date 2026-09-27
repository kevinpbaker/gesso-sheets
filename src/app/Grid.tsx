import { BehaviorSubject, combineLatest, type Observable } from 'rxjs';
import { map } from 'rxjs/operators';

import {
  Box,
  contextMenu,
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
  type UiTextSpan,
  type UiTextChangeEvent,
  type SheetRange,
  type UiVirtualSheet
} from 'gesso-core';
import { Menu, type MenuItem } from 'gesso-components';
import {
  createComponent,
  fanOut,
  FocusService,
  internalState,
  EditingService,
  ShellService,
  TextService,
  type ComponentContext,
  type FanCell,
  type Inputs
} from 'gesso-framework';

import { guessOf, type Place } from './alignment';
import { columnName, relativeRef } from '../sheet/A1';
import { chartElement, dragged, type ChartDrag, type Corner } from './ChartLayer';
import { acceptCompletion, definedFunctionsOf, hintFor, markedArgument, type DefinedFunction, type FormulaHint } from '../sheet/FormulaHint';
import type { Span } from '../sheet/Tokenizer';

import * as dims from './dimensions';
import type { CellEdge, CellPaint } from '../sheet/Format';
import {
  cellIn,
  cornerOf,
  isOneCell,
  PLAIN_PAINT,
  Sheet,
  type SheetExplain,
  type SheetMarked,
  type SheetPasteMode,
  type SheetRowFit,
  type SheetMerge,
  type SheetChart,
  type SheetChartSource,
  type SheetCharts,
  type SheetSelection,
  type SheetSeriesView,
  type SheetWindow
} from './SheetContract';
import { colouredReferences, formulaSpans } from './FormulaColours';
import { cycleAbsolute, pick, repick } from './FormulaEditing';
import { COMMANDS, commandFor, type CommandId } from './SheetCommands';
import { isPrintable, keyAction, stampText } from './SheetKeys';
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
export function Grid(
  _inputs: Inputs<{ editing: SheetEditing; zoom?: number; widen?: number; rebuilt?: boolean }>,
  ctx: ComponentContext
) {
  /**
   * The zoom, read once: the grid is built again when it changes.
   *
   * Every size handed to the engine is multiplied by it — the widths and
   * heights, the strips, the fonts, the charts — so everything the
   * engine works out from them, the offsets, `cellAt`, what is mounted,
   * what a press hit, is right by construction, because it is all real
   * pixels. The document stays at 100%: a size read from the other
   * thread is multiplied on the way in and one sent back is divided on
   * the way out, and a measurement it asks for is made at 100%. A
   * transform over the paint would click in the right place and mount
   * the wrong number of rows. See the roadmap's zoom note.
   */
  const zoom = _inputs.zoom.value ?? 1;
  /**
   * And how much wider the columns are drawn than they are, which is
   * Show formulas: a formula is longer than its answer, and Excel
   * doubles the columns while it shows them. Columns only; a row is as
   * tall as a line of text whatever that text is.
   */
  const widen = _inputs.widen.value ?? 1;
  const scaled = (size: number): number => Math.round(size * zoom);
  /** A size on screen, in the document's pixels. */
  const unscaled = (size: number): number => Math.round(size / zoom);
  const scaledWidth = (size: number): number => Math.round(size * zoom * widen);
  const unscaledWidth = (size: number): number => Math.round(size / (zoom * widen));
  const ROW_HEIGHT = scaled(dims.ROW_HEIGHT);
  const COLUMN_WIDTH = scaledWidth(dims.COLUMN_WIDTH);
  const GUTTER_WIDTH = scaled(dims.GUTTER_WIDTH);
  const HEADER_HEIGHT = scaled(dims.HEADER_HEIGHT);
  const MIN_ROW_HEIGHT = scaled(dims.MIN_ROW_HEIGHT);
  const MAX_ROW_HEIGHT = scaled(dims.MAX_ROW_HEIGHT);
  const MIN_COLUMN_WIDTH = scaledWidth(dims.MIN_COLUMN_WIDTH);
  const MAX_COLUMN_WIDTH = scaledWidth(dims.MAX_COLUMN_WIDTH);
  const CELL_PADDING = scaled(dims.CELL_PADDING);
  const CELL_FONT_SIZE = dims.CELL_FONT_SIZE * zoom;
  const { COLUMN_COUNT, ROW_COUNT } = dims;
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const shell = ctx.inject(ShellService);
  const measure = ctx.inject(TextService);
  const editing = ctx.inject(EditingService);
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
    const box = shapeBox(mounted.row, mounted.column);
    const next = box === null ? NO_SHAPES : bordersOf(
      mounted.paint.value,
      box.width,
      box.height,
      isFlagged(mounted.row, mounted.column),
      noteOf(mounted.row, mounted.column) !== ''
    );
    if (next === NO_SHAPES && mounted.shapes.value === NO_SHAPES) {
      return;
    }
    mounted.shapes.next(next);
  };

  ctx.effect(sheet.view.formats, current => {
    latestFormats = current;
    repaint();
  });

  /**
   * The cells in view that break a rule.
   *
   * Held rather than piped into the cells, on the rule this file has
   * followed since Phase 0: one reader, and a push into the cells
   * that changed. A sheet with no validations publishes an empty
   * object and `isFlagged` is one lookup that misses.
   */
  let latestValidation: Readonly<Record<string, Readonly<Record<string, string>>>> = {};
  const isFlagged = (row: number, column: number): boolean =>
    latestValidation[row]?.[column] !== undefined;

  /**
   * The values the open cell may take, when its rule is a list.
   *
   * The one kind of validation with a dropdown, because it is the
   * only kind where the acceptable values are few and known — which
   * is also what makes it the kind worth enforcing.
   */
  let allowedValues: readonly string[] = [];
  const choices = internalState<readonly string[]>([]);
  const choice = internalState(0);

  ctx.effect(sheet.view.validation, current => {
    latestValidation = current.cells;
    allowedValues = current.list;
    repaint();
  });

  /**
   * The notes on the cells in view, held the way the validation marks
   * are: one reader and a lookup that misses for almost every cell.
   */
  let latestNotes: Readonly<Record<string, Readonly<Record<string, string>>>> = {};
  const noteOf = (row: number, column: number): string => latestNotes[row]?.[column] ?? '';
  ctx.effect(sheet.view.notes, current => {
    latestNotes = current.cells;
    repaint();
  });

  /**
   * The list, narrowed to what has been typed so far.
   *
   * Narrowed rather than filtered away: a draft matching nothing
   * closes the list instead of showing an empty box, because an empty
   * box over the sheet says less than no box at all.
   */
  const refreshChoices = (): void => {
    const draft = edit.draftNow();
    if (draft === null || allowedValues.length === 0 || draft.startsWith('=')) {
      if (choices.value.length > 0) {
        choices.value = [];
        sheetWindow.invalidate();
      }
      return;
    }
    const typed = draft.trim().toUpperCase();
    const shown = allowedValues.filter(value => typed === '' || value.toUpperCase().startsWith(typed));
    const same = shown.length === choices.value.length && shown.every((value, at) => value === choices.value[at]);
    if (same) {
      return;
    }
    choices.value = shown;
    choice.value = 0;
    // The popup is a child of a row and the list is read while the
    // row is built, so changing it has to rebuild the row.
    sheetWindow.invalidate();
  };

  /** The arrows and Enter, while a list is open. */
  const onChoiceKey = (event: UiKeyboardEvent): boolean => {
    const shown = choices.value;
    if (shown.length === 0) {
      return false;
    }
    switch (event.key) {
      case 'ArrowDown':
        choice.value = (choice.value + 1) % shown.length;
        sheetWindow.invalidate();
        return true;
      case 'ArrowUp':
        choice.value = (choice.value + shown.length - 1) % shown.length;
        sheetWindow.invalidate();
        return true;
      case 'Enter':
      case 'Tab': {
        const picked = shown[Math.min(choice.value, shown.length - 1)];
        if (editorNode !== null) {
          const model = editorFor(editorNode);
          model.replaceText(picked);
          model.select(picked.length);
        }
        edit.write(picked);
        choices.value = [];
        sheetWindow.invalidate();
        // Not handled: Enter still commits the cell, which is what
        // somebody pressing it after choosing means. Taking the key
        // here would make choosing and committing two presses of the
        // same key for no reason anybody could see.
        return false;
      }
      case 'Escape':
        // The list goes and the edit stays, which is what Escape
        // means with a list open — the same as the function hints.
        choices.value = [];
        sheetWindow.invalidate();
        return true;
      default:
        return false;
    }
  };

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

  /**
   * What a cell's borders and marks are drawn round: the cell, or the
   * whole merge when it anchors one. The anchor is drawn the merge's
   * size, so a box round it drawn the width of its own column put a
   * rule down the middle of the merge, and a note's mark half way
   * along its top. A cell the merge covers draws nothing: its own
   * borders would be rules inside the merge, where Excel draws none.
   */
  const shapeBox = (row: number, column: number): { width: number; height: number } | null => {
    const merge = mergeAt(row, column);
    if (merge === null) {
      return { width: widths.value[column] ?? COLUMN_WIDTH, height: heightNow(row) };
    }
    if (merge.firstRow !== row || merge.firstColumn !== column) {
      return null;
    }
    let width = 0;
    for (let at = merge.firstColumn; at <= merge.lastColumn; at++) {
      width += widths.value[at] ?? COLUMN_WIDTH;
    }
    let height = 0;
    for (let at = merge.firstRow; at <= merge.lastRow; at++) {
      height += heightNow(at);
    }
    return { width, height };
  };
  /** Whether a column or row that moved is under the cell's shapes. */
  const underShapes = (mounted: MountedCell, moved: ReadonlySet<number>, axis: 'row' | 'column'): boolean => {
    const merge = mergeAt(mounted.row, mounted.column);
    if (merge === null) {
      return moved.has(mounted[axis]);
    }
    const [first, last] = axis === 'row' ? [merge.firstRow, merge.lastRow] : [merge.firstColumn, merge.lastColumn];
    for (const at of moved) {
      if (at >= first && at <= last) {
        return true;
      }
    }
    return false;
  };

  const hidden = internalState<ReadonlySet<number>>(new Set<number>());
  /**
   * The rows that are not the default height: dragged, or fitted to
   * wrapped text. The document's, adopted from the geometry — except
   * mid-drag, where this side leads, as it does for a column.
   */
  const sized = internalState<ReadonlyMap<number, number>>(new Map<number, number>());
  /** The exceptions the window takes: every sized row, and every hidden one at zero. */
  const heightsOf = (rows: ReadonlySet<number>, sizes: ReadonlyMap<number, number>): Map<number, number> => {
    const heights = new Map(sizes);
    for (const row of rows) {
      heights.set(row, 0);
    }
    return heights;
  };
  /**
   * A row's height now, from this side's own state.
   *
   * Not asked of the window, for the reason `columnLeft` is not: a
   * cell is built while the window is still being constructed.
   */
  const heightNow = (row: number): number => (hidden.value.has(row) ? 0 : (sized.value.get(row) ?? ROW_HEIGHT));
  /** Where a frozen row starts, past the header; there are a handful of those. */
  const rowTop = (row: number): number => {
    let top = 0;
    for (let at = 0; at < row; at++) {
      top += heightNow(at);
    }
    return top;
  };

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
      height = new BehaviorSubject(heightNow(row));
      rowHeights.set(row, height);
    }
    return height;
  };
  const refreshHeights = (): void => {
    const moved = new Set<number>();
    for (const [row, height] of rowHeights) {
      const next = heightNow(row);
      if (next !== height.value) {
        height.next(next);
        moved.add(row);
      }
    }
    if (moved.size === 0) {
      return;
    }
    // A bottom border is drawn at the foot of the cell, so a row that
    // changed height redraws the borders in it, as a column does.
    for (const mounted of cells.values()) {
      if (underShapes(mounted, moved, 'row')) {
        refreshShapes(mounted);
      }
    }
  };
  ctx.effect(hidden, refreshHeights);
  ctx.effect(sized, refreshHeights);

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
      if (underShapes(mounted, moved, 'column')) {
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
  const bordersOf = (
    paint: CellPaint,
    width: number,
    height: number,
    flagged: boolean,
    noted = false
  ): readonly DecorationShape[] => {
    const edges = paint.borders;
    if (
      !flagged &&
      !noted &&
      edges.top.width === 0 &&
      edges.right.width === 0 &&
      edges.bottom.width === 0 &&
      edges.left.width === 0
    ) {
      // The common case by a very long way, and it allocates nothing.
      return NO_SHAPES;
    }
    const shapes: DecorationShape[] = [];
    if (noted) {
      // A note's mark, square in the very corner where Excel puts its
      // triangle; the rule's dot sits beside it when a cell has both.
      shapes.push({
        kind: 'fill',
        x: width - NOTE_MARK,
        y: 0,
        width: NOTE_MARK,
        height: NOTE_MARK,
        radius: 0,
        color: 'primary',
        after: 'children'
      });
    }
    if (flagged) {
      /**
       * A corner mark for a cell that breaks its rule.
       *
       * In the cell's own paint pass like the borders, so a sheet
       * full of marked cells costs one draw instance each and no
       * extra nodes — which is what makes marking every cell in the
       * *window* the affordable thing it needs to be.
       */
      shapes.push({
        kind: 'fill',
        x: width - MARK_SIZE - 1 - (noted ? NOTE_MARK + 1 : 0),
        y: 1,
        width: MARK_SIZE,
        height: MARK_SIZE,
        radius: MARK_SIZE / 2,
        color: 'danger',
        after: 'children'
      });
    }
    const edge = (e: CellEdge, box: { x: number; y: number; width: number; height: number }): void => {
      if (e.width > 0) {
        shapes.push({ kind: 'fill', ...box, radius: 0, color: e.color === '' ? 'text' : e.color, after: 'children' });
      }
    };
    edge(edges.top, { x: 0, y: 0, width, height: edges.top.width });
    edge(edges.bottom, { x: 0, y: height - edges.bottom.width, width, height: edges.bottom.width });
    edge(edges.left, { x: 0, y: 0, width: edges.left.width, height });
    edge(edges.right, { x: width - edges.right.width, y: 0, width: edges.right.width, height });
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
    const height = heightNow(row);
    box(left, 0, OUTLINE, height);
    box(right - OUTLINE, 0, OUTLINE, height);
    if (row === firstRow) {
      box(left, 0, width, OUTLINE);
    }
    if (row === lastRow) {
      box(left, height - OUTLINE, width, OUTLINE);
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
  /**
   * The block last copied or cut, while Ctrl+V would still paste it:
   * a dashed outline that moves, so it can be seen what a paste is
   * about to do. On the sheet it was copied from, and only there.
   */
  let marked: SheetMarked | null = null;
  /** How far along the dashes have crawled; stepped by a timer while there is an outline. */
  let marchPhase = 0;
  let marching: ReturnType<typeof setInterval> | null = null;
  const DASH = 5;
  const GAP = 3;

  /** Dashes along one edge, from `start` for `length`, as rectangles on the row. */
  const dashes = (
    into: DecorationShape[],
    along: 'x' | 'y',
    start: number,
    length: number,
    at: number,
    offset: number
  ): void => {
    const period = DASH + GAP;
    for (let from = -((offset + marchPhase) % period); from < length; from += period) {
      const a = Math.max(0, from);
      const b = Math.min(length, from + DASH);
      if (b <= a) {
        continue;
      }
      into.push(
        along === 'x'
          ? { kind: 'fill', x: start + a, y: at, width: b - a, height: OUTLINE, radius: 0, color: 'primary', after: 'children' }
          : { kind: 'fill', x: at, y: start + a, width: OUTLINE, height: b - a, radius: 0, color: 'primary', after: 'children' }
      );
    }
  };

  const marqueeSegments = (rect: SheetMarked, row: number, into: DecorationShape[]): void => {
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(rect.firstColumn);
    const right = GUTTER_WIDTH + sheetWindow.offsetOf(rect.lastColumn) + sheetWindow.widthOf(rect.lastColumn);
    const height = heightNow(row);
    if (height === 0 || right <= left) {
      return;
    }
    // The sides run on down from row to row, so their dashes are laid
    // out from where the row starts on the sheet, not from its top.
    const top = sheetWindow.rowOffsetOf(row);
    dashes(into, 'y', 0, height, left, top);
    dashes(into, 'y', 0, height, right - OUTLINE, top);
    if (row === rect.firstRow) {
      dashes(into, 'x', left, right - left, 0, 0);
    }
    if (row === rect.lastRow) {
      dashes(into, 'x', left, right - left, height - OUTLINE, 0);
    }
  };

  /**
   * Where a block dragged by its border would land, drawn as a solid
   * outline while the drag is on — the same rectangles a formula's
   * references are drawn with, one row at a time.
   */
  let landing: { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number } | null = null;

  /** The charts on this sheet, and what they draw; fed below, read by the outlines as well. */
  const charts = internalState<SheetCharts>({ entries: [], selected: 0 });
  const series = internalState<SheetSeriesView>({ charts: {} });
  /** Which sheet is showing, for whether a chart's cells are on it. */
  let activeSheet = 0;
  /** The selected chart's cells, when they are on the sheet showing; see `paintOutlines`. */
  const chartSource = (): SheetChartSource | null => {
    const selected = charts.value.selected;
    if (selected === 0) {
      return null;
    }
    const source = series.value.charts[String(selected)]?.source ?? null;
    return source !== null && source.sheet === activeSheet ? source : null;
  };

  const paintOutlines = (draft: string | null): void => {
    const references = draft === null ? [] : colouredReferences(draft);
    const rows = new Map<number, DecorationShape[]>();
    if (landing !== null) {
      const range = sheetWindow.range$.value;
      const box = {
        start: { row: landing.firstRow, column: landing.firstColumn },
        end: { row: landing.lastRow, column: landing.lastColumn }
      };
      for (let row = Math.max(landing.firstRow, 0); row <= Math.min(landing.lastRow, range.lastRow); row++) {
        const shapes: DecorationShape[] = [];
        outlineSegments(box, 'primary', row, shapes);
        rows.set(row, shapes);
      }
    }
    if (marked !== null) {
      // Only the rows the window has, so a whole column copied costs
      // thirty rows of dashes and not ten thousand.
      const range = sheetWindow.range$.value;
      const first = Math.max(marked.firstRow, Math.min(range.firstRow, frozen.value.rows > 0 ? 0 : range.firstRow));
      const last = Math.min(marked.lastRow, range.lastRow);
      for (let row = first; row <= last; row++) {
        const shapes: DecorationShape[] = [];
        marqueeSegments(marked, row, shapes);
        rows.set(row, shapes);
      }
    }
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
    /**
     * The selected chart's cells, in the three colours Excel draws them
     * in: the series' names red, the categories purple, the values
     * blue. Only while the chart is selected and its cells are on the
     * sheet in view — a chart of another sheet's cells has nothing here
     * to point at — and only the rows the window has, because the proof
     * page's chart reads fifty thousand of them.
     */
    const source = chartSource();
    if (source !== null) {
      const range = sheetWindow.range$.value;
      for (const [part, color] of [
        [source.names, CHART_NAMES],
        [source.categories, CHART_CATEGORIES],
        [source.values, CHART_VALUES]
      ] as const) {
        if (part === null) {
          continue;
        }
        const box = { start: { row: part.firstRow, column: part.firstColumn }, end: { row: part.lastRow, column: part.lastColumn } };
        const first = Math.max(part.firstRow, frozen.value.rows > 0 ? 0 : range.firstRow);
        const last = Math.min(part.lastRow, range.lastRow);
        for (let row = first; row <= last; row++) {
          let shapes = rows.get(row);
          if (shapes === undefined) {
            shapes = [];
            rows.set(row, shapes);
          }
          outlineSegments(box, color, row, shapes);
        }
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
  const repaintOutlines = (): void => paintOutlines(edit.draftNow());
  ctx.effect(sheet.view.sheets, tabs => {
    if (tabs.active !== activeSheet) {
      activeSheet = tabs.active;
      repaintOutlines();
    }
  });

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
      // A covered cell draws no grid line: it is inside the merge, and
      // a box of no width with a border is still a line down it.
      borderWidth: covered ? 0 : state.pipe(map(where => (where === 2 ? 2 : 1))),
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
      fontSize: paint.pipe(map(how => (how.fontSize === 0 ? CELL_FONT_SIZE : how.fontSize * zoom))),
      /**
       * Wrapped text breaks at the cell's width, and the row is made
       * tall enough to hold it — see the `rowFit` effect below.
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
          how.align === 'auto' ? AUTO_ALIGN[guessOf(text)] : how.align === 'center' ? 'center' : how.align
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

  /** How the next paste event lands; see Ctrl+Shift+V in `onKey`. */
  let pasteMode: SheetPasteMode = 'all';

  /**
   * The format painter, when it is lit, paints what was just selected
   * by the pointer — a click, a sweep's end, a header — and goes out
   * unless it was double-clicked on.
   */
  const paintIfLit = (): void => {
    const state = edit.painter.value;
    if (state === 'off') {
      return;
    }
    sheet.send.paintFormats();
    if (state === 'once') {
      edit.setPainter('off');
    }
  };

  /** The cell the fill handle hangs off: the selection's far corner. */
  let cornerAt: { row: number; column: number } | null = null;
  /** The selection's first corner, where a finger's other handle sits. */
  let startAt: { row: number; column: number } | null = null;

  /**
   * Whether the last press was a finger's.
   *
   * A finger cannot hover to see what a sweep would select, and a
   * sweep is also a scroll, so a touch pointer gets what a mouse does
   * not need: grips wide enough to find, and two round handles on the
   * selection's corners in place of the fill handle. Read off every
   * press, so a laptop with a touchscreen changes as the hand does.
   */
  const touching = internalState(false);
  const GRIP = (finger: boolean): number => (finger ? 24 : 8);
  const HANDLE = 22;

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

  /**
   * The runs the open cell draws itself in.
   *
   * A subject rather than a pipe off the draft, because the runs
   * depend on the *caret* as well as the text: the bracket beside it
   * and the one that closes it are washed.
   *
   * Fed from three places, which between them are every way a caret
   * moves. The draft changing covers typing; `onSelectionChange`
   * covers arrows, clicks into the text and select-all, which change
   * no text at all and which nothing else reports; and a pick or an
   * F4 refreshes directly, because both move the caret by hand after
   * writing.
   *
   * The middle one is an engine signal added for this. Before it, a
   * bracket arrowed onto stayed dark until the next keystroke.
   */
  const editorSpans = new BehaviorSubject<readonly UiTextSpan[] | undefined>(undefined);

  /**
   * What to offer the person typing: a list of names, or a signature.
   *
   * Held beside the runs and refreshed by the same three things,
   * because it answers the same question — what is under the caret —
   * and two answers computed at different moments would disagree
   * about it.
   */
  const hint = internalState<FormulaHint>(null);
  /**
   * Why the selected cell is showing an error, when it is.
   *
   * Worked out on the application worker — finding the cell that
   * *made* an error is a walk back through the dependency graph, and
   * the graph is not on the wire — so this holds a sentence and an
   * address and draws them.
   */
  const explain = internalState<SheetExplain | null>(null);
  /** Which name in the list is picked out, an index into `names`. */
  const chosen = internalState(0);
  /** The workbook's own functions, named LAMBDAs, for the list and the hint. */
  let definedFunctions: readonly DefinedFunction[] = [];
  ctx.effect(sheet.view.names, current => {
    definedFunctions = definedFunctionsOf(current.formulas);
  });

  const refreshSpans = (): void => {
    const draft = edit.draftNow();
    if (draft === null) {
      editorSpans.next(undefined);
      hint.value = null;
      return;
    }
    const caret = editorNode === null ? undefined : editorFor(editorNode).focus;
    editorSpans.next(formulaSpans(draft, caret));

    const next = caret === undefined ? null : hintFor(draft, caret, definedFunctions);
    /**
     * The choice survives a keystroke that did not change the list.
     *
     * Typing another letter of a name usually narrows it, and a list
     * that jumped back to its first item on every character would be
     * unusable at exactly the speed people type. It is reset when the
     * names actually change, because holding an index into a list
     * that has been replaced points at something nobody chose.
     */
    if (!sameNames(hint.value, next)) {
      chosen.value = 0;
    }
    /**
     * A changed hint rebuilds the rows, because the popup is a child
     * of one.
     *
     * Only when it actually changed: typing inside `SUM(` leaves the
     * signature saying the same thing on most keystrokes, and a
     * rebuild per character for a hint that did not move is the
     * expensive half of this feature for none of the benefit. The
     * comparison is cheap and the rebuild is not.
     */
    const changed = !sameHint(hint.value, next);
    hint.value = next;
    if (changed) {
      sheetWindow.invalidate();
    }
  };

  ctx.effect(sheet.view.editor, view => {
    const before = explain.value;
    explain.value = view.explain;
    // The popup is a child of a row, so a change has to rebuild one.
    if ((before === null) !== (view.explain === null) || before?.code !== view.explain?.code) {
      sheetWindow.invalidate();
    }
  });

  ctx.effect(edit.draft, draft => {
    if (draft !== pickedText) {
      picking = null;
      pickedText = null;
    }
    refreshSpans();
    refreshChoices();
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
    refreshSpans();
    return true;
  };

  /**
   * What a press on the strips landed on: the corner, a column's
   * letter or a row's number — or nothing, for a press on the cells.
   *
   * One question asked of the grid's own press, rather than a handler
   * on every header: the strips are mounted for every visible row and
   * column, and a binding per header is the per-cell cliff again. A
   * frozen column's letter does not scroll, so it is read from the
   * offset as it stands; anything past the pane from where the sheet
   * has scrolled to.
   */
  type HeaderHit = { kind: 'corner' } | { kind: 'column'; column: number } | { kind: 'row'; row: number };
  const headerHit = (event: { x: number; y: number }): HeaderHit | null => {
    const box = viewport.value;
    if (box.width === 0) {
      return null;
    }
    const x = event.x - box.x;
    const y = event.y - box.y;
    const inHeader = y >= 0 && y < HEADER_HEIGHT;
    const inGutter = x >= 0 && x < GUTTER_WIDTH;
    if (inHeader && inGutter) {
      return { kind: 'corner' };
    }
    if (inHeader) {
      return { kind: 'column', column: columnAtX(x) };
    }
    if (inGutter && y >= HEADER_HEIGHT) {
      return { kind: 'row', row: rowAtY(y) };
    }
    return null;
  };
  /**
   * The same question asked of the node that was pressed, which is
   * the answer that does not depend on where: a click a screen reader
   * sends has no point, and a letter is a letter wherever on it the
   * press landed. Walks up to the strip's own nodes, which carry what
   * they are — the column in `posInSet`, the row in the label.
   */
  const headerOfNode = (node: UiNode | null): HeaderHit | null => {
    for (let at: UiNode | null = node; at !== null && at !== gridNode; at = at.parent) {
      const role = at.properties.get('role');
      if (role === 'columnheader') {
        return { kind: 'column', column: Number(at.properties.get('posInSet')) - 1 };
      }
      if (role === 'rowheader') {
        return { kind: 'row', row: Number(at.properties.get('label')) - 1 };
      }
      if (role === 'button' && at.properties.get('label') === SELECT_ALL) {
        return { kind: 'corner' };
      }
      if (role === 'cell' || role === 'row') {
        return null;
      }
    }
    return null;
  };
  const columnAtX = (x: number): number =>
    x - GUTTER_WIDTH < columnLeft(frozen.value.columns) ? sheetWindow.columnAt(x) : sheetWindow.cellAt(x, HEADER_HEIGHT).column;
  const rowAtY = (y: number): number =>
    y - HEADER_HEIGHT < rowTop(frozen.value.rows) ? sheetWindow.rowAt(y) : sheetWindow.cellAt(GUTTER_WIDTH, y).row;

  /**
   * Whole columns, whole rows, or the sheet, as the selection.
   *
   * The active cell is a corner — row 1 of the first column, the first
   * column of the first row — which is where Excel puts it, and the one
   * move of the selection that must not scroll: somebody who clicked a
   * letter half way down a sheet is looking at the half they clicked.
   */
  let keepScroll = false;
  /**
   * `active` is where the pointer is and `anchor` where the selection
   * started, as a Shift+click on the cells has it: the end that moves
   * is the active one.
   */
  const selectColumns = (corner: number, anchor: number, active?: { row: number; column: number }): void => {
    keepScroll = true;
    edit.selectRect(0, corner, rowCount.value - 1, anchor, active);
    keepScroll = false;
  };
  const selectRows = (corner: number, anchor: number, active?: { row: number; column: number }): void => {
    keepScroll = true;
    edit.selectRect(corner, 0, anchor, columnCount() - 1, active);
    keepScroll = false;
  };
  const columnCount = (): number => sheet.view.geometry.value.columnCount;
  const selectHeader = (hit: HeaderHit, extend: boolean): void => {
    if (edit.openNow()) {
      edit.commit(0, 0);
    }
    const held = edit.selectionNow();
    if (hit.kind === 'corner') {
      keepScroll = true;
      edit.selectRect(0, 0, rowCount.value - 1, columnCount() - 1);
      keepScroll = false;
    } else if (hit.kind === 'column') {
      // Shift keeps the active cell, here as on the cells.
      selectColumns(hit.column, extend ? held.anchorColumn : hit.column, extend ? held : undefined);
    } else {
      selectRows(hit.row, extend ? held.anchorRow : hit.row, extend ? held : undefined);
    }
    if (gridNode !== null) {
      focus.focus(gridNode);
    }
  };
  /**
   * Whether a point is on the selection's border — within three pixels
   * outside one of its edges or two inside, and along it — which is
   * where a press picks
   * the block up rather than starting a sweep.
   *
   * Worked out from the offsets by the grid's own press, as a header
   * is: a node along each edge would be four more nodes that move on
   * every arrow key, for a question arithmetic answers. The fill handle
   * sits on the corner and stops its own press, so it is never seen
   * here.
   */
  const BORDER = 3;
  const INSIDE = 2;
  const onBorder = (event: { x: number; y: number }): boolean => {
    const box = viewport.value;
    if (box.width === 0) {
      return false;
    }
    const x = event.x - box.x;
    const y = event.y - box.y;
    if (x < GUTTER_WIDTH || y < HEADER_HEIGHT) {
      return false;
    }
    const at = edit.selectionNow();
    const firstRow = Math.min(cornerOf(at).row, at.anchorRow);
    const lastRow = Math.max(cornerOf(at).row, at.anchorRow);
    const firstColumn = Math.min(cornerOf(at).column, at.anchorColumn);
    const lastColumn = Math.max(cornerOf(at).column, at.anchorColumn);
    const xOf = (column: number): number =>
      GUTTER_WIDTH + sheetWindow.offsetOf(column) - (column < frozen.value.columns ? 0 : scrollX.value);
    const yOf = (row: number): number =>
      HEADER_HEIGHT + sheetWindow.rowOffsetOf(row) - (row < frozen.value.rows ? 0 : scrollY.value);
    const left = xOf(firstColumn);
    const right = xOf(lastColumn) + sheetWindow.widthOf(lastColumn);
    const top = yOf(firstRow);
    const bottom = yOf(lastRow) + sheetWindow.rowHeightOf(lastRow);
    // Mostly outside the line, as Excel's is: a press a few pixels
    // into a selected cell is a sweep starting there, not a pick-up.
    // Along an edge means within its own span: past a corner, outside
    // both edges at once, is the next cell over and not the border.
    const alongX = x >= left && x <= right;
    const alongY = y >= top && y <= bottom;
    const near = (at: number, edge: number, outward: number): boolean =>
      outward < 0 ? at >= edge - BORDER && at <= edge + INSIDE : at >= edge - INSIDE && at <= edge + BORDER;
    return (
      (alongY && (near(x, left, -1) || near(x, right, 1))) || (alongX && (near(y, top, -1) || near(y, bottom, 1)))
    );
  };
  /** A block being dragged by its border: where it was picked up, and the cell under the press. */
  let carrying: { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number; row: number; column: number } | null =
    null;
  /** Over the border, the pointer says the block can be moved. */
  const pointer = internalState<'move' | undefined>(undefined);

  /** A header click, which is also somewhere the painter paints. */
  const clickHeader = (hit: HeaderHit, extend: boolean): void => {
    selectHeader(hit, extend);
    paintIfLit();
  };
  /** A drag across the letters or the numbers, from where it started. */
  let headerSweep: { kind: 'column' | 'row'; from: number } | null = null;

  /**
   * The menu a right-click opens, built from the command table like
   * every other menu so that it cannot drift: what the pointer was on
   * decides which commands are in it.
   */
  const menuOpen = new BehaviorSubject(false);
  const menuAt = new BehaviorSubject({ x: 0, y: 0 });
  const menuItems = new BehaviorSubject<readonly MenuItem[]>([]);
  const MENU_FOR: Readonly<Record<'cell' | 'column' | 'row', readonly CommandId[]>> = {
    cell: ['cut', 'copy', 'paste', 'insertRowAbove', 'insertColumnLeft', 'deleteRows', 'deleteColumns', 'sortAscending', 'sortDescending', 'clear', 'editNote'],
    column: ['cut', 'copy', 'paste', 'insertColumnLeft', 'insertColumnRight', 'deleteColumns', 'hideColumns', 'showColumns', 'autofitColumns', 'sortAscending', 'sortDescending', 'clear'],
    row: ['cut', 'copy', 'paste', 'insertRowAbove', 'insertRowBelow', 'deleteRows', 'hideRows', 'showRows', 'fitRows', 'clear']
  };
  const openMenu = (kind: 'cell' | 'column' | 'row', at: { x: number; y: number }): void => {
    menuItems.next(MENU_FOR[kind].map(id => ({ value: id, label: COMMANDS[id].label })));
    menuAt.next(at);
    menuOpen.next(true);
  };
  /**
   * A right-click: on something outside the selection, that thing is
   * selected first, as everywhere — the menu acts on what it was opened
   * over. Inside the selection, the selection stays.
   */
  const onContext = (at: { x: number; y: number }): void => {
    const hit = headerHit(at);
    const held = edit.selectionNow();
    const firstRow = Math.min(cornerOf(held).row, held.anchorRow);
    const lastRow = Math.max(cornerOf(held).row, held.anchorRow);
    const firstColumn = Math.min(cornerOf(held).column, held.anchorColumn);
    const lastColumn = Math.max(cornerOf(held).column, held.anchorColumn);
    if (hit?.kind === 'column') {
      const whole = firstRow === 0 && lastRow >= rowCount.value - 1;
      if (!whole || hit.column < firstColumn || hit.column > lastColumn) {
        selectHeader(hit, false);
      }
      openMenu('column', at);
      return;
    }
    if (hit?.kind === 'row') {
      const whole = firstColumn === 0 && lastColumn >= columnCount() - 1;
      if (!whole || hit.row < firstRow || hit.row > lastRow) {
        selectHeader(hit, false);
      }
      openMenu('row', at);
      return;
    }
    if (hit?.kind === 'corner') {
      selectHeader(hit, false);
      openMenu('cell', at);
      return;
    }
    const cellAt = cellUnder({ x: at.x, y: at.y } as UiPointerEvent);
    if (cellAt !== null && !(cellAt.row >= firstRow && cellAt.row <= lastRow && cellAt.column >= firstColumn && cellAt.column <= lastColumn)) {
      selectByPointer(cellAt.row, cellAt.column);
    }
    openMenu('cell', at);
  };
  /** Shift+F10 and the Menu key: the same menu, at the active cell. */
  const openMenuFromKeyboard = (): void => {
    const box = viewport.value;
    const at = edit.selectionNow();
    const x = box.x + GUTTER_WIDTH + sheetWindow.offsetOf(at.column) - scrollX.value + 8;
    const y = box.y + HEADER_HEIGHT + sheetWindow.rowOffsetOf(at.row) - scrollY.value + sheetWindow.rowHeightOf(at.row);
    openMenu('cell', { x: Math.max(box.x, Math.min(x, box.x + box.width - 8)), y: Math.max(box.y, Math.min(y, box.y + box.height - 8)) });
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
    paintIfLit();
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
    const box = shapeBox(row, column);
    const shapes = new BehaviorSubject<readonly DecorationShape[]>(
      box === null ? NO_SHAPES : bordersOf(
        paint.value,
        box.width,
        box.height,
        isFlagged(row, column),
        noteOf(row, column) !== ''
      )
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
          /**
           * Once per edit, not once per rebuild.
           *
           * The caret goes to the end when a cell *opens*. This ref
           * fires again every time the row is rebuilt — and the row is
           * rebuilt whenever a hint appears, a choice moves or the
           * selection changes — so placing the caret here
           * unconditionally snapped it back to the end on every one of
           * them. Arrowing left moved the caret and the next frame put
           * it back, which looks like an arrow key that does not work.
           */
          if (!caretPlaced) {
            caretPlaced = true;
            model.select(text.length);
            typed = text;
            // The letter that opened the cell is typing too, and it is
            // the one that most often has a word to finish.
            if (offerOnOpen) {
              offerOnOpen = false;
              offerCompletion(text, model.focus);
            }
          }
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
      spans: editorSpans,
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
      onInput: (event: UiTextChangeEvent) => {
        const before = typed ?? '';
        const shown = offered;
        offered = null;
        edit.write(event.value);
        typed = event.value;
        // Backspace over an offer takes the offer away and nothing else,
        // and must not be answered with the same offer again.
        if (shown !== null && event.value === before) {
          return;
        }
        const grew = event.value.length > before.length && fold(event.value).startsWith(fold(before));
        if (grew && editorNode !== null) {
          offerCompletion(event.value, editorFor(editorNode).focus);
        }
      },
      // The caret moving with the text unchanged: an arrow key, a
      // click into the text, select-all. Nothing else reports it, and
      // the bracket beside the caret depends on it.
      onSelectionChange: () => refreshSpans(),
      onKeyDown: onKey
    });
  };

  /** The node holding the open cell's editor, so focus can be put in it. */
  let editorNode: UiNode | null = null;
  /** The cell the renderer should build as an editor, read while building. */
  let openAt: { row: number; column: number } | null = null;
  /**
   * Whether this edit has had its caret placed.
   *
   * Reset when a cell opens or closes; see the editor's `ref`.
   */
  let caretPlaced = false;

  /**
   * AutoComplete: the rest of a word the column already holds, offered
   * selected after what was typed, so Enter takes it and typing on
   * replaces it.
   *
   * `typed` is what the person has typed, which is not what the cell
   * shows while an offer is up; `offered` is the offer on screen. The
   * question goes to the other thread, which holds the whole column,
   * and the answer is used only while it is still about this cell and
   * this text — an answer that arrives after another key is dropped.
   */
  let typed: string | null = null;
  let offered: string | null = null;
  let offerOnOpen = false;
  let completionSerial = 0;
  const offerCompletion = (text: string, caret: number): void => {
    if (text === '' || text.startsWith('=') || caret !== text.length || choices.value.length > 0) {
      return;
    }
    const at = edit.selectionNow();
    completionSerial++;
    sheet.send.complete(at.row, at.column, text, completionSerial);
  };
  ctx.effect(sheet.view.completion, answer => {
    if (answer.serial !== completionSerial || answer.text === '' || editorNode === null) {
      return;
    }
    const draft = edit.draftNow();
    const at = edit.selectionNow();
    if (draft !== answer.prefix || draft !== typed || at.row !== answer.row || at.column !== answer.column) {
      return;
    }
    const model = editorFor(editorNode);
    if (!model.collapsed || model.focus !== draft.length) {
      return;
    }
    // The word as the column spells it: `nor` becomes North, as it does
    // in Excel, so that Enter writes the word that is already there.
    const whole = answer.text;
    offered = whole;
    edit.write(whole);
    model.replaceText(whole);
    model.select(draft.length, whole.length);
    refreshSpans();
  });

  /** The frozen strip at the start of a row: the row's number. */
  /**
   * A row's number, with the grip along its foot that resizes it.
   *
   * The same grip a column's label has, turned on its side: a child of
   * the label, eight pixels across a one-pixel rule, and a Pan rather
   * than a Drag. A row dragged to a height keeps it, whatever the row
   * holds, until "Fit rows to contents" gives it back.
   */
  const rowHeader = (row: number): UiElement =>
    Box(
      {
        key: 'gutter',
        width: GUTTER_WIDTH,
        height: heightOf(row),
        flexShrink: 0,
        x: 'center',
        y: 'center',
        // Held at the left edge while the sheet scrolls sideways. Its
        // vertical travel is its row's, which it gets for free by being
        // inside it. Sticky positions it, so the grip is placed
        // against it rather than against the layout root.
        position: 'sticky',
        left: 0,
        zIndex: 1,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1,
        role: 'rowheader',
        label: String(row + 1)
      },
      Text({
        text: String(row + 1),
        color: 'textMuted',
        fontSize: 11,
        textAlign: 'center',
        verticalAlign: 'middle',
        selectable: false
      }),
      Box({
        key: 'grip',
        width: GUTTER_WIDTH,
        height: touching.pipe(map(GRIP)),
        position: 'absolute',
        left: 0,
        bottom: touching.pipe(map(finger => -GRIP(finger) / 2)),
        zIndex: 4,
        cursor: 'row-resize',
        onPanStart: event => {
          heightDrag = { row, height: heightNow(row), from: event.y };
          // The grid sweeps a selection on a pan, and the gutter is
          // over a row of cells as far as the offsets are concerned: a
          // drag here would resize the row and select it as well.
          event.stopPropagation();
        },
        onPanMove: event => {
          if (heightDrag !== null) {
            event.stopPropagation();
            const next = new Map(sized.value);
            next.set(
              heightDrag.row,
              Math.min(Math.max(Math.round(heightDrag.height + (event.y - heightDrag.from)), MIN_ROW_HEIGHT), MAX_ROW_HEIGHT)
            );
            sized.value = next;
          }
        },
        // Two clicks on a row's foot fit it to what it holds, which for a
        // row somebody dragged means forgetting the drag.
        onClick: (event: UiPointerEvent) => event.stopPropagation(),
        onDoubleClick: (event: UiPointerEvent) => {
          event.stopPropagation();
          sheet.send.fitRowsToContents(row, row);
        },
        onPanEnd: event => {
          if (heightDrag !== null) {
            event.stopPropagation();
            // Written down when the drag ends, as a column's width is.
            sheet.send.setRowHeight(heightDrag.row, unscaled(sized.value.get(heightDrag.row) ?? ROW_HEIGHT));
          }
          heightDrag = null;
        }
      })
    );

  let heightDrag: { row: number; height: number; from: number } | null = null;

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
     * A key that beat the editor to the keyboard.
     *
     * The first letter opens the cell, and the editor takes focus when
     * it mounts, a frame later. Keys typed inside that frame still
     * arrive here, at the grid, and the table below — rightly — leaves
     * every key of an open cell to its text field, which does not have
     * the keyboard yet: so they fell between the two and only the first
     * letter survived. Found in Chrome by typing a word quickly; the
     * harness draws its frames as it goes and never opens the gap. The
     * grid holds the draft's other end, so it writes them there.
     */
    const chord = event.modifiers.ctrl === true || event.modifiers.meta === true || event.modifiers.alt === true;
    if (edit.openNow() && event.target === gridNode && !chord && isPrintable(event.key)) {
      edit.write((edit.draftNow() ?? '') + event.key);
      typed = edit.draftNow();
      event.preventDefault();
      event.stopPropagation();
      return;
    }
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
      // Ctrl+Shift+V is the browser's paste as plain text, and the paste
      // event it sends a moment later lands as values.
      pasteMode = (event.key === 'v' || event.key === 'V') && (event.modifiers.ctrl || event.modifiers.meta) && event.modifiers.shift ? 'values' : 'all';
      if (event.key === 'Escape' && (edit.painter.value !== 'off' || sheet.view.clipboard.value.marked !== null)) {
        event.preventDefault();
        event.stopPropagation();
        edit.setPainter('off');
        sheet.send.unmark();
        return;
      }
      if (event.key === 'ContextMenu' || (event.key === 'F10' && event.modifiers.shift)) {
        event.preventDefault();
        event.stopPropagation();
        openMenuFromKeyboard();
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
    /**
     * The suggestion list takes the keys that are about it.
     *
     * Before the accelerators and before `keyAction`, because while a
     * list is open Enter means "take this name" rather than "commit
     * the cell", and Escape means "put the list away" rather than
     * "throw the edit away". That is a mode, and it is a narrow one:
     * it exists only while a list is on screen, it takes five keys,
     * and everything else still goes to the text.
     */
    if (choices.value.length > 0 && onChoiceKey(event)) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (hint.value?.kind === 'completions' && onHintKey(event)) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    /**
     * F4, which belongs to the cell rather than to the sheet.
     *
     * Handled here and not in the command table because it is only
     * ever meaningful with a cell open and a caret on a reference,
     * and a command is a thing the whole application can do. It also
     * does nothing rather than beeping when the caret is elsewhere,
     * which a command with an accelerator could not express.
     */
    if (event.key === 'F4' && edit.openNow() && cycleAtCaret()) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const action = keyAction(event.key, event.modifiers, edit.openNow());
    if (action?.kind === 'replace') {
      offerOnOpen = true;
    }
    // Alt+Enter's line break, and a date typed into a cell that is
    // already open, go in at the caret — which is this side's.
    if (action !== null && (action.kind === 'insert' || action.kind === 'stamp') && edit.openNow()) {
      if (insertAtCaret(action.kind === 'insert' ? action.text : stampText(action.what))) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
    if (edit.apply(action)) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  /**
   * A key while a list of names is open. Returns whether it was ours.
   *
   * Tab as well as Enter, because Tab is what a list of completions
   * takes everywhere else and a spreadsheet's Tab — commit and move
   * right — is what somebody who wanted that would press *after*
   * choosing.
   */
  const onHintKey = (event: UiKeyboardEvent): boolean => {
    const current = hint.value;
    if (current?.kind !== 'completions') {
      return false;
    }
    const shown = Math.min(current.names.length, 8);
    switch (event.key) {
      case 'ArrowDown':
        chosen.value = (chosen.value + 1) % shown;
        // The popup is a child of a row and the choice is read while
        // the row is built, so moving it has to rebuild the row.
        // Without this the list moves in the state and not on the
        // screen, which is the same bug as not moving at all.
        sheetWindow.invalidate();
        return true;
      case 'ArrowUp':
        chosen.value = (chosen.value + shown - 1) % shown;
        sheetWindow.invalidate();
        return true;
      case 'Enter':
      case 'Tab':
        acceptHint(current.names[Math.min(chosen.value, shown - 1)], current.span);
        return true;
      case 'Escape':
        // The list goes away and the edit stays. A second Escape is
        // the one that throws the edit away, which is what Escape
        // means with no list open.
        hint.value = null;
        sheetWindow.invalidate();
        return true;
      default:
        return false;
    }
  };

  /** Writes a chosen name into the draft, with its bracket. */
  const acceptHint = (name: string, span: Span): void => {
    const draft = edit.draftNow();
    if (draft === null || editorNode === null) {
      return;
    }
    const written = acceptCompletion(draft, span, name);
    pickedText = written.text;
    edit.write(written.text);
    const model = editorFor(editorNode);
    model.replaceText(written.text);
    model.select(written.caret);
    refreshSpans();
  };

  /**
   * `A1 → $A$1 → A$1 → $A1`, on the reference the caret is in.
   *
   * The cycle is `FormulaEditing.cycleAbsolute`, with a table beside
   * it; this is the part that knows where the caret is and puts it
   * back. Returns false when the caret is not on a reference, so the
   * key falls through to whatever else wanted it.
   */
  const cycleAtCaret = (): boolean => {
    if (editorNode === null) {
      return false;
    }
    const draft = edit.draftNow();
    if (draft === null) {
      return false;
    }
    const model = editorFor(editorNode);
    const cycled = cycleAbsolute(draft, model.focus);
    if (cycled === null) {
      return false;
    }
    pickedText = cycled.text;
    edit.write(cycled.text);
    model.replaceText(cycled.text);
    model.select(cycled.caret);
    refreshSpans();
    return true;
  };

  /** Text typed in at the caret of the open cell, over whatever is selected in it. */
  const insertAtCaret = (text: string): boolean => {
    const draft = edit.draftNow();
    if (editorNode === null || draft === null) {
      return false;
    }
    const model = editorFor(editorNode);
    const next = draft.slice(0, model.start) + text + draft.slice(model.end);
    const caret = model.start + text.length;
    edit.write(next);
    model.replaceText(next);
    model.select(caret);
    refreshSpans();
    return true;
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
  /**
   * A round handle on a corner of the selection, for a finger.
   *
   * Dragged, it moves its corner and leaves the other where it was,
   * and the active cell with it when it is still inside — Shift's rule,
   * for a hand with no Shift. It stops its own press, as the fill
   * handle does, so the drag is its own and not a scroll or a sweep.
   */
  const selectionHandle = (row: number, column: number, which: 'start' | 'end'): UiElement =>
    Box({
      key: `handle-${which}`,
      width: HANDLE,
      height: HANDLE,
      position: 'absolute',
      left:
        GUTTER_WIDTH + sheetWindow.offsetOf(column) + (which === 'end' ? sheetWindow.widthOf(column) : 0) - HANDLE / 2,
      top: (which === 'end' ? heightNow(row) : 0) - HANDLE / 2,
      zIndex: 3,
      borderRadius: HANDLE / 2,
      backgroundColor: 'background',
      borderColor: 'primary',
      borderWidth: 3,
      role: 'button',
      label: which === 'start' ? 'Selection start' : 'Selection end',
      onPanStart: (event: UiPointerEvent) => {
        const held = edit.selectionNow();
        const corner = cornerOf(held);
        const first = { row: Math.min(corner.row, held.anchorRow), column: Math.min(corner.column, held.anchorColumn) };
        const last = { row: Math.max(corner.row, held.anchorRow), column: Math.max(corner.column, held.anchorColumn) };
        handleDrag = { fixed: which === 'end' ? first : last };
        event.stopPropagation();
      },
      onPanMove: (event: UiPointerEvent) => {
        if (handleDrag === null) {
          return;
        }
        event.stopPropagation();
        const at = cellUnder(event);
        if (at === null) {
          return;
        }
        const held = edit.selectionNow();
        const fixed = handleDrag.fixed;
        const inside =
          held.row >= Math.min(at.row, fixed.row) &&
          held.row <= Math.max(at.row, fixed.row) &&
          held.column >= Math.min(at.column, fixed.column) &&
          held.column <= Math.max(at.column, fixed.column);
        edit.selectRect(at.row, at.column, fixed.row, fixed.column, inside ? held : fixed);
      },
      onPanEnd: (event: UiPointerEvent) => {
        event.stopPropagation();
        handleDrag = null;
        handlesSettled.value = handlesSettled.value + 1;
      }
    });
  let handleDrag: { fixed: { row: number; column: number } } | null = null;
  const handlesSettled = internalState(0);

  const fillHandle = (row: number, column: number): UiElement =>
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
      top: heightNow(row) - 5,
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
      // Two clicks fill down as far as the column beside goes: the fill
      // everybody uses, and one that otherwise takes a drag the length
      // of the data.
      onClick: (event: UiPointerEvent) => event.stopPropagation(),
      onDoubleClick: (event: UiPointerEvent) => {
        event.stopPropagation();
        sheet.send.fillToData();
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
   * The chart being dragged, as a live rectangle.
   *
   * On this thread for the length of the drag, like a column resize:
   * the application worker learns where a chart ended up and not
   * where it passed through, so a move is frame-rate here and one
   * command at the end.
   */
  let chartDrag: ChartDrag | null = null;
  let chartFrom: { x: number; y: number; width: number; height: number; pointerX: number; pointerY: number } | null =
    null;

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

  /**
   * The hint under the cell being typed into.
   *
   * Placed from the caret rather than from the cell's left edge, so a
   * list of function names appears under the word it is completing
   * rather than under the start of the formula. The caret's position
   * is the engine's to know — it depends on the paragraph as it was
   * laid out — and `EditingService` is where it is asked for.
   *
   * A child of the row, like the fill handle, so it travels with a
   * scroll. It is deliberately not focusable and takes no pointer
   * events: the keyboard stays in the cell, which is what makes
   * typing through a list of suggestions feel like typing rather than
   * like operating a menu.
   */
  /**
   * The error's explanation, under the cell that is showing it.
   *
   * Under the cell rather than in a panel, because a panel is
   * somewhere else and the question is about *this* cell: an error
   * code is a diagnosis in five characters and the sentence that
   * makes it actionable belongs beside the five characters.
   *
   * Only for the selected cell, and only while it is not being
   * edited — a cell somebody is typing into is a cell they are
   * already fixing, and an explanation under the caret would cover
   * the sheet they are reading to decide what to type.
   */
  const explainPopup = (row: number, column: number): UiElement => {
    const current = explain.value;
    const blamed = current?.blame == null ? '' : ` Made in ${current.blame}.`;
    return Box(
      {
        key: 'explain',
        position: 'absolute',
        left: GUTTER_WIDTH + sheetWindow.offsetOf(column),
        top: heightNow(row),
        zIndex: 4,
        maxWidth: 320,
        backgroundColor: 'surface',
        borderColor: 'danger',
        borderWidth: 1,
        paddingLeft: 6,
        paddingRight: 6,
        paddingTop: 3,
        paddingBottom: 3,
        pointerEvents: 'none',
        role: 'status',
        label: `${current?.code ?? ''} explained`
      },
      Text({
        key: 'why',
        text: `${current?.code ?? ''} — ${current?.meaning ?? ''}${blamed}`,
        fontSize: 11,
        color: 'text',
        textWrap: 'word',
        maxLines: 3
      })
    );
  };

  /**
   * A cell's note, beside it.
   *
   * To the right of the cell rather than under it, as Excel's are, so
   * it does not cover the row below — which is usually the row being
   * read. For the cell under the pointer, or else for the selected cell,
   * so the keyboard reaches a note as well as the mouse does.
   */
  const notePopup = (row: number, column: number): UiElement =>
    Box(
      {
        key: 'note',
        position: 'absolute',
        left: GUTTER_WIDTH + sheetWindow.offsetOf(column) + sheetWindow.widthOf(column) + 4,
        top: 0,
        zIndex: 4,
        maxWidth: 260,
        backgroundColor: 'surface',
        borderColor: 'primary',
        borderWidth: 1,
        paddingLeft: 6,
        paddingRight: 6,
        paddingTop: 4,
        paddingBottom: 4,
        pointerEvents: 'none',
        role: 'status',
        label: `Note on ${columnName(column)}${row + 1}`
      },
      Text({ key: 'said', text: noteOf(row, column), fontSize: 11, color: 'text', textWrap: 'word', maxLines: 8 })
    );

  /** The cell the pointer is resting on, when it has a note. */
  let hoveredNote: { row: number; column: number } | null = null;
  const hoverAt = (at: { row: number; column: number } | null): void => {
    const next = at !== null && noteOf(at.row, at.column) !== '' ? at : null;
    if (sameCell(next, hoveredNote)) {
      return;
    }
    hoveredNote = next;
    sheetWindow.invalidate();
  };

  const hintPopup = (row: number, column: number): UiElement => {
    const current = hint.value;
    const caret = editing.caretRectOf(editorNode);
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(column) + (caret?.x ?? 0);
    const common = {
      key: 'hint',
      position: 'absolute' as const,
      left,
      top: heightNow(row),
      zIndex: 4,
      backgroundColor: 'surface',
      borderColor: 'border',
      borderWidth: 1,
      paddingTop: 2,
      paddingBottom: 2,
      pointerEvents: 'none' as const
    };

    if (current?.kind === 'signature') {
      const marked = markedArgument(current.signature, current.argument);
      return Box(
        { ...common, paddingLeft: 6, paddingRight: 6, role: 'status', label: `${current.name} signature` },
        Text({
          key: 'sig',
          // The argument being filled in is the only one in the theme's
          // text colour; the rest are muted. Bold would move the
          // glyphs, and a hint that reflows as the caret crosses a
          // comma is a hint that draws the eye for the wrong reason.
          spans: [
            { text: `${current.name}(` },
            ...current.signature.args.flatMap((argument, at) => [
              { text: at === 0 ? '' : ', ' },
              { text: argument, color: at === marked ? 'text' : 'textMuted' }
            ]),
            { text: ')' }
          ],
          fontSize: 11,
          color: 'textMuted',
          textWrap: 'none'
        }),
        Text({ key: 'summary', text: current.signature.summary, fontSize: 11, color: 'textMuted', textWrap: 'none' })
      );
    }

    if (current?.kind !== 'completions') {
      return Box({ ...common, width: 0, height: 0 });
    }

    /**
     * At most eight, because a list longer than the screen is a list
     * that covers the sheet somebody is reading to decide what to
     * type.
     */
    const shown = current.names.slice(0, 8);
    const picked = Math.min(chosen.value, shown.length - 1);
    return Box(
      { ...common, role: 'listbox', label: 'Functions' },
      ...shown.map((name, at) =>
        Text({
          key: name,
          text: name,
          role: 'option',
          label: name,
          states: at === picked ? ['selected'] : [],
          fontSize: 12,
          paddingLeft: 6,
          paddingRight: 12,
          color: 'text',
          backgroundColor: at === picked ? 'selectionBackground' : undefined,
          textWrap: 'none'
        })
      )
    );
  };

  /**
   * The values a cell may take, under it.
   *
   * The same shape the function hints use — a listbox hanging off the
   * row that holds the open cell, so it travels with a scroll and
   * needs nothing kept in step.
   */
  const choicePopup = (row: number, column: number): UiElement => {
    const shown = choices.value.slice(0, 8);
    const picked = Math.min(choice.value, shown.length - 1);
    return Box(
      {
        key: 'choices',
        position: 'absolute',
        left: GUTTER_WIDTH + sheetWindow.offsetOf(column),
        top: heightNow(row),
        zIndex: 4,
        backgroundColor: 'surface',
        borderColor: 'border',
        borderWidth: 1,
        paddingTop: 2,
        paddingBottom: 2,
        pointerEvents: 'none',
        role: 'listbox',
        label: 'Allowed values'
      },
      ...shown.map((value, at) =>
        Text({
          key: value,
          text: value,
          role: 'option',
          label: value,
          states: at === picked ? ['selected'] : [],
          fontSize: 12,
          paddingLeft: 6,
          paddingRight: 12,
          color: 'text',
          backgroundColor: at === picked ? 'selectionBackground' : undefined,
          textWrap: 'none'
        })
      )
    );
  };

  /** How many rows the sheet has, for the anchor search to bound itself. */
  const rowCount = internalState(0);
  /**
   * Which charts hang off which row, rebuilt when either could have
   * moved.
   *
   * A map rather than a scan per row: a render walks thirty rows and
   * scanning the chart list in each of them is thirty times the work
   * to answer a question that changes when somebody drags something.
   */
  const anchored = new Map<number, SheetChart[]>();

  /**
   * The row a sheet-pixel offset falls in.
   *
   * A binary search over `rowOffsetOf`, which is monotonic and
   * already accounts for hidden rows and the heights that differ.
   * Twenty probes for a sheet of a million rows, and it runs when the
   * charts move rather than per frame.
   */
  const rowAtOffset = (y: number): number => {
    if (sheetWindow === undefined) {
      return 0;
    }
    let low = 0;
    let high = Math.max(0, rowCount.value - 1);
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (sheetWindow.rowOffsetOf(middle) <= y) {
        low = middle;
      } else {
        high = middle - 1;
      }
    }
    return low;
  };

  /**
   * Set when the anchors were asked for before the window could answer.
   *
   * A grid mounted over a workbook that is already open — the route
   * moving from `/` to `/d/main` mounts a second one — hears the
   * geometry and the charts before the window has been told how many
   * rows there are. Until it has, every row starts at offset 0, the
   * search above lands every chart on the last row, and the charts
   * hang off a row nobody will scroll to. So the map waits, and is
   * built the first time a render asks for it with the window caught
   * up.
   */
  let anchorsStale = false;

  const rebuildAnchors = (): void => {
    anchored.clear();
    if (sheetWindow !== undefined && rowCount.value > 1 && sheetWindow.rowOffsetOf(rowCount.value) === 0) {
      anchorsStale = charts.value.entries.length > 0;
      return;
    }
    anchorsStale = false;
    for (const chart of charts.value.entries) {
      const row = rowAtOffset(chart.y);
      const held = anchored.get(row);
      if (held === undefined) {
        anchored.set(row, [chart]);
      } else {
        held.push(chart);
      }
    }
  };

  /** The anchors, built now if they were asked for too early. */
  const anchors = (): Map<number, SheetChart[]> => {
    if (anchorsStale) {
      rebuildAnchors();
    }
    return anchored;
  };

  const chartHandlers = {
    select: (id: number) => sheet.send.selectChart(id),
    begin: (id: number, corner: Corner | null, x: number, y: number) => {
      const chart = charts.value.entries.find(entry => entry.id === id);
      if (chart === undefined) {
        return;
      }
      chartFrom = { x: chart.x, y: chart.y, width: chart.width, height: chart.height, pointerX: x, pointerY: y };
      chartDrag = { id, corner, x: chart.x, y: chart.y, width: chart.width, height: chart.height };
    },
    move: (x: number, y: number) => {
      if (chartDrag === null || chartFrom === null) {
        return;
      }
      const next = dragged(chartFrom, chartDrag.corner, x - chartFrom.pointerX, y - chartFrom.pointerY);
      chartDrag = { ...chartDrag, ...next };
      sheetWindow?.invalidate();
    },
    end: () => {
      if (chartDrag === null) {
        return;
      }
      const { id, x, y, width, height } = chartDrag;
      chartDrag = null;
      chartFrom = null;
      sheet.send.placeChart(id, x / zoom, y / zoom, width / zoom, height / zoom);
    }
  };

  /**
   * The charts anchored to a row, as its children.
   *
   * A chart hangs off the row it *starts* in, so it travels with the
   * scroll and its node is never rebuilt by one. `extendRange` below
   * keeps that row mounted while any part of the chart is on screen,
   * which is the same hook a merge reaching up out of the window
   * uses.
   */
  const chartsOn = (row: number): UiElement[] => {
    const here = anchors().get(row);
    if (here === undefined || sheetWindow === undefined) {
      return [];
    }
    const offset = sheetWindow.rowOffsetOf(row);
    return here.map(chart =>
      chartElement({
        chart,
        data: series.value.charts[String(chart.id)] ?? null,
        selected: charts.value.selected === chart.id,
        rowOffset: offset,
        drag: chartDrag,
        on: chartHandlers
      })
    );
  };

  const renderRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [rowHeader(row), ...chartsOn(row)];
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
    const finger = touching.value;
    if (corner && cornerAt !== null) {
      line.push(finger ? selectionHandle(row, cornerAt.column, 'end') : fillHandle(row, cornerAt.column));
    }
    const start = finger && startAt !== null && startAt.row === row;
    if (start && startAt !== null) {
      line.push(selectionHandle(row, startAt.column, 'start'));
    }
    // The hint hangs off the row holding the open cell, as the fill
    // handle hangs off the row holding the corner: a child of the row
    // travels with it through a scroll and needs nothing kept in step.
    const hinting = openAt !== null && openAt.row === row && hint.value !== null;
    if (hinting && openAt !== null) {
      line.push(hintPopup(row, openAt.column));
    }
    const choosing = openAt !== null && openAt.row === row && choices.value.length > 0;
    if (choosing && openAt !== null) {
      line.push(choicePopup(row, openAt.column));
    }
    // And the error's explanation, which belongs to the selected cell
    // rather than to an open one: a cell being typed into is a cell
    // somebody is already fixing.
    const explaining =
      openAt === null && explain.value !== null && latestSelection !== null && latestSelection.row === row;
    if (explaining && latestSelection !== null) {
      line.push(explainPopup(row, latestSelection.column));
    }
    // A note, for the cell the pointer rests on or else the selected
    // one — and never over an open cell or an error being explained.
    const noteAt =
      hoveredNote ??
      (latestSelection !== null && noteOf(latestSelection.row, latestSelection.column) !== ''
        ? { row: latestSelection.row, column: latestSelection.column }
        : null);
    const explainingThat = explaining && hoveredNote === null;
    if (openAt === null && noteAt !== null && noteAt.row === row && !explainingThat) {
      line.push(notePopup(row, noteAt.column));
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
    /**
     * A row holding something taller than itself is lifted.
     *
     * True of a vertical merge's anchor, and true of a chart for the
     * same reason and with the same consequence: the row below is
     * painted after this one, so without a stacking context the thing
     * reaching down out of this row is drawn and then covered up by
     * the rows it reaches over. A chart came out twenty-four pixels
     * tall in a browser, which is exactly the height of one row.
     */
    const spans =
      anchors().has(row) || merges.value.some(rect => rect.firstRow === row && rect.lastRow > row);
    return Row(
      {
        role: 'row',
        posInSet: row + 1,
        // Four coloured rectangles in the row's own paint pass when a
        // formula being typed names something on this row, and an
        // empty array the rest of the time.
        modifiers: outlineFor(row).modifiers,
        position: stuck ? 'sticky' : corner || start || spans || hinting || explaining ? 'relative' : undefined,
        top: stuck ? HEADER_HEIGHT + rowTop(row) : undefined,
        // A finger's handles reach over the rows around theirs, so their
        // rows are lifted over them.
        zIndex: stuck || spans || hinting || explaining || (finger && (corner || start)) ? 1 : undefined,
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
        role: 'button',
        label: SELECT_ALL,
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
        width: touching.pipe(map(GRIP)),
        height: HEADER_HEIGHT,
        position: 'absolute',
        right: touching.pipe(map(finger => -GRIP(finger) / 2)),
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
          // The grid sweeps a selection on a pan, and a column's letter
          // selects the column: a drag on its edge does neither.
          event.stopPropagation();
        },
        onPanMove: event => {
          if (resizing !== null) {
            event.stopPropagation();
            resizeTo(resizing.column, resizing.width + (event.x - resizing.from));
          }
        },
        // A click on the edge is not a click on the letter, and two of
        // them fit the column to what it holds — every selected column,
        // when this one is among them, as Excel does.
        onClick: (event: UiPointerEvent) => event.stopPropagation(),
        onDoubleClick: (event: UiPointerEvent) => {
          event.stopPropagation();
          const at = edit.selectionNow();
          const first = Math.min(cornerOf(at).column, at.anchorColumn);
          const last = Math.max(cornerOf(at).column, at.anchorColumn);
          const whole = Math.min(cornerOf(at).row, at.anchorRow) === 0 && Math.max(cornerOf(at).row, at.anchorRow) >= rowCount.value - 1;
          if (whole && column >= first && column <= last) {
            sheet.send.measureColumns(first, last);
          } else {
            sheet.send.measureColumns(column, column);
          }
        },
        onPanEnd: event => {
          event.stopPropagation();
          if (resizing !== null) {
            // The end of the drag, and the only moment the other
            // thread hears about it. Sending each frame would be a
            // round trip per pixel to agree on something this side has
            // already drawn; sending the result is what makes the
            // width survive a reload.
            sheet.send.setColumnWidth(resizing.column, unscaledWidth(widths.value[resizing.column] ?? COLUMN_WIDTH));
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
      rowHeights: heightsOf(hidden.value, sized.value),
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
        /**
         * A chart is drawn by the row it starts in, so a window that
         * starts below that row has nothing to draw it and the chart
         * vanishes while most of it is still on screen — which is the
         * merge problem exactly, and takes the same answer.
         *
         * Bounded by the chart's own height: a three-hundred-pixel
         * chart reaches back about twelve rows, and a sheet holds a
         * handful of charts.
         */
        for (const [row, here] of anchors()) {
          if (row >= firstRow) {
            continue;
          }
          const reaches = here.some(chart => rowAtOffset(chart.y + chart.height) >= range.firstRow);
          if (reaches) {
            firstRow = Math.min(firstRow, row);
          }
        }
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
      modifiers: [viewport.modifier, contextMenu({ onOpen: at => onContext(at) })],
      onKeyDown: onKey,
      // A press on a column's letter, a row's number or the corner. A
      // press on a cell is the cell's own click, and this lets it be.
      onClick: (event: UiPointerEvent) => {
        const hit = headerOfNode(event.target);
        if (hit !== null) {
          clickHeader(hit, event.modifiers.shift);
        }
      },
      // Sweeping a selection out with the pointer.
      //
      // On the grid rather than on every cell: three listeners instead
      // of three per cell, and the cell being swept over is found from
      // the window's offsets, which works for the ones that are not
      // mounted as well as the ones that are. A cell would have to be
      // under the pointer to hear about it, and past the edge of the
      // viewport none is.
      // Which hand the last press was made with; see `touching`.
      onPointerDown: (event: UiPointerEvent) => {
        const finger = event.pointer.kind === 'touch';
        if (finger !== touching.value) {
          touching.value = finger;
        }
      },
      onPanStart: (event: UiPointerEvent) => {
        // A finger's drag across the cells is a scroll, which the
        // engine's touch scroller takes; the selection is extended by
        // its handles instead.
        if (event.pointer.kind === 'touch') {
          return;
        }
        // A press that reached the grid is a press that missed every
        // chart, which is how a chart stops being selected.
        if (charts.value.selected !== 0) {
          sheet.send.selectChart(0);
        }
        // A press on the selection's border picks the block up: Phase
        // 20's cut and paste as one gesture, or a copy with Ctrl held.
        if (!edit.openNow() && onBorder(event)) {
          const at = cellUnder(event);
          const held = edit.selectionNow();
          if (at !== null) {
            carrying = {
              firstRow: Math.min(cornerOf(held).row, held.anchorRow),
              lastRow: Math.max(cornerOf(held).row, held.anchorRow),
              firstColumn: Math.min(cornerOf(held).column, held.anchorColumn),
              lastColumn: Math.max(cornerOf(held).column, held.anchorColumn),
              row: at.row,
              column: at.column
            };
            event.stopPropagation();
            return;
          }
        }
        // A drag across the letters or the numbers sweeps whole columns
        // or rows, as a drag across cells sweeps cells.
        const hit = headerOfNode(event.target) ?? headerHit(event);
        if (hit !== null) {
          selectHeader(hit, event.modifiers.shift);
          if (hit.kind !== 'corner') {
            const held = edit.selectionNow();
            headerSweep =
              hit.kind === 'column' ? { kind: 'column', from: held.anchorColumn } : { kind: 'row', from: held.anchorRow };
          }
          return;
        }
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
        if (carrying !== null) {
          const at = cellUnder(event);
          if (at === null || event.buttons === 0) {
            return;
          }
          const { rowCount: rows, columnCount: columns } = sheet.view.geometry.value;
          const height = carrying.lastRow - carrying.firstRow;
          const width = carrying.lastColumn - carrying.firstColumn;
          const firstRow = Math.min(Math.max(0, carrying.firstRow + at.row - carrying.row), rows - 1 - height);
          const firstColumn = Math.min(Math.max(0, carrying.firstColumn + at.column - carrying.column), columns - 1 - width);
          const next = { firstRow, lastRow: firstRow + height, firstColumn, lastColumn: firstColumn + width };
          if (landing === null || landing.firstRow !== next.firstRow || landing.firstColumn !== next.firstColumn) {
            landing = next;
            repaintOutlines();
          }
          return;
        }
        if (headerSweep !== null) {
          if (event.buttons === 0) {
            return;
          }
          const box = viewport.value;
          const x = event.x - box.x;
          const y = event.y - box.y;
          const held = edit.selectionNow();
          if (headerSweep.kind === 'column') {
            selectColumns(columnAtX(Math.max(GUTTER_WIDTH, x)), headerSweep.from, held);
          } else {
            selectRows(rowAtY(Math.max(HEADER_HEIGHT, y)), headerSweep.from, held);
          }
          return;
        }
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
      onPanEnd: (event: UiPointerEvent) => {
        if (carrying !== null) {
          const to = landing;
          const from = carrying;
          carrying = null;
          landing = null;
          repaintOutlines();
          if (to !== null && (to.firstRow !== from.firstRow || to.firstColumn !== from.firstColumn)) {
            // Ctrl at the drop, as in Excel: what the hand is holding
            // when it lets go is what it meant.
            sheet.send.moveRange(to.firstRow, to.firstColumn, event.modifiers.ctrl || event.modifiers.meta);
          }
          return;
        }
        if (sweeping || headerSweep !== null) {
          paintIfLit();
        }
        sweeping = false;
        pickingFrom = null;
        headerSweep = null;
      },
      // Text from the clipboard with no caret anywhere. Before the
      // engine offered this the paste was dropped: `paste` had nothing
      // editable to insert into and returned false.
      // A note shows while the pointer rests on its cell. A move with a
      // button down is a sweep and not a rest.
      onPointerMove: (event: UiPointerEvent) => {
        if (event.buttons === 0) {
          hoverAt(cellUnder(event));
          const over = !edit.openNow() && onBorder(event) ? 'move' : undefined;
          if (over !== pointer.value) {
            pointer.value = over;
          }
        }
      },
      cursor: pointer,
      onPointerLeave: () => hoverAt(null),
      onPaste: (event: UiPasteEvent) => {
        edit.pasteText(event.text, pasteMode);
        pasteMode = 'all';
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
  // A note that appears or goes under the selection changes what the
  // row draws beside it; see `notePopup`.
  ctx.effect(sheet.view.notes, () => sheetWindow.invalidate());

  // The window owns the offsets, so a new set of widths has to reach
  // it: everything past the column that moved sits somewhere else, and
  // the prefix sum is what says where.
  ctx.effect(widths, all => sheetWindow.setColumnWidths(all));

  // And the same for the rows, which the window needs before it can
  // place anything below a hidden one.
  ctx.effect(hidden, rows => sheetWindow.setRowHeights(heightsOf(rows, sized.value)));
  ctx.effect(sized, sizes => sheetWindow.setRowHeights(heightsOf(hidden.value, sizes)));

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
    const shown = geometry.columnWidths.map(scaledWidth);
    if (resizing === null && !sameWidths(shown, widths.value)) {
      widths.value = shown;
    }
    const rows = geometry.hiddenRows;
    if (rows.length !== hidden.value.size || rows.some(row => !hidden.value.has(row))) {
      hidden.value = new Set(rows);
    }
    const heights = geometry.rowHeights.map(([row, height]): [number, number] => [row, scaled(height)]);
    if (
      heightDrag === null &&
      (heights.length !== sized.value.size || heights.some(([row, height]) => sized.value.get(row) !== height))
    ) {
      sized.value = new Map(heights);
    }
    if (geometry.frozenRows !== frozen.value.rows || geometry.frozenColumns !== frozen.value.columns) {
      frozen.value = { rows: geometry.frozenRows, columns: geometry.frozenColumns };
    }
    if (geometry.rowCount !== rowCount.value) {
      rowCount.value = geometry.rowCount;
      rebuildAnchors();
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
   * The charts, and what they draw.
   *
   * Two effects because they are two keys, and the anchors are
   * rebuilt from the first: a chart that moved may have moved into
   * another row, and until the map says so it would go on being
   * drawn by the row it left.
   */
  ctx.effect(sheet.view.charts, value => {
    charts.value =
      zoom === 1
        ? value
        : {
            ...value,
            entries: value.entries.map(chart => ({
              ...chart,
              x: chart.x * zoom,
              y: chart.y * zoom,
              width: chart.width * zoom,
              height: chart.height * zoom
            }))
          };
    rebuildAnchors();
    sheetWindow?.invalidate();
    repaintOutlines();
  });
  ctx.effect(sheet.view.series, value => {
    series.value = value;
    repaintOutlines();
    // The rectangle did not move, so nothing has to be rebuilt: the
    // chart's `paint` inputs carry this object and the engine
    // repaints because it changed.
    sheetWindow?.invalidate();
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
    const kept = new Map<number, number>();
    let moved = false;
    for (const entry of autofit.columns) {
      let widest = 0;
      entry.samples.forEach((text, index) => {
        widest = Math.max(
          widest,
          measure.widthOf(text, { fontSize: dims.CELL_FONT_SIZE, fontWeight: entry.bold[index] ? 'bold' : 'normal' })
        );
      });
      // The padding a cell draws with, plus a hair so the widest
      // string is not flush against the gridline. Measured at 100%,
      // which is the size the document keeps.
      const wanted = widest === 0 ? dims.COLUMN_WIDTH : Math.ceil(widest) + dims.CELL_PADDING * 2 + 2;
      const clamped = Math.min(Math.max(wanted, dims.MIN_COLUMN_WIDTH), dims.MAX_COLUMN_WIDTH);
      if (next[entry.column] !== scaledWidth(clamped)) {
        next[entry.column] = scaledWidth(clamped);
        kept.set(entry.column, clamped);
        moved = true;
      }
    }
    if (!moved) {
      return;
    }
    widths.value = next;
    // The document is what keeps a width, so it has to hear about
    // this one the same way a drag tells it.
    for (const [column, width] of kept) {
      sheet.send.setColumnWidth(column, width);
    }
  });

  /**
   * Row fitting: the other half of the question autofit asks.
   *
   * The application worker sent the cells that could make a row
   * taller than one line; this thread measures them through the same
   * measurer the cells are laid out with — wrapped at the column's
   * width less its padding — and sends the heights back. A row whose
   * tallest cell fits in one line comes back at the default, which is
   * how a row whose wrapped text was emptied goes back down.
   *
   * Asked before the first frame there is no measurer yet, so the
   * question is held and tried again shortly rather than answered
   * with zeros.
   */
  let rowFitAnswered = 0;
  let rowFitRetry: ReturnType<typeof setTimeout> | null = null;
  const answerRowFit = (fit: SheetRowFit): void => {
    if (fit.serial <= rowFitAnswered) {
      return;
    }
    if (!measure.ready) {
      if (rowFitRetry === null) {
        rowFitRetry = setTimeout(() => {
          rowFitRetry = null;
          answerRowFit(latestRowFit);
        }, 50);
      }
      return;
    }
    rowFitAnswered = fit.serial;
    // The default row is one line of the default font and the room
    // round it; a taller one keeps the same room.
    // At 100%, which is the size the document keeps.
    const line = measure.measure({ text: 'X', fontSize: dims.CELL_FONT_SIZE, wrap: 'none' }).height;
    const room = Math.max(0, dims.ROW_HEIGHT - line);
    const heights: [number, number][] = fit.rows.map(entry => {
      let tallest = 0;
      for (const cell of entry.cells) {
        const fontSize = cell.fontSize === 0 ? dims.CELL_FONT_SIZE : cell.fontSize;
        const size = measure.measure({
          text: cell.text,
          fontSize,
          fontWeight: cell.bold ? 'bold' : 'normal',
          ...(cell.wrap ? { wrap: 'word' as const, maxWidth: Math.max(1, cell.width - dims.CELL_PADDING * 2) } : { wrap: 'none' as const })
        });
        tallest = Math.max(tallest, size.height);
      }
      return [entry.row, tallest === 0 ? dims.ROW_HEIGHT : Math.max(dims.ROW_HEIGHT, Math.ceil(tallest + room))];
    });
    sheet.send.fitRows(fit.serial, heights);
  };
  let latestRowFit: SheetRowFit = { serial: 0, rows: [] };
  ctx.effect(sheet.view.rowFit, fit => {
    latestRowFit = fit;
    answerRowFit(fit);
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

  /**
   * The round trip: the range the window settled on is what the
   * application worker is asked for, on the sheet it is drawn over.
   *
   * `range$` emits only when the range changes, so this is not a
   * command per frame. The sheet travels with it rather than on a
   * command of its own, which is what stops every other command
   * needing to say which sheet it means.
   */
  const askForWindow = (range: {
    firstRow: number;
    lastRow: number;
    firstColumn: number;
    lastColumn: number;
  }): void => {
    sheet.send.setViewport(
      sheet.view.sheets.value.active,
      range.firstRow,
      range.lastRow,
      range.firstColumn,
      range.lastColumn
    );
  };
  ctx.effect(sheetWindow.range$, askForWindow);

  /**
   * The copied block's outline, following the clipboard view: drawn
   * over the sheet it came from, and marching while it is up. The
   * timer runs only while there is an outline to move.
   */
  ctx.effect(combineLatest([sheet.view.clipboard, sheet.view.sheets]), ([clipboard, tabs]) => {
    const next = clipboard.marked !== null && clipboard.marked.sheet === tabs.active ? clipboard.marked : null;
    marked = next;
    if (next !== null && marching === null) {
      marching = setInterval(() => {
        marchPhase = (marchPhase + 1) % (DASH + GAP);
        repaintOutlines();
      }, 110);
    } else if (next === null && marching !== null) {
      clearInterval(marching);
      marching = null;
    }
    repaintOutlines();
  });
  ctx.onUnmount(() => {
    if (marching !== null) {
      clearInterval(marching);
    }
  });
  // New rows scrolled in need their share of the outline.
  ctx.effect(sheetWindow.range$, () => {
    if (marked !== null) {
      repaintOutlines();
    }
  });
  // A tab change is a new window over the same range, and the range
  // did not move — so `range$` says nothing and this has to ask.
  ctx.effect(sheet.view.sheets, () => askForWindow(sheetWindow.range$.value));

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
  let touchedBefore = false;
  // While a handle is being dragged the handles stay put: moving one
  // rebuilds it on another row, and the node the drag belongs to would
  // go with it. The selection's wash follows the finger meanwhile, and
  // the handles catch up when it lifts.
  ctx.effect(combineLatest([edit.selection, touching, handlesSettled]), ([at]) => {
    if (handleDrag !== null) {
      return;
    }
    const next = { row: Math.max(cornerOf(at).row, at.anchorRow), column: Math.max(cornerOf(at).column, at.anchorColumn) };
    const first = { row: Math.min(cornerOf(at).row, at.anchorRow), column: Math.min(cornerOf(at).column, at.anchorColumn) };
    const startMoved = !sameCell(first, startAt);
    startAt = first;
    if (sameCell(next, cornerBefore) && !startMoved && touchedBefore === touching.value) {
      return;
    }
    touchedBefore = touching.value;
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
    // A new edit places its caret once; see the editor's `ref`.
    caretPlaced = false;
    typed = null;
    offered = null;
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
  const bringIntoView = (at: SheetSelection): void => {
    const view = viewport.value;
    if (view.width === 0 || view.height === 0) {
      return;
    }
    const top = HEADER_HEIGHT + sheetWindow.rowOffsetOf(at.row);
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(at.column);
    const width = sheetWindow.widthOf(at.column);
    // The frozen strips cover the near edges, so a cell is only really
    // visible once it is past them.
    scrollY.value = bring(scrollY.value, top, sheetWindow.rowHeightOf(at.row), view.height, HEADER_HEIGHT);
    scrollX.value = bring(scrollX.value, left, width, view.width, GUTTER_WIDTH);
  };
  // What moved is what is brought into view: the corner, when Shift
  // moved it and left the active cell where it was; the active cell
  // otherwise — an arrow, a click, Enter walking the selection.
  let lastBrought: SheetSelection | null = null;
  ctx.effect(edit.selection, held => {
    const cornerMoved =
      lastBrought !== null &&
      lastBrought.row === held.row &&
      lastBrought.column === held.column &&
      (cornerOf(lastBrought).row !== cornerOf(held).row || cornerOf(lastBrought).column !== cornerOf(held).column);
    lastBrought = held;
    if (!keepScroll) {
      bringIntoView(cornerMoved ? { ...held, ...cornerOf(held) } : held);
    }
  });
  // A grid built again — a new zoom — starts at the top of the sheet,
  // and the selection is where somebody was looking. Once, when the
  // window first has a size to be scrolled in.
  let placedOnMount = false;
  ctx.effect(viewport, box => {
    if (!placedOnMount && box.width > 0 && box.height > 0) {
      placedOnMount = true;
      bringIntoView(edit.selectionNow());
      // And the keyboard, which went with the grid it was on: a zoom or
      // Show formulas chosen from the keyboard must leave the next key
      // on the sheet. Only for a grid built again, and only when nothing
      // else has the keyboard — the first grid takes nothing from anyone.
      if (_inputs.rebuilt.value === true && gridNode !== null && focus.focused.value === null) {
        focus.focus(gridNode);
      }
    }
  });

  /**
   * Cells that scrolled away, so their bindings go with them.
   *
   * **A frozen cell never scrolls away, and judging one by the
   * scrolling range says it always has.** `range$` is the range that
   * moves, and once a pane is frozen it *starts after the bands* —
   * rows 3 onwards and columns 1 onwards for a corner of three by one
   * — so every cell in a band fails this test on the first emission
   * after the freeze. It was then deleted from the cache and released
   * from both fan-outs while its element stayed mounted and on
   * screen, bound to a subject nothing would ever push to again: the
   * band drew at the right place, at the right size, in the right
   * colours, and empty.
   *
   * The application worker has always sent these cells — `rowsInView`
   * and `columnsInView` add the bands to the window for exactly this
   * reason — so the data was arriving and being thrown away on
   * receipt.
   *
   * Per axis, because the two are independent: a frozen row scrolled
   * sideways really does lose its right-hand cells, and a frozen
   * column scrolled down really does lose its lower ones. What it
   * cannot lose is the band itself.
   */
  ctx.effect(sheetWindow.range$, range => {
    const pane = frozen.value;
    for (const [key, mounted] of cells) {
      const rowShows = mounted.row < pane.rows || (mounted.row >= range.firstRow && mounted.row <= range.lastRow);
      const columnShows =
        mounted.column < pane.columns ||
        (mounted.column >= range.firstColumn && mounted.column <= range.lastColumn);
      if (!rowShows || !columnShows) {
        cells.delete(key);
        values.release(key);
        standings.release(key);
      }
    }
  });

  const menu = createComponent(Menu, {
    open: menuOpen,
    onOpenChange: (open: boolean) => {
      menuOpen.next(open);
      if (!open && gridNode !== null) {
        focus.focus(gridNode);
      }
    },
    items: menuItems,
    at: menuAt,
    label: 'Cell actions',
    onSelect: (value: string) => {
      menuOpen.next(false);
      if (gridNode !== null) {
        focus.focus(gridNode);
      }
      edit.runCommand(value as CommandId);
    }
  });

  return Box({ flex: 1, minHeight: 0, width: percent(100) }, grid, menu);
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
  const firstRow = Math.min(cornerOf(selection).row, selection.anchorRow);
  const lastRow = Math.max(cornerOf(selection).row, selection.anchorRow);
  const firstColumn = Math.min(cornerOf(selection).column, selection.anchorColumn);
  const lastColumn = Math.max(cornerOf(selection).column, selection.anchorColumn);
  return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
}


/** How the grid draws each place `guessOf` can answer. */
const AUTO_ALIGN: Readonly<Record<Place, 'right' | 'center' | 'start'>> = { end: 'right', center: 'center', start: 'start' };

export type { SheetWindow };

/** Shared, because the overwhelming majority of cells have no border. */
/** The dot in a cell's corner when it breaks its validation rule. */
const MARK_SIZE = 6;
/** How big a note's corner mark is. */
const NOTE_MARK = 5;

const NO_SHAPES: DecorationShape[] = [];

/** The corner above row 1, which selects the sheet. */
const SELECT_ALL = 'Select all';

/**
 * How thick a reference's outline is.
 *
 * Two pixels rather than one: it has to read as a deliberate mark
 * against the grid's own one-pixel rules, which are the same colour
 * family and everywhere.
 */
const OUTLINE = 2;
/** The selected chart's parts, in Excel's colours for them. */
const CHART_NAMES = '#ea4335';
const CHART_CATEGORIES = '#b061f5';
const CHART_VALUES = '#4285f4';

function sameWidths(a: readonly number[], b: readonly number[]): boolean {
  return a.length === b.length && a.every((width, index) => width === b[index]);
}

/** Whether two hints offer the same list of names, in the same order. */
function sameNames(before: FormulaHint, after: FormulaHint): boolean {
  if (before?.kind !== 'completions' || after?.kind !== 'completions') {
    return false;
  }
  return before.names.length === after.names.length && before.names.every((name, at) => name === after.names[at]);
}

/** Whether two hints would draw the same popup. */
function sameHint(before: FormulaHint, after: FormulaHint): boolean {
  if (before === null || after === null) {
    return before === after;
  }
  if (before.kind !== after.kind) {
    return false;
  }
  if (before.kind === 'signature' && after.kind === 'signature') {
    return before.name === after.name && before.argument === after.argument;
  }
  return sameNames(before, after);
}

/** Case folded, for comparing what was typed with what is offered. */
function fold(text: string): string {
  return text.toLocaleLowerCase();
}
