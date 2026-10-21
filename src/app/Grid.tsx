import { BehaviorSubject, combineLatest, type Observable } from 'rxjs';
import { distinctUntilChanged, map } from 'rxjs/operators';

import {
  Box,
  boxShadow,
  Button,
  Column,
  contextMenu,
  darkTheme,
  decorated,
  EditableText,
  editorFor,
  LazySheet,
  percent,
  Row,
  scrollPosition,
  Text,
  type DecorationShape,
  type UiElement,
  type UiKeyboardEvent,
  type UiNode,
  type UiPasteEvent,
  type UiModifier,
  type UiPointerEvent,
  type UiChild,
  type UiTextSpan,
  type UiTextChangeEvent,
  type SheetRange,
  type UiVirtualSheet
} from 'gesso-core';
import { Icon, Menu, type MenuItem } from 'gesso-components';
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

import { appearanceTheme } from './appearance';
import { inkFor, shownPalette } from './cellInk';
import { CHART_CATEGORIES, CHART_NAMES, CHART_VALUES } from './chartColours';
import { guessOf, type Place } from './alignment';
import { columnName, relativeRef } from '../sheet/A1';
import { chartElement, dragged, type ChartDrag, type Corner } from './ChartLayer';
import { acceptCompletion, definedFunctionsOf, hintFor, markedArgument, type DefinedFunction, type FormulaHint } from '../sheet/FormulaHint';
import type { Span } from '../sheet/Tokenizer';

import * as dims from './dimensions';
import type { CellEdge, CellPaint } from '../sheet/Format';
import {
  cellIn,
  type SheetUncertainty,
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
  type SheetRect,
  type SheetCharts,
  type SheetSelection,
  type SheetSeriesView,
  type SheetWindow
} from './SheetContract';
import { colouredReferences, formulaSpans, type ColouredReference } from './FormulaColours';
import { fontStack } from './fonts';
import { ICONS } from './icons';
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
 *   - **A cell that holds nothing is not a node.** Most of a sheet is
 *     empty, and an empty cell is a box with a grid line round it: the
 *     row draws the lines for all of its columns in its own paint pass
 *     (see `gridLines`), and the selection's wash is a rectangle or two
 *     per row over the top (see `washFor`). What gets a node is a cell
 *     with something to show — a value, a format, a note, a rule's mark,
 *     a merge — and the active cell, which carries the ring and is what
 *     a screen reader is told the cursor is on. Counted in Chrome on the
 *     proof's Sales sheet scrolled past its orders, this took the cells
 *     on screen from 722 to 14 — the frozen headings and the active one
 *     — and the thumb drag's median frame from about 4.7ms to 2.6ms.
 */

/**
 * A cell on screen: the element, and the subjects that feed it.
 *
 * The row and column are kept beside them because the feeds walk this
 * map every time the window or the selection moves, and parsing them
 * back out of the key was the only reason the key had a shape.
 */
interface MountedCell {
  /**
   * The cell it shows. Fixed for a cell kept by what it shows; moved,
   * by `repoint`, for one kept by its place on screen.
   */
  row: number;
  column: number;
  /**
   * The row slot it is kept in, or undefined for a cell kept by the cell
   * it shows: a frozen row's, a merge's, and every cell of the compare
   * pane. See `cell`.
   */
  readonly slot: number | undefined;
  /**
   * Where a click on it lands, read when the click happens rather than
   * captured when the cell was built, so a cell moved to another row
   * selects the row it shows now.
   */
  readonly at: { row: number; column: number };
  /** Its own height for a cell kept by slot; a row's for one kept by what it shows. */
  readonly height: BehaviorSubject<number> | null;
  readonly element: UiElement;
  readonly value: FanCell<string | null, { row: number; column: number }>;
  /**
   * Whether it was built as the active cell, with the ring. Decided when
   * it is built, like being stuck or merged: there is one active cell, it
   * moves at a person's pace, and a cell that stops being it is built
   * again rather than every cell binding to whether it is.
   */
  readonly active: boolean;
  /**
   * The row line that holds it now, so a line that lets go of a cell
   * does not release one another line has since taken up. Only a merge's
   * anchor, kept by the cell it shows, can move between lines: its row
   * scrolls away from one slot and comes back in another.
   */
  owner: string;
  /**
   * How the cell is painted, pushed in like the value.
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

/**
 * The selection's wash, laid over the cells rather than painted into
 * them, so it is see-through: the theme's `selectionBackground` over the
 * theme's own background, and a tint over a fill, as a spreadsheet draws
 * a range. One per theme because a decoration-free box cannot name a
 * palette entry with an alpha; the two are the theme's opaque wash
 * worked back to the blue it is made of.
 */
const SELECTED_WASH = { light: 'rgba(59, 130, 246, 0.16)', dark: 'rgba(59, 130, 246, 0.33)' } as const;
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
  _inputs: Inputs<{ editing: SheetEditing; zoom?: number; widen?: number; fonts?: boolean; rebuilt?: boolean }>,
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
  /**
   * Whether any cell in the workbook has a font of its own — Phase 40.
   *
   * Read once, like the zoom, and the grid is built again the moment
   * the first font arrives. A cell's font is one more binding per cell,
   * and a binding per cell is the cost `fanOut` was written about — so
   * a workbook with no fonts, which is every workbook the proof
   * scrolls, does not pay it.
   */
  const fonts = _inputs.fonts.value === true;
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
   *
   * **The selection no longer reaches the cells at all.** It had a
   * fan-out of its own, a standing per cell that chose its wash and its
   * ring; with empty cells drawn by their row there is no cell to wash,
   * so the wash is laid over the rows (`washFor`) and the ring is the
   * active cell's own border, a cell built for being active.
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

  /**
   * Whether the window exists yet, and so whether the rows can be asked
   * which cells they hold. The effects below run once as they are
   * registered, which is before the window — and before most of what a
   * row is built from — has been made.
   */
  let ready = false;
  /** The active cell, as a merge's anchor where it is in one; worked out when first asked. */
  let activeAt: { row: number; column: number } | null = null;

  // After the fan-out, which has already pushed the new values into the
  // cells there are: what is left is which cells there should be.
  ctx.effect(window$, current => {
    latestWindow = current;
    if (ready) {
      refreshLines();
    }
  });

  ctx.effect(selection$, selection => {
    latestSelection = selection;
    activeAt = null;
    if (ready) {
      refreshLines();
    }
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
  let latestUncertainty: SheetUncertainty = { cells: {} };
  /** Whether the theme is dark: the fills below and the charts are drawn for it. */
  const dark$ = appearanceTheme(ctx).pipe(
    map(theme => theme === darkTheme),
    distinctUntilChanged()
  );
  /**
   * The palette as the theme draws it: in the dark theme a light fill
   * is shown as its dark equivalent (see `shownFill`), and the ink is
   * chosen against what is shown, because `inkFor` is asked of the
   * paint the cell holds. Done to the palette — a handful of entries —
   * and not per cell, so a theme change is one repaint and no cell
   * binds to the theme.
   */
  const shownPalette$ = combineLatest([sheet.view.palette, dark$]).pipe(map(([palette, dark]) => ({ entries: shownPalette(palette.entries, dark) })));

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
    // A cell given a format, a note or a rule's mark has something to
    // show, and one that lost the last of them may have nothing left.
    if (ready) {
      refreshLines();
    }
  };

  /**
   * A cell's borders, recomputed only when they could have changed.
   *
   * `bordersOf` hands back the *same* empty array for a cell with no
   * borders, so the comparison below is a pointer check that is true
   * for almost every cell on the screen and pushes nothing at all.
   */
  /** Everything drawn in a cell's own paint pass: its histogram, under its borders and marks. */
  const shapesFor = (row: number, column: number, paint: CellPaint): readonly DecorationShape[] => {
    const box = shapeBox(row, column);
    if (box === null) {
      return NO_SHAPES;
    }
    const borders = bordersOf(paint, box.width, box.height, isFlagged(row, column), noteOf(row, column) !== '');
    const bars = latestUncertainty.cells[row]?.[column];
    return bars === undefined ? borders : [...histogramOf(bars, box.width, box.height), ...borders];
  };

  const refreshShapes = (mounted: MountedCell): void => {
    const next = shapesFor(mounted.row, mounted.column, mounted.paint.value);
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
   * Where each visible cell's value fell across the last simulation,
   * as bar heights from 0 to 8; see `SheetUncertainty`. Drawn under the
   * cell's figure, among its decorations, so a cell with a spread costs
   * sixteen rectangles in its own paint pass and a cell without one
   * costs nothing.
   */
  ctx.effect(sheet.view.uncertainty, current => {
    const before = latestUncertainty;
    latestUncertainty = current;
    for (const mounted of cells.values()) {
      if (before.cells[mounted.row]?.[mounted.column] !== current.cells[mounted.row]?.[mounted.column]) {
        refreshShapes(mounted);
      }
    }
    if (ready) {
      refreshLines();
    }
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

  ctx.effect(shownPalette$, current => {
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
    // Every cell with a node asks, now that a cell is placed by its own
    // `left`, so the sum is kept per set of widths rather than walked.
    const starts = prefixOf(widths.value);
    if (column < starts.length) {
      return starts[column];
    }
    return starts[starts.length - 1] + (column - starts.length + 1) * COLUMN_WIDTH;
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
    // A cell kept by slot holds its own height, for whichever row it
    // shows; the subjects above are only for cells kept by row.
    for (const mounted of cells.values()) {
      if (mounted.height !== null) {
        const next = heightNow(mounted.row);
        if (next !== mounted.height.value) {
          mounted.height.next(next);
          moved.add(mounted.row);
        }
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
  /**
   * Where each column starts, as a subject per column, for the same
   * reason as its width.
   *
   * A cell is placed in its row by its own `left` rather than by the
   * cells before it, because the cells before it are mostly not there:
   * an empty cell has no node to push the next one along. So a column
   * that is dragged wider moves every cell to its right, which is the
   * cells with nodes in the columns after it and nothing else. Two per
   * column, because a frozen column's cell is placed inside the frozen
   * strip and the rest inside the row, which starts a gutter earlier.
   */
  const columnLefts = new Map<string, { readonly column: number; readonly from: number; readonly left: BehaviorSubject<number> }>();
  const leftOf = (column: number, stuck: boolean): Observable<number> => {
    const key = `${stuck ? 'f' : 's'}${column}`;
    let entry = columnLefts.get(key);
    if (entry === undefined) {
      const from = stuck ? 0 : GUTTER_WIDTH;
      entry = { column, from, left: new BehaviorSubject(columnLeft(column) + from) };
      columnLefts.set(key, entry);
    }
    return entry.left;
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
    const starts = prefixOf(all);
    for (const { column, from, left } of columnLefts.values()) {
      const next = (starts[column] ?? columnLeft(column)) + from;
      if (next !== left.value) {
        left.next(next);
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
  /**
   * Whether the charts are drawn for the dark theme. Read when a row
   * builds its charts, and the rows are built again when it changes,
   * so a chart's picture is made again in the new theme at once.
   */
  let chartsDark = false;
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

  /**
   * The range a chart's corner handle is being dragged to, while it is:
   * drawn as one outline in place of the chart's three, and sent as a
   * command when the drag ends.
   */
  let chartRangeDraft: SheetRect | null = null;
  /**
   * Where the handle was when the drag began, which is where it stays
   * until it ends: a handle moved to another row is a new node, and a
   * drag does not survive its node being replaced.
   */
  let chartRangeFrom: SheetRect | null = null;

  /** The whole of what a chart reads: the three parts together. */
  const sourceRect = (source: SheetChartSource): SheetRect | null => {
    const parts = [source.names, source.categories, source.values].filter((part): part is SheetRect => part !== null);
    if (parts.length === 0) {
      return null;
    }
    return {
      firstRow: Math.min(...parts.map(part => part.firstRow)),
      lastRow: Math.max(...parts.map(part => part.lastRow)),
      firstColumn: Math.min(...parts.map(part => part.firstColumn)),
      lastColumn: Math.max(...parts.map(part => part.lastColumn))
    };
  };

  /** The selected cell's formula, for View ▸ Show references; see `paintOutlines`. */
  let shownFormula = '';
  /** The sheet showing, by name, so a reference to another sheet is not drawn on this one. */
  let activeSheetName = '';

  /**
   * The references to outline: the draft's while a cell is being typed
   * into, and otherwise the selected cell's when Show references is on.
   * Only the ones on the sheet in view, because a box round `B2` of this
   * sheet for `Sales!B2` would point at the wrong cell.
   */
  const outlinedReferences = (draft: string | null): readonly ColouredReference[] => {
    const text = draft ?? (edit.referencesShown.value && shownFormula.startsWith('=') ? shownFormula : null);
    if (text === null) {
      return [];
    }
    const here = activeSheetName.toUpperCase();
    return colouredReferences(text).filter(reference => {
      const sheet = reference.range.start.sheet;
      return sheet === undefined || sheet.toUpperCase() === here;
    });
  };

  const paintOutlines = (draft: string | null): void => {
    const references = outlinedReferences(draft);
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
      const parts: readonly (readonly [SheetRect | null, string])[] =
        chartRangeDraft !== null
          ? [[chartRangeDraft, CHART_VALUES]]
          : [
              [source.names, CHART_NAMES],
              [source.categories, CHART_CATEGORIES],
              [source.values, CHART_VALUES]
            ];
      for (const [part, color] of parts) {
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
    const name = tabs.entries[tabs.active]?.name ?? '';
    if (tabs.active !== activeSheet || name !== activeSheetName) {
      activeSheet = tabs.active;
      activeSheetName = name;
      repaintOutlines();
    }
  });
  ctx.effect(edit.referencesShown, repaintOutlines);
  ctx.effect(sheet.view.editor, editor => {
    if (editor.input !== shownFormula) {
      shownFormula = editor.input;
      if (edit.referencesShown.value) {
        repaintOutlines();
      }
    }
  });

  const buildCell = (
    at: { readonly row: number; readonly column: number },
    value: Observable<string | null>,
    paint: Observable<CellPaint>,
    shapes: Observable<readonly DecorationShape[]>,
    slotHeight: Observable<number> | null,
    active: boolean
  ): UiElement => {
    // Read once, at build time: a cell kept by slot is only ever moved to
    // another row in its own column, and never into or out of a merge
    // (see `cell`), so nothing below changes when it is.
    const { row, column } = at;
    /**
     * A frozen column's cells live in the row's frozen strip.
     *
     * The strip is what is stuck — `position: 'sticky'`, one gutter along,
     * the way the gutter has always been — and the cells in it are placed
     * against it, so a frozen cell is measured from the strip's edge and
     * not from the row's. Read once, at build time, because the cells are
     * rebuilt when the pane moves — see the effect that empties `cells` on
     * a freeze.
     */
    const stuck = column < frozen.value.columns;
    /**
     * A merged cell is drawn by its anchor and by nothing else.
     *
     * The anchor is as wide as the columns it covers and as tall as
     * the rows, and it overflows its own row downwards to reach them
     * — rows are not merged, only cells are, so there is nothing else
     * for a vertical merge to be. The cells it covers have no node at
     * all (see `needsNode`), and the anchor is opaque, so the grid lines
     * the rows draw under it are covered up.
     */
    const merge = mergeAt(row, column);
    const anchors = merge !== null && merge.firstRow === row && merge.firstColumn === column;
    const spanWidth = merge === null ? null : columnLeft(merge.lastColumn + 1) - columnLeft(merge.firstColumn);
    const spanRows = merge === null ? 1 : merge.lastRow - merge.firstRow + 1;

    return Text({
      key: column,
      /**
       * Placed by its own offset, and not by the cells before it.
       *
       * A row used to be a `Row` of every cell in its window, each pushing
       * the next along. Most of those were empty and are not there any
       * more, so a cell says where it is: absolute, against the row (or
       * the frozen strip), at the column's start. `leftOf` is a subject
       * per column, moved by a resize.
       */
      position: 'absolute',
      left: leftOf(column, stuck),
      top: 0,
      zIndex: anchors && spanRows > 1 ? 1 : undefined,
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
      // With no colour of its own, the ink follows the fill the cell is
      // showing: the theme's own where there is none. See `inkFor`. The
      // selection's wash is laid over the top and changes neither.
      color: combineLatest([value, paint]).pipe(
        map(([text, how]) => (text === null ? 'placeholder' : how.color !== '' ? how.color : inkFor(how.fill)))
      ),
      /**
       * See-through, unless it has a fill or anchors a merge.
       *
       * The grid lines are the row's, drawn under its cells, so a cell
       * with nothing behind its text lets them show and needs to paint
       * nothing of its own. A merge's anchor is opaque because the lines
       * of the columns and rows it covers run under it.
       */
      backgroundColor: paint.pipe(map(how => (how.fill !== '' ? how.fill : anchors ? 'background' : 'transparent'))),
      /**
       * The ring, on the active cell. It is the cell's own border, so it
       * is clipped and covered with the cell — by the gutter and the
       * frozen strip, which a ring drawn by the row would paint over.
       * Every other cell draws the grid line it sits on, which is where
       * the row draws it too, for a fill that would otherwise cover it.
       */
      borderColor: active ? 'primary' : GRID_LINE,
      borderWidth: active ? 2 : 1,
      width: anchors && spanWidth !== null ? spanWidth : widthOf(column),
      height:
        anchors && spanRows > 1
          ? heightOf(row).pipe(map(height => (height === 0 ? 0 : height * spanRows)))
          : (slotHeight ?? heightOf(row)),
      flexShrink: 0,
      paddingLeft: CELL_PADDING,
      paddingRight: CELL_PADDING,
      fontSize: paint.pipe(map(how => (how.fontSize === 0 ? CELL_FONT_SIZE : how.fontSize * zoom))),
      ...(fonts ? { fontFamily: paint.pipe(map(how => fontStack(how.fontFamily))) } : {}),
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
      // The active cell is the grid's active descendant, which is how a
      // screen reader hears which cell the keyboard is on: the grid keeps
      // focus, and a cell never takes it. See `activeCell$`.
      // Declared on every cell, so a cell that stops being the active one
      // hands its node back: a ref left out is a ref the builder keeps.
      ref: active ? activeRef() : undefined,
      onClick: (event: UiPointerEvent) => selectByPointer(at.row, at.column, event.modifiers.shift)
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
  /** And the names that are read rather than called: ranges, and formulas that are not functions. */
  let definedValues: readonly string[] = [];
  ctx.effect(sheet.view.names, current => {
    definedFunctions = definedFunctionsOf(current.formulas);
    const functions = new Set(definedFunctions.map(each => each.name));
    definedValues = [
      ...current.entries.map(entry => entry.name),
      ...current.formulas.filter(entry => !functions.has(entry.name)).map(entry => entry.name)
    ];
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

    const next = caret === undefined ? null : hintFor(draft, caret, definedFunctions, definedValues);
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
    column: ['cut', 'copy', 'paste', 'insertColumnLeft', 'insertColumnRight', 'deleteColumns', 'hideColumns', 'showColumns', 'autofitColumns', 'sortAscending', 'sortDescending', 'clear', 'freezeThroughColumns', 'unfreezeColumns'],
    row: ['cut', 'copy', 'paste', 'insertRowAbove', 'insertRowBelow', 'deleteRows', 'hideRows', 'showRows', 'fitRows', 'clear', 'freezeThroughRows', 'unfreezeRows']
  };
  /** Unfreezing is offered only when there is something to unfreeze. */
  const inMenu = (id: CommandId): boolean =>
    (id !== 'unfreezeRows' || frozen.value.rows > 0) && (id !== 'unfreezeColumns' || frozen.value.columns > 0);
  const openMenu = (kind: 'cell' | 'column' | 'row', at: { x: number; y: number }): void => {
    const items = MENU_FOR[kind].filter(inMenu).map(id => ({ value: id, label: COMMANDS[id].label }));
    menuItems.next(items);
    // Asked for where the pointer was: Gesso slides a menu that fits
    // on neither side of it until all of it is on the screen.
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
    // A click on a cell lets a selected chart go, as a sweep across the
    // cells already did: a chart's handles and its bar are about the
    // chart, and the person has gone back to the sheet.
    if (charts.value.selected !== 0) {
      sheet.send.selectChart(0);
    }
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

  /**
   * The cell on screen at `row` and `column`, for a cell that has a node.
   *
   * **Kept by its place on screen, where it can be.** The window mounts
   * each scrolled row in a slot (`recycleRows`), and a cell is kept by
   * that slot and its column, not by the cell it shows. A row arriving
   * in a slot gets the cells a row leaving it left behind, moved to show
   * the new row by `repoint`: their elements are the same objects, so
   * the reconciler passes over their nodes and their bindings stay
   * connected, and what the move costs is the values they push. Kept by
   * cell, a scrollbar drag replaced the whole window every frame — about
   * nine hundred cells and five thousand bindings built, and as many
   * torn down — and drew at a 20 ms median.
   *
   * A frozen row has no slot, and a cell in a merge is shaped by the
   * merge when it is built; both are kept by the cell they show, as
   * every cell used to be. Keys of the two kinds never meet: `s` marks
   * a slot's.
   *
   * Which cells get here at all is `needsNode`'s question, asked by the
   * row's line; `owner` is that line, so it knows what it holds.
   */
  const keyOf = (row: number, column: number, slot: number | undefined): string =>
    slot !== undefined && mergeAt(row, column) === null ? `s${slot}:${column}` : `${row}:${column}`;

  const cell = (row: number, column: number, slot: number | undefined, owner: string): UiElement => {
    if (openAt !== null && openAt.row === row && openAt.column === column) {
      return editorCell(row, column);
    }
    const key = keyOf(row, column, slot);
    const active = isActive(row, column);
    let mounted = cells.get(key);
    // The ring is built in, so a cell that became or stopped being the
    // active one is made again rather than moved.
    if (mounted !== undefined && mounted.active !== active) {
      release(key);
      mounted = undefined;
    }
    if (mounted !== undefined) {
      mounted.owner = owner;
      if (mounted.row !== row) {
        repoint(key, mounted, row);
      }
      return mounted.element;
    }
    // Seeded from what the window says *now*, because a cell is built
    // during the frame that reveals it and the feed above has already
    // run for this one.
    const kept = key.startsWith('s');
    const at = { row, column };
    const value = values.for(key, at);
    const paint = new BehaviorSubject<CellPaint>(paintOf(row, column));
    const shapes = new BehaviorSubject<readonly DecorationShape[]>(shapesFor(row, column, paint.value));
    const height = kept ? new BehaviorSubject(heightNow(row)) : null;
    const element = buildCell(at, value, paint, shapes, height, active);
    cells.set(key, { row, column, slot: kept ? slot : undefined, at, height, element, value, active, owner, paint, shapes });
    return element;
  };

  /** Lets a cell go: out of the cache, and off the window's feed. */
  const release = (key: string): void => {
    cells.delete(key);
    values.release(key);
  };

  /**
   * Moves a cell kept by slot to show another row of its column.
   *
   * Everything a cell shows that depends on its row is pushed in here,
   * through the subjects it was built on: its value through the
   * fan-out, which reads the new row and pushes only what differs; its
   * paint, its borders and marks, and its height. On a sheet nobody has
   * formatted most of that is a pointer comparison that pushes nothing.
   */
  const repoint = (key: string, mounted: MountedCell, row: number): void => {
    mounted.row = row;
    mounted.at.row = row;
    const at = { row, column: mounted.column };
    values.move(key, at);
    const paint = paintOf(row, mounted.column);
    if (paint !== mounted.paint.value) {
      mounted.paint.next(paint);
    }
    refreshShapes(mounted);
    const height = heightNow(row);
    if (mounted.height !== null && height !== mounted.height.value) {
      mounted.height.next(height);
    }
  };

  /** Whether a cell is the one the selection is on — a merge's anchor, when it is in one. */
  const isActive = (row: number, column: number): boolean => {
    if (latestSelection === null) {
      return false;
    }
    activeAt ??= anchorOf(latestSelection.row, latestSelection.column);
    return activeAt.row === row && activeAt.column === column;
  };

  /**
   * Whether a cell has to be a node.
   *
   * **Most of a sheet is empty, and an empty cell is drawn by its row.**
   * The row paints the grid lines of every column it holds (`gridLines`)
   * and the selection is laid over the rows (`washFor`), so a cell with
   * nothing in it, nothing on it and nothing round it has nothing left
   * to draw. Everything else is a node: a value, any format — a fill or
   * a border shows with nothing typed, and a font would the moment
   * something is — a note, a rule's mark, a spread from the simulation,
   * a merge's anchor, and the active cell, whose ring is its border and
   * which is what a screen reader is told the keyboard is on. A cell a
   * merge covers never is: the anchor is drawn over it.
   *
   * A row the window has not sent yet keeps whatever its slot holds.
   * Rows reach a slot a frame before their values reach the window, and
   * a slot that let its cells go when it was handed an unknown row would
   * build them all again a frame later: a thumb dragged through a full
   * sheet did exactly that, every frame. Kept, they show the new row's
   * nothing-yet until it arrives, as every cell did before.
   */
  const needsNode = (row: number, column: number, key: string): boolean => {
    const merge = mergeAt(row, column);
    if (merge !== null) {
      return merge.firstRow === row && merge.firstColumn === column;
    }
    if (isActive(row, column)) {
      return true;
    }
    const value = latestWindow === null ? null : cellIn(latestWindow, row, column);
    if (value === null) {
      return cells.has(key);
    }
    return (
      value !== '' ||
      paintOf(row, column) !== PLAIN_PAINT ||
      noteOf(row, column) !== '' ||
      isFlagged(row, column) ||
      latestUncertainty.cells[row]?.[column] !== undefined
    );
  };

  /**
   * What a mounted row holds, kept per slot like its cells.
   *
   * A row's cells are a stream of its own rather than children written
   * into the row, so a cell that gains a value — the window arriving a
   * frame after its row did, a format, a note — is one row's list pushed
   * again, and not every row on the screen built again to find it. The
   * window rebuilds a row when its columns or its slot's row change, and
   * `fillLine` runs then too; between those, `refreshLines` asks every
   * mounted row whether its list is still right and pushes the ones that
   * are not.
   *
   * It is also what lets a cell go. A cell leaves when its line stops
   * holding it — a column scrolled away, a value cleared — and not when
   * the scrolling range stops covering it, which was the old test and
   * which a frozen cell always failed: the range starts past the bands,
   * so every frozen cell was released while it was still on screen and
   * drew empty. A frozen cell is in its row's line like any other.
   */
  interface RowLine {
    readonly key: string;
    readonly slot: number | undefined;
    /** The row it holds now. */
    row: number;
    /** The cells with nodes in the scrolled columns, and the wash over them. */
    readonly cells: BehaviorSubject<readonly UiElement[]>;
    /** The same for the frozen columns, inside the strip. */
    readonly stuck: BehaviorSubject<readonly UiElement[]>;
    /** The keys of the cells it holds, so the ones it stops holding are let go. */
    keys: ReadonlySet<string>;
    /** The grid lines under the scrolled columns, and under the frozen ones. */
    readonly lines: BehaviorSubject<readonly DecorationShape[]>;
    readonly stripLines: BehaviorSubject<readonly DecorationShape[]>;
    /**
     * The row's modifiers: its lines, and the reference outlines of the
     * row it holds. Hoisted, for the reason `RowOutline` gives — a fresh
     * list in a render body re-attaches — and rebuilt when the row it
     * holds changes, since the outlines are kept by row.
     */
    readonly drawn: UiModifier<unknown>;
    modifiers: readonly UiModifier<unknown>[];
    readonly stripModifiers: readonly UiModifier<unknown>[];
    /** The wash last laid over it, by where it went, so the same wash is the same elements. */
    washed: string;
    wash: { readonly cells: readonly UiElement[]; readonly stuck: readonly UiElement[] };
  }

  const NO_WASH = { cells: [] as readonly UiElement[], stuck: [] as readonly UiElement[] };
  const rowLines = new Map<string, RowLine>();
  /** Which line each mounted row is in, written as the window builds it. */
  const lineOfRow = new Map<number, RowLine>();

  const lineFor = (row: number, slot: number | undefined): RowLine => {
    const key = slot === undefined ? `r${row}` : `s${slot}`;
    let line = rowLines.get(key);
    if (line === undefined) {
      const lines = new BehaviorSubject<readonly DecorationShape[]>(NO_SHAPES);
      const stripLines = new BehaviorSubject<readonly DecorationShape[]>(NO_SHAPES);
      const drawn = decorated(lines) as UiModifier<unknown>;
      line = {
        key,
        slot,
        row: -1,
        cells: new BehaviorSubject<readonly UiElement[]>([]),
        stuck: new BehaviorSubject<readonly UiElement[]>([]),
        keys: new Set<string>(),
        lines,
        stripLines,
        drawn,
        modifiers: [drawn],
        stripModifiers: [decorated(stripLines) as UiModifier<unknown>],
        washed: '',
        wash: NO_WASH
      };
      rowLines.set(key, line);
    }
    if (line.row !== row) {
      // The row it held has gone, so the map forgets it: it would
      // otherwise hold every row ever scrolled past.
      if (lineOfRow.get(line.row) === line) {
        lineOfRow.delete(line.row);
      }
      line.row = row;
      line.modifiers = [line.drawn, ...outlineFor(row).modifiers];
    }
    lineOfRow.set(row, line);
    return line;
  };

  /**
   * Works out what a row holds, and pushes what changed.
   *
   * Pushes nothing when the list comes out the same elements in the same
   * order, which on a scroll through an empty sheet is every row: no
   * cells, no wash, and the lines of a row as tall as the last one.
   */
  const fillLine = (line: RowLine, firstColumn: number, lastColumn: number): void => {
    const { row, slot, key: owner } = line;
    const keys = new Set<string>();
    const take = (column: number, into: UiElement[]): void => {
      const key = keyOf(row, column, slot);
      if (openAt !== null && openAt.row === row && openAt.column === column) {
        into.push(cell(row, column, slot, owner));
      } else if (needsNode(row, column, key)) {
        into.push(cell(row, column, slot, owner));
        keys.add(key);
      }
    };
    const stuck: UiElement[] = [];
    const shown: UiElement[] = [];
    const pane = frozen.value.columns;
    for (let column = 0; column < pane; column++) {
      take(column, stuck);
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      take(column, shown);
    }
    for (const key of line.keys) {
      if (!keys.has(key) && cells.get(key)?.owner === owner) {
        release(key);
      }
    }
    line.keys = keys;
    const wash = washFor(line, firstColumn, lastColumn);
    stuck.push(...wash.stuck);
    shown.push(...wash.cells);
    if (!sameElements(stuck, line.stuck.value)) {
      line.stuck.next(stuck);
    }
    if (!sameElements(shown, line.cells.value)) {
      line.cells.next(shown);
    }
    const height = heightNow(row);
    const lines = gridLines(height, firstColumn, lastColumn);
    if (lines !== line.lines.value) {
      line.lines.next(lines);
    }
    const stripLines = pane === 0 ? NO_SHAPES : gridLines(height, 0, pane - 1, 0);
    if (stripLines !== line.stripLines.value) {
      line.stripLines.next(stripLines);
    }
  };

  /**
   * Asks every mounted row whether it still holds the right cells.
   *
   * Run when something that decides it moved without the window building
   * its rows again: the window's values, the formats and the palette, the
   * notes and the rules' marks, a spread, the selection. Forty-odd rows of
   * twenty-odd lookups, almost all of which miss.
   */
  const refreshLines = (): void => {
    const range = sheetWindow.range$.value;
    const visit = (row: number): void => {
      const line = lineOfRow.get(row);
      if (line !== undefined && line.row === row) {
        fillLine(line, range.firstColumn, range.lastColumn);
      }
    };
    for (let row = 0; row < frozen.value.rows; row++) {
      visit(row);
    }
    for (let row = range.firstRow; row <= range.lastRow; row++) {
      visit(row);
    }
  };

  /**
   * The grid lines of a row, as rectangles in the row's own paint pass.
   *
   * What every cell used to draw for itself with a one-pixel border, and
   * the reason an empty cell no longer needs to exist: the same two lines
   * between each pair of columns — one at the right of each cell, one at
   * the left of the next — and one along the top and the foot. Drawn
   * before the row's children, so a cell with a fill covers them as it
   * always did, and a merge's anchor hides the ones inside it.
   *
   * The same array for every row of the same height with the same
   * columns, so a scroll through a sheet of default rows pushes nothing:
   * the comparison in `fillLine` is a pointer check. Built again when the
   * widths or the columns move, which rebuilds the rows anyway.
   */
  let linesFor: { widths: readonly number[]; firstColumn: number; lastColumn: number; from: number; byHeight: Map<number, readonly DecorationShape[]> }[] = [];
  const gridLines = (height: number, firstColumn: number, lastColumn: number, from = GUTTER_WIDTH): readonly DecorationShape[] => {
    if (height === 0 || lastColumn < firstColumn) {
      return NO_SHAPES;
    }
    const all = widths.value;
    let set = linesFor.find(each => each.firstColumn === firstColumn && each.lastColumn === lastColumn && each.from === from);
    if (set === undefined || set.widths !== all) {
      // One set for the scrolled columns and one for the frozen strip;
      // anything else is a window that has moved on.
      linesFor = linesFor.filter(each => each.widths === all && each.from !== from);
      set = { widths: all, firstColumn, lastColumn, from, byHeight: new Map() };
      linesFor.push(set);
    }
    let shapes = set.byHeight.get(height);
    if (shapes !== undefined) {
      return shapes;
    }
    const starts = prefixOf(all);
    const xOf = (column: number): number => from + (starts[column] ?? columnLeft(column));
    const left = xOf(firstColumn);
    const right = xOf(lastColumn + 1);
    const built: DecorationShape[] = [];
    const line = (x: number, y: number, width: number, tall: number): void => {
      built.push({ kind: 'fill', x, y, width, height: tall, radius: 0, color: GRID_LINE });
    };
    line(left, 0, right - left, 1);
    line(left, height - 1, right - left, 1);
    // A pixel at each side of each column, and the two that meet at a
    // boundary drawn as one rectangle two pixels wide.
    let edgeX = -1;
    let edgeWidth = 0;
    for (let column = firstColumn; column <= lastColumn; column++) {
      const width = all[column] ?? COLUMN_WIDTH;
      if (width <= 0) {
        continue;
      }
      const x = xOf(column);
      for (const at of [x, x + width - 1]) {
        if (edgeWidth > 0 && edgeX + edgeWidth === at) {
          edgeWidth += 1;
          continue;
        }
        if (edgeWidth > 0) {
          line(edgeX, 0, edgeWidth, height);
        }
        edgeX = at;
        edgeWidth = 1;
      }
    }
    if (edgeWidth > 0) {
      line(edgeX, 0, edgeWidth, height);
    }
    shapes = built;
    set.byHeight.set(height, shapes);
    return shapes;
  };

  /**
   * The selection's wash over one row: a see-through rectangle across
   * the selected columns, on top of the cells.
   *
   * **Laid over the cells rather than painted into them.** Each cell
   * used to take a standing in the selection and choose its background
   * from it, which put the selection's fan-out on every cell and needed
   * every selected cell to exist. Now the selection is drawn by the rows
   * it crosses: up to two rectangles each — the active cell is left out,
   * it keeps its own fill and carries the ring — and none at all for a
   * selection of one cell, which is nearly every selection there is.
   *
   * Boxes, not the row's decorations, because of what has to cover it.
   * The gutter and the frozen strip are stuck over the scrolled cells,
   * and a decoration on the row paints after every child of the row,
   * stuck ones included, so a wash scrolled under the gutter would be
   * drawn over the row numbers. A box sits in the row's paint order,
   * above the cells and under the strips. The frozen columns' part is a
   * box in the strip, placed from the strip's edge.
   */
  let washColour: string = SELECTED_WASH.light;
  const washFor = (line: RowLine, firstColumn: number, lastColumn: number): RowLine['wash'] => {
    const held = latestSelection;
    const row = line.row;
    const height = heightNow(row);
    if (held === null || height === 0) {
      return clearWash(line);
    }
    const corner = cornerOf(held);
    const top = Math.min(corner.row, held.anchorRow);
    const bottom = Math.max(corner.row, held.anchorRow);
    const first = Math.min(corner.column, held.anchorColumn);
    const last = Math.max(corner.column, held.anchorColumn);
    if (row < top || row > bottom) {
      return clearWash(line);
    }
    activeAt ??= anchorOf(held.row, held.column);
    const merge = mergeAt(activeAt.row, activeAt.column);
    const own = merge ?? { firstRow: activeAt.row, lastRow: activeAt.row, firstColumn: activeAt.column, lastColumn: activeAt.column };
    if (own.firstRow <= top && own.lastRow >= bottom && own.firstColumn <= first && own.lastColumn >= last) {
      // One cell, or one merge: the ring says it all.
      return clearWash(line);
    }
    const spans: [number, number][] =
      row >= own.firstRow && row <= own.lastRow
        ? [
            [first, Math.min(last, own.firstColumn - 1)],
            [Math.max(first, own.lastColumn + 1), last]
          ]
        : [[first, last]];
    const pane = frozen.value.columns;
    const pieces: { stuck: boolean; left: number; width: number }[] = [];
    for (const [from, to] of spans) {
      // The frozen part, in the strip.
      const a = from;
      const b = Math.min(to, pane - 1);
      if (b >= a) {
        pieces.push({ stuck: true, left: columnLeft(a), width: columnLeft(b + 1) - columnLeft(a) });
      }
      // And the scrolled part, as far as the window goes either way.
      const c = Math.max(from, pane, firstColumn);
      const d = Math.min(to, lastColumn);
      if (d >= c) {
        pieces.push({ stuck: false, left: GUTTER_WIDTH + columnLeft(c), width: columnLeft(d + 1) - columnLeft(c) });
      }
    }
    const said = `${washColour}|${height}|${pieces.map(piece => `${piece.stuck ? 'f' : 's'}${piece.left}+${piece.width}`).join(',')}`;
    if (said === line.washed) {
      return line.wash;
    }
    const box = (piece: { left: number; width: number }, index: number): UiElement =>
      Box({
        key: `wash-${index}`,
        position: 'absolute',
        left: piece.left,
        top: 0,
        width: piece.width,
        height,
        // Over the cells, and over a merge that reaches down out of its
        // row, which is lifted by the same amount and comes before it.
        zIndex: 1,
        backgroundColor: washColour,
        pointerEvents: 'none'
      });
    line.washed = said;
    line.wash = {
      cells: pieces.filter(piece => !piece.stuck).map(box),
      stuck: pieces.filter(piece => piece.stuck).map(box)
    };
    return line.wash;
  };
  const clearWash = (line: RowLine): RowLine['wash'] => {
    line.washed = '';
    line.wash = NO_WASH;
    return NO_WASH;
  };

  /**
   * The active cell's node, as the grid's active descendant.
   *
   * Each active cell is given its own ref, so the one that lets go of its
   * node clears the descendant only if it was still the one named: a new
   * active cell can be handed its node before the old one hands its back.
   */
  const activeCell$ = new BehaviorSubject<UiNode | null>(null);
  const activeRef = (): ((node: UiNode | null) => void) => {
    let mine: UiNode | null = null;
    return node => {
      if (node !== null) {
        mine = node;
        activeCell$.next(node);
      } else if (activeCell$.value === mine) {
        activeCell$.next(null);
      }
    };
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
   * Kept for the edit, and only for the edit: there is one of these at
   * a time, it is let go when a cell opens or closes, and in between its
   * row's list asks for it every time something else in the row moved.
   * Built again each time, every recalculation that republished the
   * window would hand the field a new element.
   */
  let editorBuilt: { readonly row: number; readonly column: number; readonly element: UiElement } | null = null;
  const editorCell = (row: number, column: number): UiElement => {
    if (editorBuilt === null || editorBuilt.row !== row || editorBuilt.column !== column) {
      editorBuilt = { row, column, element: buildEditor(row, column) };
    }
    return editorBuilt.element;
  };
  const buildEditor = (row: number, column: number): UiElement => {
    // The open cell is always a merge's anchor — `anchorOf` moves the
    // selection there before anything can open it — so this needs the
    // anchor's arithmetic and not the covered case.
    const merge = mergeAt(row, column);
    const spanWidth = merge === null ? null : columnLeft(merge.lastColumn + 1) - columnLeft(merge.firstColumn);
    const spanRows = merge === null ? 1 : merge.lastRow - merge.firstRow + 1;

    return EditableText({
      key: column,
      // Placed by its column, as every cell is; see `buildCell`.
      position: 'absolute',
      left: leftOf(column, column < frozen.value.columns),
      top: 0,
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
       * Built from `widthOf` alone, opening a merge collapsed the
       * anchor to one column — which on screen looks exactly like the
       * merge coming apart to show the original cells underneath it.
       * Same arithmetic as `buildCell` above, and for the same reason.
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
        // against it rather than against the layout root. Over the
        // cells, the wash over them and a merge reaching down from the
        // row above, which are all lifted to one.
        position: 'sticky',
        left: 0,
        zIndex: 2,
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
    const written = acceptCompletion(draft, span, name, !definedValues.includes(name));
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

  /**
   * The corner of the cells the selected chart reads, to drag: the
   * range grows or shrinks to the cell let go over, from the same top
   * left, as Excel's does. Blue, like the values it usually ends in.
   */
  const chartRangeHandle = (row: number, column: number): UiElement =>
    Box({
      key: 'chart-range',
      width: 8,
      height: 8,
      position: 'absolute',
      left: GUTTER_WIDTH + sheetWindow.offsetOf(column) + sheetWindow.widthOf(column) - 5,
      top: heightNow(row) - 5,
      zIndex: 3,
      backgroundColor: CHART_VALUES,
      borderColor: 'background',
      borderWidth: 1,
      cursor: 'nwse-resize',
      role: 'button',
      label: 'Chart range',
      onPointerDown: (event: UiPointerEvent) => event.stopPropagation(),
      onPanStart: (event: UiPointerEvent) => {
        const source = chartSource();
        chartRangeDraft = source === null ? null : sourceRect(source);
        chartRangeFrom = chartRangeDraft;
        event.stopPropagation();
      },
      onPanMove: (event: UiPointerEvent) => {
        if (chartRangeDraft === null) {
          return;
        }
        event.stopPropagation();
        const box = viewport.value;
        const to = sheetWindow.cellAt(event.x - box.x, event.y - box.y);
        const next = {
          ...chartRangeDraft,
          lastRow: Math.max(chartRangeDraft.firstRow, to.row),
          lastColumn: Math.max(chartRangeDraft.firstColumn, to.column)
        };
        if (next.lastRow !== chartRangeDraft.lastRow || next.lastColumn !== chartRangeDraft.lastColumn) {
          chartRangeDraft = next;
          repaintOutlines();
        }
      },
      onClick: (event: UiPointerEvent) => event.stopPropagation(),
      onPanEnd: (event: UiPointerEvent) => {
        event.stopPropagation();
        const draft = chartRangeDraft;
        chartRangeDraft = null;
        chartRangeFrom = null;
        if (draft !== null && charts.value.selected !== 0) {
          sheet.send.setChartRange(charts.value.selected, draft.firstRow, draft.firstColumn, draft.lastRow, draft.lastColumn);
        }
        repaintOutlines();
        sheetWindow.invalidate();
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
  /**
   * Whether a click landed on a cell that has no node — on its row, with
   * nothing of the row's own in between.
   *
   * The first node with a role on the way up decides: a cell's own click
   * has already been handled, and a chart, a handle, a popup or a header
   * is not a cell however much of the sheet it covers. The frozen strip
   * has no role, so an empty frozen cell is found through it.
   */
  const onEmptyCell = (target: UiNode | null): boolean => {
    for (let at: UiNode | null = target; at !== null && at !== gridNode; at = at.parent) {
      const role = at.properties.get('role');
      if (role !== undefined && role !== null) {
        return role === 'row' && at.properties.get('posInSet') !== undefined;
      }
    }
    return false;
  };

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
    const why = `${current?.code ?? ''} — ${current?.meaning ?? ''}${blamed}`;
    // Measured here rather than left to the panel, for the reason
    // `popupPanel` gives: under the row's bottom edge an unsized box is
    // laid out zero tall. At most three lines, as the text is.
    const line = measure.measure({ text: 'X', fontSize: 11, wrap: 'none' }).height;
    const said = measure.measure({ text: why, fontSize: 11, wrap: 'word', maxWidth: EXPLAIN_WIDTH - 12 });
    return Box(
      {
        ...popupPanel('explain', GUTTER_WIDTH + sheetWindow.offsetOf(column), heightNow(row), Math.ceil(Math.min(said.height, 3 * line)) + 2),
        width: Math.ceil(said.width) + 12,
        maxWidth: EXPLAIN_WIDTH,
        borderColor: 'danger',
        paddingLeft: 6,
        paddingRight: 6,
        paddingTop: 3,
        paddingBottom: 3,
        role: 'status',
        label: `${current?.code ?? ''} explained`
      },
      Text({
        key: 'why',
        text: why,
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
   *
   * Past the whole of a merge, not past its first column: a note on a
   * merged figure was drawn over the second half of the figure.
   */
  const notePopup = (row: number, column: number): UiElement =>
    Box(
      {
        key: 'note',
        position: 'absolute',
        left: GUTTER_WIDTH + sheetWindow.offsetOf(column) + (shapeBox(row, column)?.width ?? sheetWindow.widthOf(column)) + 4,
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
  const hintPopup = (row: number, column: number): UiElement => {
    const current = hint.value;
    const caret = editing.caretRectOf(editorNode);
    const left = GUTTER_WIDTH + sheetWindow.offsetOf(column) + (caret?.x ?? 0);

    if (current?.kind === 'signature') {
      const marked = markedArgument(current.signature, current.argument);
      return Column(
        {
          ...popupPanel('hint', left, heightNow(row), 2 * POPUP_LINE),
          paddingLeft: 6,
          paddingRight: 6,
          role: 'status',
          label: `${current.name} signature`
        },
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
          height: POPUP_LINE,
          verticalAlign: 'middle',
          color: 'textMuted',
          textWrap: 'none'
        }),
        Text({ key: 'summary', text: current.signature.summary, fontSize: 11, height: POPUP_LINE, verticalAlign: 'middle', color: 'textMuted', textWrap: 'none' })
      );
    }

    if (current?.kind !== 'completions') {
      return Box({ key: 'hint', position: 'absolute', left, top: heightNow(row), width: 0, height: 0, pointerEvents: 'none' });
    }

    /**
     * At most eight, because a list longer than the screen is a list
     * that covers the sheet somebody is reading to decide what to
     * type.
     */
    return listPopup('hint', 'Functions', left, heightNow(row), current.names.slice(0, 8), chosen.value);
  };

  /**
   * The values a cell may take, under it.
   *
   * The same shape the function hints use — a listbox hanging off the
   * row that holds the open cell, so it travels with a scroll and
   * needs nothing kept in step.
   */
  const choicePopup = (row: number, column: number): UiElement =>
    listPopup('choices', 'Allowed values', GUTTER_WIDTH + sheetWindow.offsetOf(column), heightNow(row), choices.value.slice(0, 8), choice.value);

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
        scrollX,
        on: chartHandlers,
        dark: chartsDark
      })
    );
  };

  /**
   * The frozen columns of one row, as one stuck strip.
   *
   * The frozen cells used to be stuck one by one, each a sticky node in
   * the row's flow. With empty cells gone there is nothing in the flow to
   * put a stuck cell after, and an empty frozen cell still has to cover
   * the scrolled cells passing under it — so the strip is what is stuck,
   * opaque, drawing the frozen columns' grid lines, and the frozen cells
   * with nodes are placed in it as the scrolled ones are in the row.
   */
  const frozenStrip = (row: number, line: RowLine): UiElement =>
    Box(
      {
        key: 'frozen',
        position: 'sticky',
        left: GUTTER_WIDTH,
        zIndex: 2,
        width: columnLeft(frozen.value.columns),
        height: heightNow(row),
        flexShrink: 0,
        backgroundColor: 'background',
        modifiers: line.stripModifiers
      },
      line.stuck
    );

  const renderRow = (row: number, firstColumn: number, lastColumn: number, slot?: number): UiElement => {
    const held = lineFor(row, slot);
    fillLine(held, firstColumn, lastColumn);
    /**
     * In-flow, the gutter, then the frozen strip, and nothing else.
     *
     * The window puts its leading spacer after the gutter and as many
     * children again as there are frozen columns (see `withLead`); with
     * the strip second, the spacer lands after it whatever the count,
     * and since the cells are absolute it pushes nothing along.
     */
    const line: UiChild[] = [rowHeader(row)];
    if (frozen.value.columns > 0) {
      line.push(frozenStrip(row, held));
    }
    line.push(held.cells, ...chartsOn(row));
    const corner = cornerAt !== null && cornerAt.row === row;
    const finger = touching.value;
    if (corner && cornerAt !== null) {
      line.push(finger ? selectionHandle(row, cornerAt.column, 'end') : fillHandle(row, cornerAt.column));
    }
    const start = finger && startAt !== null && startAt.row === row;
    if (start && startAt !== null) {
      line.push(selectionHandle(row, startAt.column, 'start'));
    }
    // The selected chart's range, by its corner.
    const reading = chartSource();
    const readRect = chartRangeFrom ?? (reading === null ? null : sourceRect(reading));
    const ranging = readRect !== null && readRect.lastRow === row;
    if (ranging && readRect !== null) {
      line.push(chartRangeHandle(row, readRect.lastColumn));
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
        // The grid lines under the cells, and four coloured rectangles
        // over them when a formula being typed names something on this
        // row; see `RowLine`.
        modifiers: held.modifiers,
        // Positioned always, now that every cell is placed against it.
        position: stuck ? 'sticky' : 'relative',
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
    /**
     * Whether this column is the one a filter was chosen by — Phase 41.
     * A funnel at the right of its letter, as every spreadsheet marks
     * a filtered column, and said in its name for a screen reader.
     */
    const filteredHere = sheet.view.status.pipe(
      map(current => current.filterColumn === column),
      distinctUntilChanged()
    );
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
        label: filteredHere.pipe(map(is => (is ? `${columnName(column)}, filtered` : columnName(column)))),
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
      Box(
        {
          key: 'filtered',
          position: 'absolute',
          right: 9,
          top: 6,
          opacity: filteredHere.pipe(map(is => (is ? 1 : 0)))
        },
        createComponent(Icon, {
          path: ICONS.filter.path,
          viewBox: ICONS.filter.viewBox,
          size: 12,
          color: 'primary',
          style: ICONS.filter.style,
          strokeWidth: 2,
          fillRule: 'nonzero'
        })
      ),
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

  /** Where the grid last moved the sheet itself, until a layout has reported it; see `scrollPosition` below. */
  let broughtTo: { x: number; y: number } | null = null;
  /**
   * Where the sheet is scrolled: written when the selection is brought
   * into view, and kept current by `scrollPosition` when anything else
   * scrolls it.
   */
  const scrollX = internalState(0);
  const scrollY = internalState(0);
  let gridNode: UiNode | null = null;
  /** The same node as a stream, for the second pane to scroll with; see `comparePane`. */
  const gridNode$ = new BehaviorSubject<UiNode | null>(null);
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
      // The whole of the box the grid is given, and not the rows it has
      // built so far: a scroller with no size of its own is its
      // content's size, which here is only the rows the window mounted
      // to fill the size it was going to be.
      height: percent(100),
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
      // Rows in slots, so a jump reuses the cells on screen; see `cell`.
      recycleRows: true,
      role: 'grid',
      label: 'Sheet',
      focusable: true,
      /**
       * The cell the keyboard is on, for a screen reader.
       *
       * The grid keeps focus and moves a selection inside itself, so
       * without this a reader hears the grid and nothing as the arrows
       * move. The active cell always has a node — it carries the ring —
       * so there is always something to name. The cells with nothing in
       * them are not in the tree at all: there is no node to put there,
       * and a grid of empty cells read out one by one is noise. A row
       * says which row it is (`posInSet`), and the formula bar and the
       * name box say where the cursor is.
       */
      activeDescendant: activeCell$,
      ref: node => {
        gridNode = node;
        gridNode$.next(node);
      },
      modifiers: [
        viewport.modifier,
        contextMenu({ onOpen: at => onContext(at) }),
        /**
         * Where the sheet really is, kept in the two states bound into it.
         *
         * The wheel, a trackpad and the scrollbar move the window without
         * writing these, so they said wherever the selection had last been
         * brought to, and everything that measures from them was off by
         * however far the sheet had been scrolled by hand since: a click on
         * a cell in plain view moved the sheet, and one after scrolling back
         * up put the cell at the top. The selection's border and the
         * popups hung off a cell are placed from them too. Gesso's own
         * answer for a bound offset, `scrollPosition`.
         */
        scrollPosition({
          onChange: offset => {
            // A move the grid made itself is written before the layout
            // that applies it reports: the report it skips is of the
            // layout before the move, which would put the sheet back.
            if (broughtTo !== null) {
              const brought = broughtTo;
              broughtTo = null;
              if (offset.x !== brought.x || offset.y !== brought.y) {
                return;
              }
            }
            scrollX.value = offset.x;
            scrollY.value = offset.y;
          }
        })
      ],
      onKeyDown: onKey,
      // A press on a column's letter, a row's number or the corner. A
      // press on a cell with a node is the cell's own click, and this
      // lets it be; one on a cell without — most of them — lands on its
      // row, and is worked out from where it was.
      onClick: (event: UiPointerEvent) => {
        const hit = headerOfNode(event.target);
        if (hit !== null) {
          clickHeader(hit, event.modifiers.shift);
          return;
        }
        if (onEmptyCell(event.target)) {
          const at = cellUnder(event);
          if (at !== null) {
            selectByPointer(at.row, at.column, event.modifiers.shift);
          }
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
  ready = true;
  // The charts' theme, once there is a window to rebuild; see `chartsDark`.
  // The selection's wash is drawn for it as well; see `SELECTED_WASH`.
  ctx.effect(dark$, dark => {
    washColour = dark ? SELECTED_WASH.dark : SELECTED_WASH.light;
    if (dark !== chartsDark) {
      chartsDark = dark;
      sheetWindow.invalidate();
    }
  });

  /**
   * Lets every cell and every row line go, for a change that decides how
   * a cell is built — the pane, a merge — so that each is built again.
   */
  const forgetAll = (): void => {
    cells.clear();
    values.releaseAll();
    rowLines.clear();
    lineOfRow.clear();
    editorBuilt = null;
  };
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
    forgetAll();
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
      forgetAll();
      activeAt = null;
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
   * written, so the open cell's row hands the editor out in its place
   * (see `fillLine`) and lets the cell it was showing go; closing it
   * builds the cell again. The rows are rebuilt as well, for the hints
   * and the lists that hang off the open cell's row, at the pace of a
   * cell being edited rather than per frame.
   *
   * The fill handle moves with the selection in the same way: it is a
   * child of the row holding the selection's corner, so moving the
   * corner rebuilds the rows. The cells themselves are not touched —
   * the handle hangs off the row, not off a cell — and the active cell
   * looks after itself, built again by its row when it stops or starts
   * being the one with the ring.
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
    openBefore = next;
    openAt = next;
    editorBuilt = null;
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
    const y = bring(scrollY.value, top, sheetWindow.rowHeightOf(at.row), view.height, HEADER_HEIGHT);
    const x = bring(scrollX.value, left, width, view.width, GUTTER_WIDTH);
    if (y !== scrollY.value || x !== scrollX.value) {
      broughtTo = { x, y };
    }
    scrollY.value = y;
    scrollX.value = x;
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

  // -------------------------------------------------------------------
  // The second pane
  // -------------------------------------------------------------------

  /**
   * Another version of the workbook beside this one, scrolled with it.
   *
   * Read-only and drawn from two keys of its own, `compareWindow` and
   * `compareFormats`, which the application worker fills for this
   * grid's viewport: the pane never asks for a window, because it is
   * always looking at the cells the grid is. It follows the grid's
   * scroll through `scrollWith`, so the two cannot drift, and it uses
   * the grid's widths, heights and frozen panes, so a row on the left is
   * the row on the right.
   *
   * Its own cells rather than the grid's, which carry the selection,
   * the editor, the fill handle and the charts — none of which mean
   * anything in a pane nobody types into. A merge spans its columns in
   * its own row, as the grid's does across; borders, notes, charts and
   * a rule's paint are the grid's alone.
   */
  const compare$ = sheet.view.compare;
  const compareValues = fanOut<SheetWindow, string | null, At>(
    sheet.view.compareWindow,
    (current, _key, at) => cellIn(current, at.row, at.column),
    { initial: null }
  );
  const comparePaintSource = combineLatest([sheet.view.compareFormats, shownPalette$]);
  const comparePaints = fanOut<
    [{ cells: Readonly<Record<string, Readonly<Record<string, number>>>> }, { entries: readonly CellPaint[] }],
    CellPaint,
    At
  >(
    comparePaintSource,
    ([formats, palette], _key, at) => {
      const id = formats.cells[at.row]?.[at.column] ?? 0;
      return id === 0 ? PLAIN_PAINT : (palette.entries[id] ?? PLAIN_PAINT);
    },
    { initial: PLAIN_PAINT }
  );
  const compareCells = new Map<string, { row: number; column: number; element: UiElement }>();

  const compareCell = (row: number, column: number): UiElement => {
    const key = `${row}:${column}`;
    const kept = compareCells.get(key);
    if (kept !== undefined) {
      return kept.element;
    }
    const at = { row, column };
    const value = compareValues.for(key, at);
    const paint = comparePaints.for(key, at);
    const stuck = column < frozen.value.columns;
    const merge = mergeAt(row, column);
    const anchorsHere = merge !== null && merge.firstRow === row && merge.firstColumn === column;
    const covered = merge !== null && !anchorsHere && row === merge.firstRow;
    const spanWidth = merge === null ? null : columnLeft(merge.lastColumn + 1) - columnLeft(merge.firstColumn);
    const element = Text({
      key: column,
      position: stuck ? 'sticky' : undefined,
      left: stuck ? GUTTER_WIDTH + columnLeft(column) : undefined,
      zIndex: stuck ? 1 : undefined,
      text: value.pipe(map(text => text ?? '')),
      color: combineLatest([value, paint]).pipe(
        map(([text, how]) => (text === null ? 'placeholder' : how.color !== '' ? how.color : inkFor(how.fill)))
      ),
      backgroundColor: paint.pipe(map(how => (how.fill === '' ? 'background' : how.fill))),
      borderColor: GRID_LINE,
      borderWidth: covered ? 0 : 1,
      width: covered ? 0 : anchorsHere && spanWidth !== null ? spanWidth : widthOf(column),
      height: heightOf(row),
      flexShrink: 0,
      paddingLeft: covered ? 0 : CELL_PADDING,
      paddingRight: covered ? 0 : CELL_PADDING,
      fontSize: paint.pipe(map(how => (how.fontSize === 0 ? CELL_FONT_SIZE : how.fontSize * zoom))),
      ...(fonts ? { fontFamily: paint.pipe(map(how => fontStack(how.fontFamily))) } : {}),
      textWrap: paint.pipe(map(how => (how.wrap ? 'word' : 'none'))),
      fontWeight: paint.pipe(map(how => (how.bold ? 'bold' : 'normal'))),
      fontStyle: paint.pipe(map(how => (how.italic ? 'italic' : 'normal'))),
      textOverflow: 'clip',
      verticalAlign: 'middle',
      textAlign: combineLatest([value, paint]).pipe(
        map(([text, how]) =>
          how.align === 'auto' ? AUTO_ALIGN[guessOf(text)] : how.align === 'center' ? 'center' : how.align
        )
      ),
      selectable: false,
      role: 'cell'
    });
    compareCells.set(key, { row, column, element });
    return element;
  };

  const compareGutter = (row: number): UiElement =>
    Box(
      {
        key: 'gutter',
        width: GUTTER_WIDTH,
        height: heightOf(row),
        flexShrink: 0,
        x: 'center',
        y: 'center',
        position: 'sticky',
        left: 0,
        zIndex: 1,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1,
        role: 'rowheader',
        label: String(row + 1)
      },
      Text({ text: String(row + 1), color: 'textMuted', fontSize: 11, textAlign: 'center', verticalAlign: 'middle', selectable: false })
    );

  const compareRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [compareGutter(row)];
    for (let column = 0; column < frozen.value.columns; column++) {
      line.push(compareCell(row, column));
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(compareCell(row, column));
    }
    const stuck = row < frozen.value.rows;
    return Row(
      {
        role: 'row',
        posInSet: row + 1,
        position: stuck ? 'sticky' : undefined,
        top: stuck ? HEADER_HEIGHT + rowTop(row) : undefined,
        zIndex: stuck ? 1 : undefined,
        backgroundColor: stuck ? 'background' : undefined,
        overflow: heightOf(row).pipe(map(height => (height === 0 ? 'hidden' : undefined)))
      },
      ...line
    );
  };

  const compareHeader = (firstColumn: number, lastColumn: number): UiElement => {
    const letter = (column: number): UiElement =>
      Text({
        key: column,
        text: columnName(column),
        width: widthOf(column),
        height: HEADER_HEIGHT,
        flexShrink: 0,
        position: column < frozen.value.columns ? 'sticky' : undefined,
        left: column < frozen.value.columns ? GUTTER_WIDTH + columnLeft(column) : undefined,
        zIndex: column < frozen.value.columns ? 1 : undefined,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1,
        color: 'textMuted',
        fontSize: 11,
        textAlign: 'center',
        verticalAlign: 'middle',
        selectable: false,
        role: 'columnheader',
        label: columnName(column)
      });
    const line: UiElement[] = [
      Text({
        key: 'corner',
        text: '',
        width: GUTTER_WIDTH,
        height: HEADER_HEIGHT,
        flexShrink: 0,
        position: 'sticky',
        left: 0,
        zIndex: 2,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1
      })
    ];
    for (let column = 0; column < frozen.value.columns; column++) {
      line.push(letter(column));
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(letter(column));
    }
    return Row({ position: 'sticky', top: 0, zIndex: 2 }, ...line);
  };

  let compareWindow: UiVirtualSheet | undefined;
  const comparePane = LazySheet(
    {
      flex: 1,
      minHeight: 0,
      width: percent(100),
      height: percent(100),
      backgroundColor: 'background',
      rowCount: sheet.view.geometry.pipe(map(g => g.rowCount)),
      columnCount: sheet.view.geometry.pipe(map(g => g.columnCount)),
      rowHeight: ROW_HEIGHT,
      rowHeights: heightsOf(hidden.value, sized.value),
      columnWidth: widths.value,
      frozenRows: frozen.value.rows,
      frozenColumns: frozen.value.columns,
      gutterWidth: GUTTER_WIDTH,
      headerHeight: HEADER_HEIGHT,
      rowOverscan: 2,
      columnOverscan: 1,
      scrollWith: gridNode$,
      role: 'grid',
      label: compare$.pipe(map(view => `${view.name}, beside the sheet`)),
      header: compareHeader,
      sheetRef: found => (compareWindow = found)
    },
    compareRow
  );
  const compareSheet = compareWindow as UiVirtualSheet | undefined;
  if (compareSheet !== undefined) {
    ctx.effect(widths, all => compareSheet.setColumnWidths(all));
    ctx.effect(hidden, rows => compareSheet.setRowHeights(heightsOf(rows, sized.value)));
    ctx.effect(sized, sizes => compareSheet.setRowHeights(heightsOf(hidden.value, sizes)));
    ctx.effect(frozen, pane => {
      compareCells.clear();
      compareValues.releaseAll();
      comparePaints.releaseAll();
      compareSheet.setFrozen(pane.rows, pane.columns);
    });
    // Cells that leave the pane's window let go of their streams, as the
    // grid's do; see the effect on `sheetWindow.range$`.
    ctx.effect(compareSheet.range$, range => {
      const pane = frozen.value;
      for (const [key, kept] of compareCells) {
        const rowShows = kept.row < pane.rows || (kept.row >= range.firstRow && kept.row <= range.lastRow);
        const columnShows =
          kept.column < pane.columns || (kept.column >= range.firstColumn && kept.column <= range.lastColumn);
        if (!rowShows || !columnShows) {
          compareCells.delete(key);
          compareValues.release(key);
          comparePaints.release(key);
        }
      }
    });
  }

  /**
   * A line above each pane naming what it shows, while there are two:
   * the same height on both sides, so the rows stay level.
   */
  const paneTitle = (text: Observable<string>, closing: boolean): UiElement =>
    Row(
      {
        width: percent(100),
        height: scaled(24),
        flexShrink: 0,
        y: 'center',
        gap: 8,
        paddingLeft: 8,
        paddingRight: 4,
        backgroundColor: 'surface',
        borderColor: GRID_LINE,
        borderWidth: 1
      },
      Text({ text, fontSize: 11, fontWeight: 'bold', color: 'text', selectable: false, textWrap: 'none' }),
      Box({ flex: 1, minWidth: 0 }),
      ...(closing
        ? [
            Button(
              {
                key: 'close',
                label: 'Close the side-by-side view',
                onClick: () => {
                  sheet.send.setCompare(false, null);
                  edit.focusSheet();
                },
                paddingLeft: 6,
                paddingRight: 6,
                borderRadius: 4,
                backgroundColor: 'surface',
                cursor: 'pointer'
              },
              Text({ text: '✕', fontSize: 11, color: 'textMuted', selectable: false })
            )
          ]
        : [])
    );

  const shownTitle = combineLatest([sheet.view.scenarios]).pipe(
    map(([view]) => view.entries.find(entry => entry.id === view.shown)?.name ?? 'Base')
  );
  const splitOpen = compare$.pipe(
    map(view => view.open),
    distinctUntilChanged()
  );
  const mainTitle = paneTitle(shownTitle, false);
  const besideColumn = Column(
    { key: 'beside', flex: 1, minWidth: 0, height: percent(100), borderColor: 'primary', borderWidth: 0 },
    paneTitle(compare$.pipe(map(view => view.name)), true),
    comparePane
  );

  return Row(
    { flex: 1, minHeight: 0, width: percent(100) },
    Column(
      { flex: 1, minWidth: 0, height: percent(100) },
      splitOpen.pipe(map(open => (open ? [mainTitle] : []))),
      Box({ flex: 1, minHeight: 0, width: percent(100) }, grid, menu)
    ),
    splitOpen.pipe(map(open => (open ? [Box({ key: 'divider', width: 2, height: percent(100), backgroundColor: 'border' }), besideColumn] : [])))
  );
}

/**
 * Where each column starts, from the widths: one more entry than there
 * are columns, the last being where the last column ends. Kept per array,
 * and the widths are replaced rather than changed, so it is summed once
 * per set of widths however many cells ask.
 */
const prefixes = new WeakMap<readonly number[], readonly number[]>();
function prefixOf(widths: readonly number[]): readonly number[] {
  let starts = prefixes.get(widths);
  if (starts === undefined) {
    const built = new Array<number>(widths.length + 1);
    let running = 0;
    for (let column = 0; column < widths.length; column++) {
      built[column] = running;
      running += widths[column];
    }
    built[widths.length] = running;
    starts = built;
    prefixes.set(widths, starts);
  }
  return starts;
}

/** Whether two lists hold the same elements in the same order. */
function sameElements(a: readonly UiElement[], b: readonly UiElement[]): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let at = 0; at < a.length; at++) {
    if (a[at] !== b[at]) {
      return false;
    }
  }
  return true;
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

/** How the grid draws each place `guessOf` can answer. */
const AUTO_ALIGN: Readonly<Record<Place, 'right' | 'center' | 'start'>> = { end: 'right', center: 'center', start: 'start' };

export type { SheetWindow };

/** Shared, because the overwhelming majority of cells have no border. */
/** The dot in a cell's corner when it breaks its validation rule. */
const MARK_SIZE = 6;
/** How big a note's corner mark is. */
const NOTE_MARK = 5;

const NO_SHAPES: DecorationShape[] = [];

/** The violet of the scenarios, faint: a histogram is read through, not at. */
const HISTOGRAM = 'rgba(124, 58, 237, 0.26)';

/**
 * A cell's histogram as rectangles: sixteen bars along the bottom of
 * the cell, the tallest reaching a little under half its height, behind
 * the figure. A bar of height zero is left out.
 */
function histogramOf(bars: readonly number[], width: number, height: number): DecorationShape[] {
  const inset = 3;
  const step = (width - inset * 2) / bars.length;
  const reach = Math.max(4, height * 0.45);
  const shapes: DecorationShape[] = [];
  bars.forEach((bar, index) => {
    if (bar <= 0) {
      return;
    }
    const tall = (bar / 8) * reach;
    shapes.push({
      kind: 'fill',
      x: inset + index * step,
      y: height - 1 - tall,
      width: Math.max(1, step - 1),
      height: tall,
      radius: 0,
      color: HISTOGRAM
    });
  });
  return shapes;
}

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

/** How wide an error's explanation may grow before it wraps. */
const EXPLAIN_WIDTH = 320;
/** How tall one name in a list hanging off the open cell is. */
const POPUP_ROW = 20;
/** How tall one line of a signature hint is. */
const POPUP_LINE = 16;
/** What a popup adds round its lines: two pixels of padding, top and bottom. A border is painted inside the box and takes no room. */
const POPUP_FRAME = 4;

/**
 * The panel a popup under the open cell is drawn on.
 *
 * **The height is written out, and it has to be.** A popup hangs off
 * the row that holds the open cell, starting at that row's bottom
 * edge, and the engine measures an absolute box against the room its
 * containing block has left below its `top` — which, for a box that
 * starts at the block's bottom edge, is none. Measured under a
 * maximum height of zero, the panel had no area to paint its
 * background, border or shadow in, and a `Box` stacks its children
 * rather than laying them out, so every name in the list and both
 * lines of a signature were drawn over one another in the same place,
 * bare on the gridlines. An explicit height is the box's own (the
 * engine's fix for charts in the same position), so the lines are
 * given fixed heights and the panel is their sum.
 */
function popupPanel(key: string, left: number, top: number, content: number) {
  return {
    key,
    position: 'absolute' as const,
    left,
    top: top + 1,
    height: content + POPUP_FRAME,
    zIndex: 4,
    backgroundColor: 'surface',
    borderColor: 'controlBorder',
    borderWidth: 1,
    borderRadius: 4,
    boxShadows: [boxShadow(0, 2, 8, 0, 'shadow')],
    paddingTop: 2,
    paddingBottom: 2,
    pointerEvents: 'none' as const
  };
}

/**
 * A listbox under the open cell: function names as a formula is
 * typed, or the values a validated cell allows.
 *
 * Not focusable and taking no pointer events — the keyboard stays in
 * the cell, and the chosen row is a band of the selection colour
 * that moves as the arrows do.
 */
function listPopup(key: string, label: string, left: number, top: number, shown: readonly string[], chosen: number): UiElement {
  const picked = Math.min(chosen, shown.length - 1);
  return Column(
    { ...popupPanel(key, left, top, shown.length * POPUP_ROW), role: 'listbox', label },
    ...shown.map((name, at) =>
      Text({
        key: name,
        text: name,
        role: 'option',
        label: name,
        states: at === picked ? ['selected'] : [],
        fontSize: 12,
        height: POPUP_ROW,
        verticalAlign: 'middle',
        paddingLeft: 8,
        paddingRight: 16,
        color: at === picked ? 'selectionForeground' : 'text',
        backgroundColor: at === picked ? 'selectionBackground' : undefined,
        textWrap: 'none'
      })
    )
  );
}


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
