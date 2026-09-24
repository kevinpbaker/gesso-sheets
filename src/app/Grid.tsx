import { combineLatest, type Observable } from 'rxjs';
import { distinctUntilChanged, map } from 'rxjs/operators';

import { Box, LazySheet, percent, Row, Text, type UiElement, type UiVirtualSheet } from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName } from '../sheet/A1';

import {
  COLUMN_COUNT,
  COLUMN_WIDTH,
  GUTTER_WIDTH,
  HEADER_HEIGHT,
  MIN_COLUMN_WIDTH,
  ROW_COUNT,
  ROW_HEIGHT
} from './dimensions';
import { cellIn, Sheet, type SheetSelection, type SheetWindow } from './SheetContract';

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

const SELECTED_WASH = 'selectionBackground';
const GRID_LINE = 'border';

export function Grid(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const window$ = sheet.view.window;
  const selection$ = sheet.view.selection;

  /**
   * The cells now mounted, by the cell they hold.
   *
   * An element is immutable, so a cell still in the window can simply
   * be handed back. Without this a column step rebuilds every row and
   * every row rebuilds its cells, and building a cell allocates its
   * bindings — about eleven hundred of them, once or twice a frame.
   */
  const cells = new Map<string, UiElement>();

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

  /** Where this cell stands in the selection, as one cheap number. */
  const standing = (row: number, column: number): Observable<0 | 1 | 2> =>
    selection$.pipe(
      map(selection => {
        if (selection.row === row && selection.column === column) {
          return 2;
        }
        return inRange(selection, row, column) ? 1 : 0;
      }),
      distinctUntilChanged()
    );

  const buildCell = (row: number, column: number): UiElement => {
    const value = window$.pipe(
      map(current => cellIn(current, row, column)),
      distinctUntilChanged()
    );
    const state = standing(row, column);
    return Text({
      key: column,
      text: value.pipe(map(text => text ?? '')),
      // A cell the application worker has not sent yet is drawn as a
      // rule rather than left blank, so a gap on a fling reads as
      // "not here yet" instead of as the end of the sheet.
      color: value.pipe(map(text => (text === null ? 'placeholder' : 'text'))),
      backgroundColor: state.pipe(map(where => (where === 0 ? 'background' : SELECTED_WASH))),
      borderColor: state.pipe(map(where => (where === 2 ? 'primary' : GRID_LINE))),
      borderWidth: state.pipe(map(where => (where === 2 ? 2 : 1))),
      width: widths.pipe(map(all => all[column] ?? COLUMN_WIDTH)),
      height: ROW_HEIGHT,
      flexShrink: 0,
      paddingLeft: 6,
      paddingRight: 6,
      fontSize: 12,
      textWrap: 'none',
      textOverflow: 'clip',
      verticalAlign: 'middle',
      textAlign: value.pipe(map(text => (isNumeric(text) ? 'right' : 'start'))),
      role: 'cell',
      onClick: () => sheet.send.setSelection(row, column, row, column)
    });
  };

  const cell = (row: number, column: number): UiElement => {
    const key = `${row}:${column}`;
    let built = cells.get(key);
    if (built === undefined) {
      built = buildCell(row, column);
      cells.set(key, built);
    }
    return built;
  };

  /** The frozen strip at the start of a row: the row's number. */
  const rowHeader = (row: number): UiElement =>
    Text({
      key: 'gutter',
      text: String(row + 1),
      width: GUTTER_WIDTH,
      height: ROW_HEIGHT,
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
      role: 'rowheader'
    });

  const renderRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [rowHeader(row)];
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(cell(row, column));
    }
    return Row({ role: 'row', posInSet: row + 1 }, ...line);
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
  const columnHeader = (column: number): UiElement =>
    Box(
      {
        key: column,
        width: widths.pipe(map(all => all[column] ?? COLUMN_WIDTH)),
        height: HEADER_HEIGHT,
        flexShrink: 0,
        x: 'center',
        y: 'center',
        // The grip is absolute, and an absolute child is placed
        // against its nearest *positioned* ancestor — without this it
        // would walk past the header to the layout root and be drawn
        // in the corner of the screen.
        position: 'relative',
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
        verticalAlign: 'middle'
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
          resizing = null;
        }
      })
    );

  let resizing: { column: number; width: number; from: number } | null = null;

  /** A column may be dragged narrow, but not to nothing. */
  const resizeTo = (column: number, width: number): void => {
    const next = widths.value.slice();
    next[column] = Math.max(MIN_COLUMN_WIDTH, Math.round(width));
    widths.value = next;
  };

  let window: UiVirtualSheet | undefined;
  const grid = LazySheet(
    {
      flex: 1,
      minHeight: 0,
      width: percent(100),
      backgroundColor: 'background',
      rowCount: ROW_COUNT,
      columnCount: COLUMN_COUNT,
      rowHeight: ROW_HEIGHT,
      columnWidth: widths.value,
      gutterWidth: GUTTER_WIDTH,
      headerHeight: HEADER_HEIGHT,
      // Two rows and one column: the partial cells at the edges. The
      // lookahead is the application worker's; see PHASE0.md §3.
      rowOverscan: 2,
      columnOverscan: 1,
      role: 'grid',
      label: 'Sheet',
      focusable: true,
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

  // The round trip: the range the window settled on is what the
  // application worker is asked for. `range$` emits only when the
  // range changes, so this is not a command per frame.
  ctx.effect(sheetWindow.range$, range =>
    sheet.send.setViewport(range.firstRow, range.lastRow, range.firstColumn, range.lastColumn)
  );

  // Cells that scrolled away, so their bindings go with them.
  ctx.effect(sheetWindow.range$, range => {
    for (const key of cells.keys()) {
      const colon = key.indexOf(':');
      const row = Number(key.slice(0, colon));
      const column = Number(key.slice(colon + 1));
      if (row < range.firstRow || row > range.lastRow || column < range.firstColumn || column > range.lastColumn) {
        cells.delete(key);
      }
    }
  });

  return grid;
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
