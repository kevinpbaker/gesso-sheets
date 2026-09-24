import { combineLatest, type Observable } from 'rxjs';
import { distinctUntilChanged, map } from 'rxjs/operators';

import {
  Box,
  EditableText,
  editorFor,
  LazySheet,
  percent,
  Row,
  Text,
  type UiElement,
  type UiKeyboardEvent,
  type UiNode,
  type UiTextChangeEvent,
  type UiVirtualSheet
} from 'gesso-core';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

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
      onClick: () => selectByPointer(row, column)
    });
  };

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
  const selectByPointer = (row: number, column: number): void => {
    // A click elsewhere commits what is open, as it does everywhere.
    if (edit.openNow()) {
      edit.commit(0, 0);
    }
    edit.moveTo(row, column);
    if (gridNode !== null) {
      focus.focus(gridNode);
    }
  };

  const cell = (row: number, column: number): UiElement => {
    if (openAt !== null && openAt.row === row && openAt.column === column) {
      return editorCell(row, column);
    }
    const key = `${row}:${column}`;
    let built = cells.get(key);
    if (built === undefined) {
      built = buildCell(row, column);
      cells.set(key, built);
    }
    return built;
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
  const editorCell = (row: number, column: number): UiElement =>
    EditableText({
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
        }
      },
      value: edit.draft.pipe(map(text => text ?? '')),
      width: widths.pipe(map(all => all[column] ?? COLUMN_WIDTH)),
      height: ROW_HEIGHT,
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
    const action = keyAction(event.key, event.modifiers, edit.openNow());
    if (edit.apply(action)) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

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

  /** Written only when the selection needs bringing into view. */
  const scrollX = internalState(0);
  const scrollY = internalState(0);
  let gridNode: UiNode | null = null;
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
      ref: node => (gridNode = node),
      modifiers: [viewport.modifier],
      onKeyDown: onKey,
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
  let openBefore: { row: number; column: number } | null = null;
  ctx.effect(combineLatest([edit.selection, edit.open]), ([at, isOpen]) => {
    const next = isOpen ? { row: at.row, column: at.column } : null;
    if (sameCell(next, openBefore)) {
      return;
    }
    for (const which of [openBefore, next]) {
      if (which !== null) {
        cells.delete(`${which.row}:${which.column}`);
      }
    }
    openBefore = next;
    openAt = next;
    sheetWindow.invalidate();
  });

  // Focus follows the edit: into the cell when one opens, back to the
  // grid when it closes, or the next keystroke goes nowhere.
  ctx.effect(edit.open, isOpen => {
    const target = isOpen ? editorNode : gridNode;
    if (target !== null) {
      focus.focus(target);
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
