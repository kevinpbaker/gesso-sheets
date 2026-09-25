import { BehaviorSubject, combineLatest, type Observable } from 'rxjs';
import { map } from 'rxjs/operators';

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
  type UiPasteEvent,
  type UiPointerEvent,
  type UiTextChangeEvent,
  type UiVirtualSheet
} from 'gesso-core';
import { FocusService, internalState, ShellService, type ComponentContext, type Inputs } from 'gesso-framework';

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
  readonly value: BehaviorSubject<string | null>;
  readonly standing: BehaviorSubject<Standing>;
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
   * worker — and drew at 22fps.
   *
   * So the subjects a cell subscribes to are its own, with two or three
   * observers each, and one subscriber per source fills them. Removing
   * a cell now scans a list of three. The same drag draws at over a
   * hundred.
   */
  let latestWindow: SheetWindow | null = null;
  let latestSelection: SheetSelection | null = null;

  ctx.effect(window$, current => {
    latestWindow = current;
    for (const mounted of cells.values()) {
      const next = cellIn(current, mounted.row, mounted.column);
      if (next !== mounted.value.value) {
        mounted.value.next(next);
      }
    }
  });

  ctx.effect(selection$, selection => {
    latestSelection = selection;
    for (const mounted of cells.values()) {
      const next = standingOf(selection, mounted.row, mounted.column);
      if (next !== mounted.standing.value) {
        mounted.standing.next(next);
      }
    }
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
    for (const [column, width] of columnWidths) {
      const next = all[column] ?? COLUMN_WIDTH;
      if (next !== width.value) {
        width.next(next);
      }
    }
  });

  const buildCell = (row: number, column: number, value: Observable<string | null>, state: Observable<Standing>): UiElement => {
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
      width: widthOf(column),
      height: ROW_HEIGHT,
      flexShrink: 0,
      paddingLeft: 6,
      paddingRight: 6,
      fontSize: 12,
      textWrap: 'none',
      textOverflow: 'clip',
      verticalAlign: 'middle',
      textAlign: value.pipe(map(text => (isNumeric(text) ? 'right' : 'start'))),
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
    // A click elsewhere commits what is open, as it does everywhere.
    if (edit.openNow()) {
      edit.commit(0, 0);
    }
    if (extend) {
      edit.extendTo(row, column);
    } else {
      edit.moveTo(row, column);
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
    const value = new BehaviorSubject<string | null>(
      latestWindow === null ? null : cellIn(latestWindow, row, column)
    );
    const standing = new BehaviorSubject<Standing>(
      latestSelection === null ? 0 : standingOf(latestSelection, row, column)
    );
    const element = buildCell(row, column, value, standing);
    cells.set(key, { row, column, element, value, standing });
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
      width: widthOf(column),
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
    return sheetWindow.cellAt(event.x - box.x, event.y - box.y);
  };

  const renderRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [rowHeader(row)];
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(cell(row, column));
    }
    const corner = cornerAt !== null && cornerAt.row === row;
    if (corner && cornerAt !== null) {
      line.push(fillHandle(cornerAt.column));
    }
    return Row({ role: 'row', posInSet: row + 1, position: corner ? 'relative' : undefined }, ...line);
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
        width: widthOf(column),
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
