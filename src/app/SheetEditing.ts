import { combineLatest, distinctUntilChanged, map, type Observable } from 'rxjs';

import { internalState, type ChannelReplica, type ComponentContext } from 'gesso-framework';

import type { CommandId } from './SheetCommands';
import { cornerOf, type SheetCommands, type SheetNames, type SheetPasteMode, type SheetSelection, type SheetView } from './SheetContract';
import { stampText, type SheetAction } from './SheetKeys';

export interface SheetEditing {
  /** Where the selection is, answered without a round trip. */
  readonly selection: Observable<SheetSelection>;
  readonly selectionNow: () => SheetSelection;
  /** The text of the cell being typed into, or null when none is open. */
  readonly draft: Observable<string | null>;
  readonly draftNow: () => string | null;
  readonly open: Observable<boolean>;
  readonly openNow: () => boolean;
  /** Whether a given cell is the one being typed into. */
  isOpen(row: number, column: number): Observable<boolean>;
  /** Applies what a key meant. Returns false when the key was not ours. */
  apply(action: SheetAction | null): boolean;
  /** Moves the far corner of the selection, keeping the anchor. */
  extendTo(row: number, column: number): void;
  /** Text arrived from the clipboard with the grid holding focus. */
  pasteText(text: string, mode?: SheetPasteMode): void;
  write(text: string): void;
  commit(rows: number, columns: number): void;
  cancel(): void;
  /** Opens the cell at the selection with what is already in it. */
  openCell(): void;
  /** Where the open edit was started, for deciding what takes focus. */
  readonly startedIn: () => 'grid' | 'bar' | null;
  /**
   * The format painter: off, lit for the next thing clicked, or lit
   * until Escape. Here because the toolbar lights it and the grid is
   * where the next click lands.
   */
  readonly painter: Observable<'off' | 'once' | 'held'> & { readonly value: 'off' | 'once' | 'held' };
  /**
   * View ▸ Show references: whether the selected cell's formula has the
   * cells it reads outlined, as they are while it is typed. On until it
   * is turned off, following the selection.
   */
  readonly referencesShown: Observable<boolean> & { readonly value: boolean };
  setReferencesShown(on: boolean): void;
  setPainter(state: 'off' | 'once' | 'held'): void;
  moveTo(row: number, column: number): void;
  /**
   * A rectangle as the selection, with the active cell given: what a
   * click on a column's letter or a row's number means. The active
   * cell is one corner and the anchor the other.
   */
  selectRect(row: number, column: number, anchorRow: number, anchorColumn: number, active?: { row: number; column: number }): void;
  /**
   * Puts the keyboard back on the sheet.
   *
   * The chrome needs this and cannot do it: the grid's node belongs
   * to the grid, and the menu bar, the find bar and the name box all
   * have to hand the keyboard back when they are done with it — a
   * menu that closed and left focus on itself is a menu you have to
   * Tab out of before you can type a number.
   *
   * It lives on the editing handle because that is already the one
   * thing the grid and the screen around it share, and a second
   * shared object would be a second thing to keep in step.
   */
  focusSheet(): void;
  /** The named ranges, for the name box to resolve and to answer with. */
  readonly names: Observable<SheetNames>;
  /** Gives the selection a name. The answer arrives on `names`. */
  defineName(name: string): void;
  /** Gives a name a formula to hold. The answer arrives on `names`. */
  defineFormulaName(name: string, formula: string): void;
  /** The grid, saying which node that is. Called once, on mount. */
  provideFocus(run: () => void): void;
  /**
   * Runs an application command.
   *
   * The grid answers the accelerators — it is what holds focus while
   * somebody is using the sheet, so it is where the key arrives — and
   * the chrome is what knows how to run them. Routed through here for
   * the reason `focusSheet` is: this handle is already the one thing
   * the two share.
   */
  runCommand(id: CommandId): void;
  /** The chrome, saying how. Called once, on mount. */
  provideCommands(run: (id: CommandId) => void): void;
  /**
   * Closes whatever the chrome has open. True when something closed.
   *
   * Escape has to be routed rather than left to bubble, because a
   * dialog is in the overlay layer and the grid is not its child, so
   * a key pressed at the grid never passes through it.
   *
   * It would not need routing if the dialog held the keyboard, and
   * `Dialog` in `gesso-components` means to: it calls `focus.trap` on
   * its body as that body mounts. But it never calls `focus.focus`,
   * and its body is not `focusable` — so a dialog opened while focus
   * was on something else traps a keyboard it does not have, and its
   * own `Escape` handler is bound to a node no key arrives at. Found
   * in a browser, by pressing Escape; every spec passed without it,
   * because a spec that opens a dialog and asserts it is open never
   * asks what has the keyboard.
   */
  dismiss(): boolean;
  /** The chrome, saying what it has open. Called once, on mount. */
  provideDismiss(run: () => boolean): void;
  /**
   * Puts the tab strip into rename mode, for `Sheet ▸ Rename`.
   *
   * The command is dispatched by the top bar and the box that has to
   * open is along the bottom, and the two are siblings — so it goes
   * the way every other cross-chrome call already goes, through the
   * handle both of them hold. The alternative is a rename that works
   * from the strip and not from the menu, which is the half of the
   * feature a keyboard cannot reach.
   */
  renameSheet(): void;
  provideRename(run: () => void): void;
  /** Puts the keyboard on the tab strip, for Alt+F10. */
  focusTabs(): void;
  provideTabs(run: () => void): void;
  /** Sheet ▸ Tab colour: the palette, beside the active tab. */
  tabColour(): void;
  provideTabColour(run: () => void): void;
}

/**
 * Selection and the open cell, on the render thread.
 *
 * **Why the selection is held here and not simply read off the
 * channel.** It is on the channel too — the application worker needs
 * it to say what the formula bar should show, and Phase 5 needs it to
 * know what a copy copies. But a selection read back across a barrier
 * lags by a frame or two, and every arrow key pressed inside that
 * window would move from the same stale cell: hold an arrow down and
 * the selection travels one cell and stops. So this leads and the
 * channel follows, and when the channel moves the selection for a
 * reason of its own — an undo jumping to the cell it put back — this
 * adopts it.
 *
 * The draft is only here. A cell being typed into holds text the
 * application thread has not seen and must not see until it is
 * committed, which is what makes Escape possible at all.
 */
export function editing(
  ctx: ComponentContext,
  sheet: ChannelReplica<SheetView, SheetCommands>
): SheetEditing {
  const selection = internalState<SheetSelection>({ row: 0, column: 0, anchorRow: 0, anchorColumn: 0 });
  const draft = internalState<string | null>(null);
  // How far the sheet goes is the application worker's to say, and it
  // says so on the `geometry` key. Reading it from a constant on this
  // side would be the same number written twice, and `End` would walk
  // to wherever the render thread happened to believe the edge was.
  const extent = () => sheet.view.geometry.value;
  /**
   * Which of the two views the open edit was started from.
   *
   * Only focus cares. An edit begun in the grid puts the caret in the
   * cell; one begun by clicking the formula bar must leave the caret
   * where the person put it, or the click would bounce them into the
   * cell they were trying to avoid.
   */
  let startedIn: 'grid' | 'bar' | null = null;
  /** Set by the grid on mount; does nothing until then. */
  let focusSheet: () => void = () => {};
  /** Set by the chrome on mount; does nothing until then. */
  let runCommand: (id: CommandId) => void = () => {};
  let dismiss: () => boolean = () => false;
  let renameSheet: () => void = () => {};
  let focusTabs: () => void = () => {};
  let tabColour: () => void = () => {};

  // The application worker's selection, when it is not one we caused.
  // Sending `setSelection` echoes the value straight back, which
  // `same` filters out, so this is not a loop.
  ctx.effect(sheet.view.selection, next => {
    if (!same(next, selection.value)) {
      selection.value = next;
    }
  });

  /** A selection, set here and sent over. */
  const put = (next: SheetSelection): void => {
    if (same(next, selection.value)) {
      return;
    }
    selection.value = next;
    const corner = cornerOf(next);
    if (corner.row === next.row && corner.column === next.column) {
      sheet.send.setSelection(next.row, next.column, next.anchorRow, next.anchorColumn);
    } else {
      sheet.send.setSelection(next.row, next.column, next.anchorRow, next.anchorColumn, corner.row, corner.column);
    }
  };

  /**
   * A cell as the selection: collapsed to it, or — extending — the far
   * corner moved to it with the active cell left where it was, which is
   * Shift's rule in Excel. The name box and the formula bar go on
   * showing the cell the selection started from.
   */
  const place = (row: number, column: number, keepAnchor: boolean): void => {
    const { rowCount, columnCount } = extent();
    const clampedRow = clamp(row, 0, Math.max(0, rowCount - 1));
    const clampedColumn = clamp(column, 0, Math.max(0, columnCount - 1));
    const held = selection.value;
    if (!keepAnchor) {
      put({ row: clampedRow, column: clampedColumn, anchorRow: clampedRow, anchorColumn: clampedColumn });
      return;
    }
    put({
      row: held.row,
      column: held.column,
      anchorRow: held.anchorRow,
      anchorColumn: held.anchorColumn,
      cornerRow: clampedRow,
      cornerColumn: clampedColumn
    });
  };

  const painter = internalState<'off' | 'once' | 'held'>('off');
  const referencesShown = internalState(false);
  /**
   * A rectangle from `row, column` to the anchor. The active cell is
   * the first corner unless it is given — a Shift+click on a header
   * keeps the one the selection had.
   */
  const selectRect = (
    row: number,
    column: number,
    anchorRow: number,
    anchorColumn: number,
    active?: { row: number; column: number }
  ): void => {
    tabFrom = null;
    const { rowCount, columnCount } = extent();
    const lastRow = Math.max(0, rowCount - 1);
    const lastColumn = Math.max(0, columnCount - 1);
    const corner = { row: clamp(row, 0, lastRow), column: clamp(column, 0, lastColumn) };
    const at = active === undefined ? corner : { row: clamp(active.row, 0, lastRow), column: clamp(active.column, 0, lastColumn) };
    put({
      row: at.row,
      column: at.column,
      anchorRow: clamp(anchorRow, 0, lastRow),
      anchorColumn: clamp(anchorColumn, 0, lastColumn),
      ...(at.row === corner.row && at.column === corner.column ? {} : { cornerRow: corner.row, cornerColumn: corner.column })
    });
  };

  /**
   * The column a run of Tabs started in, or null outside one.
   *
   * Typing a table row is Tab, Tab, Tab, Enter — and the Enter goes to
   * the start of the *next* row, not to the cell below the last Tab.
   * Anything but a Tab or that Enter ends the run: an arrow, a click,
   * a jump.
   */
  let tabFrom: number | null = null;

  const moveTo = (row: number, column: number): void => {
    tabFrom = null;
    place(row, column, false);
  };
  const extendTo = (row: number, column: number): void => {
    tabFrom = null;
    place(row, column, true);
  };

  /**
   * Where Enter (or Tab) goes from the cursor, and what that leaves the
   * run of Tabs as. Down one row after a run of Tabs goes back to the
   * column the run began in.
   */
  const step = (rows: number, columns: number, tab: boolean, walks = true): void => {
    const at = selection.value;
    if (walks && walk(at, rows, columns)) {
      tabFrom = null;
      return;
    }
    const back = !tab && rows === 1 && columns === 0 && tabFrom !== null ? tabFrom : null;
    tabFrom = tab ? (tabFrom ?? at.column) : null;
    place(at.row + rows, back ?? at.column + columns, false);
  };

  /**
   * Enter and Tab inside a selection of more than one cell move the
   * active cell through it and leave it selected: Enter down the
   * column and on to the top of the next, Tab along the row and on to
   * the start of the next, Shift going back, both wrapping at the end —
   * how a block of figures is typed into a range selected first.
   * Returns false for a selection of one cell, which moves as it always
   * did.
   */
  const walk = (at: SheetSelection, rows: number, columns: number): boolean => {
    const corner = cornerOf(at);
    const firstRow = Math.min(corner.row, at.anchorRow);
    const lastRow = Math.max(corner.row, at.anchorRow);
    const firstColumn = Math.min(corner.column, at.anchorColumn);
    const lastColumn = Math.max(corner.column, at.anchorColumn);
    if (firstRow === lastRow && firstColumn === lastColumn) {
      return false;
    }
    if (Math.abs(rows) + Math.abs(columns) !== 1) {
      return false;
    }
    const height = lastRow - firstRow + 1;
    const width = lastColumn - firstColumn + 1;
    // The cells in the order the key walks them: down columns for
    // Enter, along rows for Tab.
    const down = rows !== 0;
    const index = down
      ? (at.column - firstColumn) * height + (at.row - firstRow)
      : (at.row - firstRow) * width + (at.column - firstColumn);
    const count = height * width;
    const next = (((index + (rows + columns)) % count) + count) % count;
    const row = down ? firstRow + (next % height) : firstRow + Math.floor(next / width);
    const column = down ? firstColumn + Math.floor(next / height) : firstColumn + (next % width);
    put({
      row,
      column,
      anchorRow: at.anchorRow,
      anchorColumn: at.anchorColumn,
      cornerRow: corner.row,
      cornerColumn: corner.column
    });
    return true;
  };

  const commit = (rows: number, columns: number, tab = false): void => {
    const text = draft.value;
    const at = selection.value;
    startedIn = null;
    draft.value = null;
    if (text !== null) {
      sheet.send.setCell(at.row, at.column, text);
    }
    if (rows !== 0 || columns !== 0) {
      step(rows, columns, tab);
    }
  };

  /**
   * Opens the cell with what was typed into it.
   *
   * The text comes from the `editor` view key, which the application
   * worker republishes whenever the selection moves. It can lag the
   * local selection by a frame, so it is used only when it is about
   * the cell being opened; opening a cell the worker has not caught up
   * with yet starts from empty rather than from a neighbour's formula,
   * which is the safe direction of the two.
   */
  const openCell = (where: 'grid' | 'bar' = 'grid'): void => {
    const at = selection.value;
    const current = sheet.view.editor.value;
    startedIn = where;
    draft.value = current.row === at.row && current.column === at.column ? current.input : '';
  };

  const apply = (action: SheetAction | null): boolean => {
    if (action === null) {
      return false;
    }
    const at = selection.value;
    // Anything else that moves the cursor ends a run of Tabs. Typing
    // into the cell between them is the point of the run, so it does not.
    if (action.kind === 'edge' || action.kind === 'jump' || action.kind === 'selectLine' || action.kind === 'selectAll') {
      tabFrom = null;
    }
    switch (action.kind) {
      case 'move':
        if (action.extend) {
          tabFrom = null;
          // From the corner, which is the end Shift moves.
          place(cornerOf(at).row + action.rows, cornerOf(at).column + action.columns, true);
        } else {
          step(action.rows, action.columns, action.tab === true, action.walks === true);
        }
        return true;
      case 'edge':
        sheet.send.jumpToEdge(
          action.extend ? cornerOf(at).row : at.row,
          action.extend ? cornerOf(at).column : at.column,
          at.anchorRow,
          at.anchorColumn,
          action.rows,
          action.columns,
          action.extend
        );
        return true;
      case 'selectLine': {
        const { rowCount, columnCount } = extent();
        if (action.axis === 'columns') {
          selectRect(0, at.column, rowCount - 1, at.column);
        } else {
          selectRect(at.row, 0, at.row, columnCount - 1);
        }
        return true;
      }
      case 'stamp':
        // With a cell open the stamp goes in at the caret, which is the
        // grid's to place; see `Grid`'s key handler.
        if (draft.value !== null) {
          return false;
        }
        startedIn = 'grid';
        draft.value = stampText(action.what);
        return true;
      case 'insert':
        // The caret is the grid's, as above.
        return false;
      case 'commitAll': {
        const text = draft.value;
        startedIn = null;
        draft.value = null;
        if (text !== null) {
          sheet.send.writeSelection(text);
        }
        return true;
      }
      case 'jump':
        switch (action.to) {
          case 'rowStart':
            place(action.extend ? cornerOf(at).row : at.row, 0, action.extend);
            return true;
          case 'rowEnd':
            place(action.extend ? cornerOf(at).row : at.row, extent().columnCount - 1, action.extend);
            return true;
          case 'sheetStart':
            place(0, 0, action.extend);
            return true;
          default:
            place(extent().rowCount - 1, extent().columnCount - 1, action.extend);
            return true;
        }
      case 'edit':
        openCell();
        return true;
      case 'replace':
        startedIn = 'grid';
        draft.value = action.text;
        return true;
      case 'commit':
        commit(action.rows, action.columns, action.tab === true);
        return true;
      case 'cancel':
        startedIn = null;
        draft.value = null;
        return true;
      case 'clear':
        // The whole selection, which for one cell is that cell.
        sheet.send.clearRange();
        return true;
      case 'copy':
        sheet.send.copy(action.cut);
        return true;
      case 'selectAll': {
        const { rowCount, columnCount } = extent();
        selectRect(0, 0, rowCount - 1, columnCount - 1);
        return true;
      }
      case 'undo':
        sheet.send.undo();
        return true;
      case 'redo':
        sheet.send.redo();
        return true;
    }
  };

  return {
    selection,
    selectionNow: () => selection.value,
    names: sheet.view.names,
    defineName: (name: string) => sheet.send.defineName(name),
    defineFormulaName: (name: string, formula: string) => sheet.send.defineFormulaName(name, formula),
    draft,
    draftNow: () => draft.value,
    open: draft.pipe(
      map(text => text !== null),
      distinctUntilChanged()
    ),
    openNow: () => draft.value !== null,
    isOpen: (row, column) =>
      combineLatest([selection, draft]).pipe(
        map(([at, text]) => text !== null && at.row === row && at.column === column),
        distinctUntilChanged()
      ),
    apply,
    /**
     * Text reported by one of the two fields.
     *
     * This is also how an edit begins in the formula bar: nothing
     * happens when the caret arrives there, and the first character
     * typed or deleted opens the cell. An edit that opens this way is
     * a bar edit, so the cell that appears underneath does not steal
     * the caret back.
     */
    write: text => {
      startedIn ??= 'bar';
      draft.value = text;
    },
    commit,
    cancel: () => (draft.value = null),
    openCell: () => openCell('grid'),
    startedIn: () => startedIn,
    extendTo,
    pasteText: (text, mode) => sheet.send.paste(text, mode),
    moveTo,
    selectRect,
    painter,
    setPainter: state => {
      painter.value = state;
    },
    referencesShown,
    setReferencesShown: on => {
      referencesShown.value = on;
    },
    focusSheet: () => focusSheet(),
    provideFocus: run => {
      focusSheet = run;
    },
    runCommand: id => runCommand(id),
    provideCommands: run => {
      runCommand = run;
    },
    dismiss: () => dismiss(),
    provideDismiss: run => {
      dismiss = run;
    },
    renameSheet: () => renameSheet(),
    provideRename: run => {
      renameSheet = run;
    },
    tabColour: () => tabColour(),
    provideTabColour: run => {
      tabColour = run;
    },
    focusTabs: () => focusTabs(),
    provideTabs: run => {
      focusTabs = run;
    }
  };
}

function same(a: SheetSelection, b: SheetSelection): boolean {
  const cornerA = cornerOf(a);
  const cornerB = cornerOf(b);
  return (
    a.row === b.row &&
    a.column === b.column &&
    a.anchorRow === b.anchorRow &&
    a.anchorColumn === b.anchorColumn &&
    cornerA.row === cornerB.row &&
    cornerA.column === cornerB.column
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}
