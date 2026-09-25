import { combineLatest, distinctUntilChanged, map, type Observable } from 'rxjs';

import { internalState, type ChannelReplica, type ComponentContext } from 'gesso-framework';

import type { CommandId } from './SheetCommands';
import type { SheetCommands, SheetNames, SheetSelection, SheetView } from './SheetContract';
import type { SheetAction } from './SheetKeys';

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
  pasteText(text: string): void;
  write(text: string): void;
  commit(rows: number, columns: number): void;
  cancel(): void;
  /** Opens the cell at the selection with what is already in it. */
  openCell(): void;
  /** Where the open edit was started, for deciding what takes focus. */
  readonly startedIn: () => 'grid' | 'bar' | null;
  moveTo(row: number, column: number): void;
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

  // The application worker's selection, when it is not one we caused.
  // Sending `setSelection` echoes the value straight back, which
  // `same` filters out, so this is not a loop.
  ctx.effect(sheet.view.selection, next => {
    if (!same(next, selection.value)) {
      selection.value = next;
    }
  });

  const place = (row: number, column: number, keepAnchor: boolean): void => {
    const { rowCount, columnCount } = extent();
    const clampedRow = clamp(row, 0, Math.max(0, rowCount - 1));
    const clampedColumn = clamp(column, 0, Math.max(0, columnCount - 1));
    const held = selection.value;
    const next = {
      row: clampedRow,
      column: clampedColumn,
      anchorRow: keepAnchor ? held.anchorRow : clampedRow,
      anchorColumn: keepAnchor ? held.anchorColumn : clampedColumn
    };
    if (same(next, held)) {
      return;
    }
    selection.value = next;
    sheet.send.setSelection(next.row, next.column, next.anchorRow, next.anchorColumn);
  };

  const moveTo = (row: number, column: number): void => place(row, column, false);
  const extendTo = (row: number, column: number): void => place(row, column, true);

  const commit = (rows: number, columns: number): void => {
    const text = draft.value;
    const at = selection.value;
    startedIn = null;
    draft.value = null;
    if (text !== null) {
      sheet.send.setCell(at.row, at.column, text);
    }
    if (rows !== 0 || columns !== 0) {
      moveTo(at.row + rows, at.column + columns);
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
    switch (action.kind) {
      case 'move':
        place(at.row + action.rows, at.column + action.columns, action.extend);
        return true;
      case 'jump':
        switch (action.to) {
          case 'rowStart':
            place(at.row, 0, action.extend);
            return true;
          case 'rowEnd':
            place(at.row, extent().columnCount - 1, action.extend);
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
        commit(action.rows, action.columns);
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
        selection.value = { row: 0, column: 0, anchorRow: rowCount - 1, anchorColumn: columnCount - 1 };
        sheet.send.setSelection(0, 0, rowCount - 1, columnCount - 1);
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
    pasteText: text => sheet.send.paste(text),
    moveTo,
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
    }
  };
}

function same(a: SheetSelection, b: SheetSelection): boolean {
  return (
    a.row === b.row && a.column === b.column && a.anchorRow === b.anchorRow && a.anchorColumn === b.anchorColumn
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}
