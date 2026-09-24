import { combineLatest, distinctUntilChanged, map, type Observable } from 'rxjs';

import { internalState, type ChannelReplica, type ComponentContext } from 'gesso-framework';

import type { SheetCommands, SheetSelection, SheetView } from './SheetContract';
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
  write(text: string): void;
  commit(rows: number, columns: number): void;
  cancel(): void;
  /** Opens the cell at the selection with what is already in it. */
  openCell(): void;
  /** Where the open edit was started, for deciding what takes focus. */
  readonly startedIn: () => 'grid' | 'bar' | null;
  moveTo(row: number, column: number): void;
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

  // The application worker's selection, when it is not one we caused.
  // Sending `setSelection` echoes the value straight back, which
  // `same` filters out, so this is not a loop.
  ctx.effect(sheet.view.selection, next => {
    if (!same(next, selection.value)) {
      selection.value = next;
    }
  });

  const moveTo = (row: number, column: number): void => {
    const { rowCount, columnCount } = extent();
    const clampedRow = clamp(row, 0, Math.max(0, rowCount - 1));
    const clampedColumn = clamp(column, 0, Math.max(0, columnCount - 1));
    const next = { row: clampedRow, column: clampedColumn, anchorRow: clampedRow, anchorColumn: clampedColumn };
    if (same(next, selection.value)) {
      return;
    }
    selection.value = next;
    sheet.send.setSelection(clampedRow, clampedColumn, clampedRow, clampedColumn);
  };

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
        moveTo(at.row + action.rows, at.column + action.columns);
        return true;
      case 'jump':
        switch (action.to) {
          case 'rowStart':
            moveTo(at.row, 0);
            return true;
          case 'rowEnd':
            moveTo(at.row, extent().columnCount - 1);
            return true;
          case 'sheetStart':
            moveTo(0, 0);
            return true;
          default:
            moveTo(extent().rowCount - 1, extent().columnCount - 1);
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
        sheet.send.setCell(at.row, at.column, '');
        return true;
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
    moveTo
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
