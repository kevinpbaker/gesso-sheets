import { channel } from 'gesso-framework';

/**
 * The barrier.
 *
 * Imported by the render worker and by the application worker, and
 * holding nothing but names and shapes — the store, the parser, the
 * dependency graph and the recalc are behind it and the render worker
 * never loads a line of them.
 *
 * Five view keys rather than one object, for the reason
 * `NotesContract.ts` splits `rows` from `open`: the differ walks a
 * projection structurally on every publish, so two things that change
 * at different rates must not share a key. A keystroke moves `window`
 * and `editor`; it does not touch `geometry`, and the differ should
 * not have to walk a hundred column widths to find that out.
 */

/**
 * A window of cells, as display strings, keyed by absolute row and
 * then absolute column.
 *
 * **Not arrays**, and this is the single most consequential line in
 * the file. Phase 0 measured both: a row-major array costs 1,199
 * patches and 90 KiB per publish against 6.2 patches and 1.1 KiB for
 * this. `diffArray` trims a common prefix and suffix, and a window
 * that scrolled by one row has neither; the two windows are also the
 * same length, so it does not even splice — it recurses and emits one
 * patch per cell. Keyed by where a cell *is*, a scroll leaves every
 * other key structurally equal and the patch is the row that entered
 * and the row that left. See PHASE0.md section 4.
 *
 * A key that is absent is a cell that has not arrived. That is what
 * makes coverage countable, which is what made the round trip
 * measurable in Phase 0 and is why the shape is worth keeping.
 */
export interface SheetWindow {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  readonly cells: Readonly<Record<string, Readonly<Record<string, string>>>>;
}

/** The sheet's shape. Changes when a column is resized, and not otherwise. */
export interface SheetGeometry {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowHeight: number;
  readonly columnWidth: number;
}

/** The active cell, and the rectangle anchored from it. */
export interface SheetSelection {
  readonly row: number;
  readonly column: number;
  readonly anchorRow: number;
  readonly anchorColumn: number;
}

/**
 * What the formula bar shows: the active cell's text as it was typed,
 * which is not what the window holds for it.
 *
 * `A1` displays `3` and was typed `=1+2`, and the two live on separate
 * keys because they change at different moments — moving the selection
 * moves this and leaves the window alone.
 */
export interface SheetEditor {
  readonly row: number;
  readonly column: number;
  readonly input: string;
}

/**
 * Text the sheet wants put on the system clipboard.
 *
 * A command returns nothing and an effect comes back as a patch, so a
 * copy is a request one way and an answer the other: the render worker
 * asks, the application worker builds the block and publishes it here,
 * and the render worker hands it to the shell — which is the only
 * thread with a clipboard.
 *
 * `serial` is what makes copying the same cells twice a change. Without
 * it the second copy is structurally equal to the first, the differ
 * says nothing happened, and the clipboard is never written.
 */
export interface SheetClipboard {
  readonly text: string;
  readonly serial: number;
}

export interface SheetStatus {
  /** Cells whose value is still out of date. Zero when settled. */
  readonly pending: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

export interface SheetCommands {
  /** The range the render worker has mounted. Sent when it changes. */
  setViewport(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void;
  /** Commits what was typed into a cell. */
  setCell(row: number, column: number, input: string): void;
  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void;
  undo(): void;
  redo(): void;
  /**
   * Puts the selection on the clipboard, and empties it when cutting.
   *
   * The text comes back on the `clipboard` view key rather than as a
   * return value, because a command has none.
   */
  copy(cut: boolean): void;
  /**
   * Writes clipboard text at the selection's top-left corner.
   *
   * The text crosses raw and is read here, because reading it is the
   * sheet's business: which cell a tab means, whether a block is
   * rectangular, and whether the formulas in it should move are all
   * questions only this side can answer.
   */
  paste(text: string): void;
  /** Empties every cell in the selection. */
  clearRange(): void;
  /** Extends the selection over a cell, repeating it with its formulas moved. */
  fill(toRow: number, toColumn: number): void;
}

export interface SheetView {
  readonly window: SheetWindow;
  readonly geometry: SheetGeometry;
  readonly selection: SheetSelection;
  readonly editor: SheetEditor;
  readonly status: SheetStatus;
  readonly clipboard: SheetClipboard;
}

export const EMPTY_WINDOW: SheetWindow = {
  firstRow: 0,
  lastRow: -1,
  firstColumn: 0,
  lastColumn: -1,
  cells: {}
};

/**
 * One cell out of a window, or null when the window does not cover it.
 *
 * Null is the interesting answer: a cell the render worker has mounted
 * and the application worker has not sent yet. Phase 0's `Missed`
 * readout counted exactly these.
 */
export function cellIn(window: SheetWindow, row: number, column: number): string | null {
  return window.cells[row]?.[column] ?? null;
}

export const Sheet = channel<SheetView, SheetCommands>('sheet', {
  window: EMPTY_WINDOW,
  geometry: { rowCount: 0, columnCount: 0, rowHeight: 24, columnWidth: 104 },
  selection: { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 },
  editor: { row: 0, column: 0, input: '' },
  status: { pending: 0, canUndo: false, canRedo: false },
  clipboard: { text: '', serial: 0 }
});
