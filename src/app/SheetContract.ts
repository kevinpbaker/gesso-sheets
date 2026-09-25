import { channel } from 'gesso-framework';

import { NO_STATS, type SheetStats } from './Statistics';

/**
 * The barrier.
 *
 * Imported by the render worker and by the application worker, and
 * holding nothing but names and shapes — the store, the parser, the
 * dependency graph and the recalc are behind it and the render worker
 * never loads a line of them.
 *
 * A view key per thing that changes at its own rate, rather than one
 * object, for the reason `NotesContract.ts` splits `rows` from
 * `open`: the differ walks a projection structurally on every
 * publish, so two things that change at different rates must not
 * share a key. A keystroke moves `window` and `editor`; it does not
 * touch `geometry`, and the differ should not have to walk a hundred
 * column widths to find that out.
 *
 * Phase 8 adds two more on the same argument. `stats` moves whenever
 * the selection does, which is on every arrow key; `find` moves only
 * while somebody has the find bar open, which is almost never. Shared
 * with `status` — which moves on every slice of a recalc — either one
 * would be walked tens of times a second to discover it had not
 * changed.
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
  /** The width a column has until somebody drags it. */
  readonly columnWidth: number;
  /**
   * Every column's width, in order from A.
   *
   * On the contract as of Phase 6, and not before. How wide a column
   * is drawn is not the application's business — which is why the drag
   * itself stays on the render thread, where it costs no round trip —
   * right up until it has to survive a reload, and then it is the
   * document's.
   */
  readonly columnWidths: readonly number[];
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

/**
 * What the find bar is looking for and how it is going.
 *
 * The matches are a count and a position, not the list. The list can
 * be tens of thousands of cell keys and the render worker has no use
 * for any of them: it shows "3 of 412" and the application worker
 * moves the selection. Sending the list would put the largest thing
 * in the application on the wire to draw eight characters.
 */
export interface SheetFindView {
  readonly query: string;
  readonly matchCase: boolean;
  readonly wholeCell: boolean;
  readonly inFormulas: boolean;
  readonly matches: number;
  /** Which match the selection is on, from one, or zero for none. */
  readonly active: number;
}

export interface SheetStatus {
  /** Cells whose value is still out of date. Zero when settled. */
  readonly pending: number;
  /** Cells the application thread has evaluated, ever. */
  readonly evaluated: number;
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
  /**
   * A column was dragged to a new width.
   *
   * Sent when the drag *ends*, not on every frame of it. The render
   * thread already knows how wide it is drawing the column and does
   * not need an answer from another thread to go on drawing it; what
   * the application worker needs is the number to write down.
   */
  setColumnWidth(column: number, width: number): void;
  /**
   * Builds a chain of dependent cells and then disturbs its head.
   *
   * The proof surface's one command, and the reason it is on the
   * contract rather than in a script: the claim is about what happens
   * on *this* thread while somebody scrolls on the other, and a claim
   * you cannot make happen from the screen is one nobody can check.
   * Built once and bumped on every call after.
   */
  stress(cells: number): void;
  /**
   * Repeats the top row of the selection down it, or the left column
   * across it.
   *
   * Ctrl+D and Ctrl+R, which are muscle memory. On a selection one
   * cell tall or wide they take from the neighbouring cell instead,
   * which is what every spreadsheet does and what makes them usable
   * without selecting anything first.
   */
  fillDown(): void;
  fillRight(): void;
  /**
   * Searches the sheet and moves to the first match.
   *
   * On this side because only this side has the sheet: the render
   * worker holds the thirty rows it has mounted, so a find run there
   * could only ever search what somebody was already looking at.
   */
  find(query: string, matchCase: boolean, wholeCell: boolean, inFormulas: boolean): void;
  /** Moves to the next match, or the previous one. */
  findStep(forward: boolean): void;
  /** Replaces the match the selection is on, and moves to the next. */
  replaceOne(replacement: string): void;
  replaceAll(replacement: string): void;
  /** Closes the search, so the highlight and the count go away. */
  clearFind(): void;
}

export interface SheetView {
  readonly window: SheetWindow;
  readonly geometry: SheetGeometry;
  readonly selection: SheetSelection;
  readonly editor: SheetEditor;
  readonly status: SheetStatus;
  readonly clipboard: SheetClipboard;
  /** Sum, average and count over the selection. */
  readonly stats: SheetStats;
  readonly find: SheetFindView;
}

export const NO_FIND: SheetFindView = {
  query: '',
  matchCase: false,
  wholeCell: false,
  inFormulas: true,
  matches: 0,
  active: 0
};

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
  geometry: { rowCount: 0, columnCount: 0, rowHeight: 24, columnWidth: 104, columnWidths: [] },
  selection: { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 },
  editor: { row: 0, column: 0, input: '' },
  status: { pending: 0, evaluated: 0, canUndo: false, canRedo: false },
  clipboard: { text: '', serial: 0 },
  stats: NO_STATS,
  find: NO_FIND
});
