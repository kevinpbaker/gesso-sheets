import { channel } from 'gesso-framework';

import { NO_BORDERS, type CellPaint } from '../sheet/Format';
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
 * Phase 9 adds two more again, and they are the clearest case in the
 * file. `formats` is a palette *index* per visible cell and moves
 * whenever the viewport does; `palette` is the table those indices
 * point into and moves only when somebody formats something. Sharing
 * one key, a scroll would put the whole palette back on the wire and
 * a single click on Bold would put the whole window back.
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
  /**
   * The rows that are hidden, in order.
   *
   * A list of the exceptions and not a height per row, because a
   * sheet is ten thousand rows tall and all but a handful of them are
   * the same — which is the same shape `UiVirtualSheet.rowHeights`
   * takes, and for the same reason.
   */
  readonly hiddenRows: readonly number[];
  /**
   * How many rows and columns stay put while the rest scrolls.
   *
   * The document's, not the screen's, for the reason the column
   * widths are: it survives a reload, and a sheet whose panes came
   * back unfrozen would have lost something somebody set up.
   */
  readonly frozenRows: number;
  readonly frozenColumns: number;
  /**
   * The merged rectangles, all of them.
   *
   * On the geometry because that is what they are — a merge changes
   * where things are drawn and nothing about what a cell holds — and
   * because there are tens of them at most. The render worker needs
   * every one of them rather than the ones in view: a merge anchored
   * above the window still has to paint into it, which is the whole
   * reason `extendRange` exists.
   */
  readonly merges: readonly SheetMerge[];
}

export interface SheetMerge {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
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
  /**
   * Why this cell is showing an error, when it is.
   *
   * Travels with the editor rather than as a key of its own because
   * it changes exactly when the active cell does, and a key that
   * changes with another key is a second publish for one event.
   *
   * Computed on the application worker, because finding the cell that
   * *made* the error is a walk back through the dependency graph and
   * the graph is not on the wire. The render worker gets a sentence
   * and an address, which is all it draws.
   */
  readonly explain: SheetExplain | null;
}

/**
 * The names a sheet knows, and what the last attempt to add one said.
 *
 * `refused` is a sentence or the empty string, and it travels here
 * rather than as its own key because it is only ever read beside the
 * list: it answers "did that work", which is a question about this
 * list at this moment.
 */
export interface SheetNames {
  readonly entries: readonly SheetName[];
  readonly refused: string;
}

export interface SheetName {
  readonly name: string;
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

export interface SheetExplain {
  readonly code: string;
  readonly meaning: string;
  /** The cell that produced it, already written as `B7`, or null. */
  readonly blame: string | null;
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

/**
 * Which format each visible cell has, as an index into `palette`.
 *
 * Keyed exactly as `window` is, and an index rather than a record for
 * the reason the window is a map rather than an array: the shape
 * decides the patch count more than the contents do. A column of
 * fifty thousand cells formatted as currency is one palette entry
 * and one small integer per cell *in view*; a record per cell would
 * put a nested object on the wire for every cell in the window and
 * give the differ one to walk on every publish, to discover that all
 * of them are identical.
 *
 * A key that is absent is a cell with the default format, which is
 * also what index 0 means — so an unformatted sheet sends nothing at
 * all here, forever.
 */
export interface SheetFormatWindow {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  readonly cells: Readonly<Record<string, Readonly<Record<string, number>>>>;
}

/**
 * The formats the indices point at — the paint half only.
 *
 * The number format is deliberately not here. Turning 1234.5 into
 * `$1,234.50` happens on the application worker and what crosses is
 * the finished string, so the render worker never learns a locale, a
 * currency symbol or a thousands separator. It learns that a cell is
 * bold.
 */
export interface SheetPalette {
  readonly entries: readonly CellPaint[];
}

/**
 * The strings a column would have to be wide enough for.
 *
 * Autofit is the one thing neither thread can do alone. The
 * application worker knows every string in a column and nothing about
 * fonts; the render worker knows the font and holds thirty rows. So
 * this side narrows a million cells to a handful of candidates — by
 * character count, which is the right *shortlist* even though it is
 * the wrong *answer* in a proportional font — and the render worker
 * measures those exactly.
 *
 * `serial` is what makes asking twice a change, for the reason
 * `SheetClipboard` has one: the same answer twice is structurally
 * equal, the differ says nothing happened, and the second autofit
 * does nothing.
 */
export interface SheetAutofit {
  readonly serial: number;
  readonly columns: readonly SheetAutofitColumn[];
}

export interface SheetAutofitColumn {
  readonly column: number;
  /** The longest strings in it, longest first. */
  readonly samples: readonly string[];
  /** Whether any of them is a heading, which is drawn bold. */
  readonly bold: readonly boolean[];
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
  /**
   * The range the render worker has mounted, on the sheet it is
   * mounted over.
   *
   * The viewport **names its sheet**, and that is how the application
   * worker knows which one is showing. The alternative was a sheet
   * argument on every one of the forty commands below, which says the
   * same thing forty times and gets it wrong once.
   */
  setViewport(sheet: number, firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void;
  /** Shows a sheet, without waiting for its viewport to arrive. */
  activateSheet(sheet: number): void;
  /** Adds a sheet at the end and shows it. */
  addSheet(): void;
  renameSheet(sheet: number, name: string): void;
  removeSheet(sheet: number): void;
  moveSheet(from: number, to: number): void;
  duplicateSheet(sheet: number): void;
  setSheetColour(sheet: number, colour: string | null): void;
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
  /**
   * Applies a change to every cell in the selection.
   *
   * A *change*, not a format: `{ bold: true }` means "make these
   * bold" and has to leave the italics, the currency symbol and the
   * alignment of each cell as they were. Sending a whole format
   * would make every button on the toolbar destroy what the others
   * had done, which is the bug every naive formatting model has.
   */
  format(change: SheetFormatChange): void;
  /** Puts the selection back to the default format. */
  clearFormat(): void;
  /**
   * Inserts or deletes whole rows or columns at the selection.
   *
   * `count` is how many, taken from how many the selection covers, so
   * selecting three rows and inserting gives three. The sheet keeps
   * its size: an insert pushes the last rows off the end and a delete
   * brings empty ones in at it, which is what a sheet of a fixed
   * extent means.
   */
  insertRows(at: number, count: number): void;
  deleteRows(at: number, count: number): void;
  insertColumns(at: number, count: number): void;
  deleteColumns(at: number, count: number): void;
  /**
   * Draws borders over the selection.
   *
   * A named pattern rather than four edges, because the interesting
   * part is what "outline" means over a *range*: the outer edge of
   * the block and not a box round every cell in it. Only this side
   * knows where the block's edges are.
   */
  setBorders(pattern: BorderPattern, width: number, color: string): void;
  /**
   * Reorders the rows of the selection by one of its columns.
   *
   * Whole rows, always. Sorting one column of a block and leaving the
   * rest where it was is the most destructive thing a spreadsheet can
   * do quietly, so the selection *is* the block — a selection of one
   * cell is widened to the run of columns around it before this is
   * sent, which the screen does because only it knows what somebody
   * can see.
   */
  sortRange(column: number, ascending: boolean, hasHeader: boolean): void;
  /**
   * Hides the columns the selection covers, or shows what is hidden.
   *
   * A hidden column is one of width zero, which is all it takes: the
   * widths are already an array and the offsets already a prefix sum.
   * Rows cannot be hidden the same way — `LazySheet` takes one height
   * for every row — and that is recorded with the phase.
   */
  hideColumns(first: number, last: number): void;
  showColumns(first: number, last: number): void;
  hideRows(first: number, last: number): void;
  showRows(first: number, last: number): void;
  /**
   * Keeps the first `rows` rows and `columns` columns on screen.
   *
   * Counted from the top-left rather than given as a cell, because
   * that is what it means — "everything above and to the left of
   * here" — and a screen that sent a cell would be sending a
   * coordinate to describe a quantity.
   */
  freeze(rows: number, columns: number): void;
  /**
   * Merges the selection into one cell, or takes a merge apart.
   *
   * Destructive, and knowingly: everything but the top-left cell
   * loses what it held, because a merged cell has nowhere to show it.
   * Every spreadsheet warns about this and this one does it as one
   * step of undo instead, which is the same promise kept differently.
   */
  mergeCells(): void;
  /**
   * Gives the selection a name.
   *
   * A command rather than a question, so the render worker does not
   * have to hold the rules: the answer comes back as the names view
   * changing, or not changing, and the sentence explaining why is
   * published beside it.
   */
  defineName(name: string): void;
  removeName(name: string): void;
  unmergeCells(): void;
  /**
   * Asks what the columns would have to be wide enough for.
   *
   * A request one way and an answer the other, on the `autofit` view
   * key — the shape `copy` already uses, because a command has no
   * return value and only the render worker can finish the job.
   */
  measureColumns(first: number, last: number): void;
  /**
   * Keeps the rows whose cell in this column matches the one the
   * selection is on, and hides the rest.
   *
   * A snapshot rather than a rule: an edit afterwards does not re-run
   * it, which is what every spreadsheet does and what keeps an edit
   * from making rows vanish under somebody's hands.
   */
  filterToSelection(): void;
  clearFilter(): void;
}

/** What a border command draws. */
export type BorderPattern = 'all' | 'outline' | 'top' | 'bottom' | 'none';

/**
 * What one press of a formatting control means.
 *
 * Every field optional and every one of them meaning "leave it
 * alone" when absent. `number` carries the whole number format
 * because its parts are not independent — places belong to a
 * currency format, not to a cell — and the paint fields are
 * independent and so are listed one by one.
 */
export interface SheetFormatChange {
  readonly number?: NumberFormatPatch;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly fontSize?: number;
  readonly color?: string;
  readonly fill?: string;
  readonly align?: 'auto' | 'start' | 'center' | 'end';
  readonly wrap?: boolean;
  /**
   * Which edges to set, and to what.
   *
   * An edge named is an edge changed; an edge absent is left alone,
   * on the same rule as everything else here — putting a rule under a
   * row must not remove the box somebody drew around it.
   */
  readonly borders?: {
    readonly top?: SheetEdge;
    readonly right?: SheetEdge;
    readonly bottom?: SheetEdge;
    readonly left?: SheetEdge;
  };
  /** More or fewer decimal places, relative to what each cell has. */
  readonly places?: number;
}

/** One edge, as it crosses. Width 0 is no border. */
export interface SheetEdge {
  readonly width: number;
  readonly color: string;
}

/**
 * A number format as it crosses: plain data, so a tagged object and
 * not a class, and named so the render worker can ask for one
 * without importing the engine's own type.
 */
export type NumberFormatPatch =
  | { readonly kind: 'general' }
  | { readonly kind: 'number'; readonly places: number; readonly thousands: boolean }
  | { readonly kind: 'currency'; readonly places: number; readonly symbol: string }
  | { readonly kind: 'percent'; readonly places: number }
  | { readonly kind: 'scientific'; readonly places: number }
  | { readonly kind: 'date'; readonly pattern: 'ymd' | 'dmy' | 'mdy' }
  | { readonly kind: 'time'; readonly pattern: 'hm' | 'hms' }
  | { readonly kind: 'datetime'; readonly date: 'ymd' | 'dmy' | 'mdy'; readonly time: 'hm' | 'hms' }
  | { readonly kind: 'text' };

/** The tab strip, as the render worker draws it. */
export interface SheetTabs {
  readonly entries: readonly SheetTab[];
  /** Which one is showing, as an index into `entries`. */
  readonly active: number;
}

export interface SheetTab {
  readonly name: string;
  /** A tab colour somebody chose, or null for the plain one. */
  readonly colour: string | null;
}

export interface SheetView {
  readonly window: SheetWindow;
  /**
   * The tabs along the bottom: every sheet, and which one is shown.
   *
   * A list and an index rather than a key per sheet, because it is
   * one question — "what are the tabs" — and a workbook holds tens of
   * sheets rather than thousands. It changes when somebody adds,
   * renames, moves, colours or removes one, which is a thing a person
   * does by hand.
   *
   * What it deliberately does **not** carry is anything about the
   * sheets nobody is looking at. The window, the formats, the
   * geometry and the rest are all the *active* sheet's, so a formula
   * depending on fifty thousand cells on another sheet publishes
   * nothing at all while that sheet is out of view — which is the
   * claim this phase exists to defend.
   */
  readonly sheets: SheetTabs;
  readonly geometry: SheetGeometry;
  readonly selection: SheetSelection;
  readonly editor: SheetEditor;
  /** The named ranges, for the name box to resolve and to list. */
  readonly names: SheetNames;
  readonly status: SheetStatus;
  readonly clipboard: SheetClipboard;
  /** Sum, average and count over the selection. */
  readonly stats: SheetStats;
  readonly find: SheetFindView;
  readonly formats: SheetFormatWindow;
  readonly palette: SheetPalette;
  /**
   * The active cell's own format, for the toolbar to show itself
   * pressed with.
   *
   * One cell and not the selection's, because a toolbar has one Bold
   * button and a selection can hold both. Every spreadsheet answers
   * this from the active cell and this one does too.
   */
  readonly activeFormat: SheetActiveFormat;
  readonly autofit: SheetAutofit;
}

/** What the controls read to draw themselves. */
export interface SheetActiveFormat {
  readonly paint: CellPaint;
  readonly number: NumberFormatPatch;
}

export const EMPTY_FORMATS: SheetFormatWindow = {
  firstRow: 0,
  lastRow: -1,
  firstColumn: 0,
  lastColumn: -1,
  cells: {}
};

export const PLAIN_PAINT: CellPaint = {
  bold: false,
  italic: false,
  underline: false,
  fontSize: 0,
  color: '',
  fill: '',
  align: 'auto',
  wrap: false,
  borders: NO_BORDERS
};

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
  sheets: { entries: [{ name: 'Sheet1', colour: null }], active: 0 },
  geometry: {
    rowCount: 0,
    columnCount: 0,
    rowHeight: 24,
    columnWidth: 104,
    columnWidths: [],
    hiddenRows: [],
    frozenRows: 0,
    frozenColumns: 0,
    merges: []
  },
  selection: { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 },
  editor: { row: 0, column: 0, input: '', explain: null },
  names: { entries: [], refused: '' },
  status: { pending: 0, evaluated: 0, canUndo: false, canRedo: false },
  clipboard: { text: '', serial: 0 },
  stats: NO_STATS,
  find: NO_FIND,
  formats: EMPTY_FORMATS,
  palette: { entries: [PLAIN_PAINT] },
  activeFormat: { paint: PLAIN_PAINT, number: { kind: 'general' } },
  autofit: { serial: 0, columns: [] }
});
