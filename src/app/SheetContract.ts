import { channel } from 'gesso-framework';
import type { ColourScale, ConditionalPaint, ConditionalTest } from '../sheet/Conditional';
import type { ValidationRule } from '../sheet/Validation';
import type { ChartKind } from '../sheet/Chart';
import type { Series } from '../sheet/Series';

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
   * The rows that are not the default height, as `[row, height]`, in
   * order: set by hand, or fitted to what they hold. Sparse for the
   * reason the hidden rows are. A hidden row is still hidden whatever
   * height it has here, and shows at that height again when shown.
   */
  readonly rowHeights: readonly (readonly [number, number])[];
  /**
   * How many rows and columns stay put while the rest scrolls.
   *
   * The document's, not the screen's, for the reason the column
   * widths are: it survives a reload, and a sheet whose panes came
   * back unfrozen would have lost something somebody set up.
   */
  readonly frozenRows: number;
  readonly frozenColumns: number;
  /** How large the sheet is drawn, 1 being 100%; see `SheetCommands.setZoom`. */
  readonly zoom: number;
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
/**
 * A range and the active cell inside it, as Excel has them.
 *
 * `row` and `column` are the **active cell**: where typing goes, what
 * the name box and the formula bar show, where the editor opens. The
 * range runs from the anchor to the **corner**, which is the end Shift
 * moves. The corner is absent when it is the active cell — which it is
 * after a click, an arrow, or any selection made before Phase 24 — so
 * every selection written the old way still means what it meant.
 *
 * What wants the rectangle asks `rectOf` in `SheetRanges`, or
 * `cornerOf` here; what wants the active cell reads `row` and `column`.
 */
export interface SheetSelection {
  readonly row: number;
  readonly column: number;
  readonly anchorRow: number;
  readonly anchorColumn: number;
  readonly cornerRow?: number;
  readonly cornerColumn?: number;
}

/** The end of the range Shift moves: the corner when there is one, the active cell otherwise. */
export function cornerOf(selection: SheetSelection): { row: number; column: number } {
  return { row: selection.cornerRow ?? selection.row, column: selection.cornerColumn ?? selection.column };
}

/** Whether the range is one cell. */
export function isOneCell(selection: SheetSelection): boolean {
  const corner = cornerOf(selection);
  return corner.row === selection.anchorRow && corner.column === selection.anchorColumn;
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
  /**
   * The formula whose array this cell is showing a piece of, when it
   * is one — with nothing typed in it, and an answer that came from
   * somewhere. The formula bar shows it greyed, as Excel's does, and
   * will not be typed into as if it were this cell's.
   */
  readonly spilledFrom: { readonly row: number; readonly column: number; readonly input: string } | null;
  /** The note on this cell, or empty; what Shift+F2 opens with. */
  readonly note: string;
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
  /** The names that hold formulas, which have no range to go to. */
  readonly formulas: readonly SheetFormulaName[];
  readonly refused: string;
}

export interface SheetFormulaName {
  readonly name: string;
  /** As written, `=` and all. */
  readonly formula: string;
}

export interface SheetName {
  readonly name: string;
  /** What it holds, as the Names dialog shows it and takes it back: `=Sales!$A$4:$A$27`. */
  readonly refersTo: string;
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
  /**
   * The block last copied or cut, while it is still what Ctrl+V would
   * paste: drawn with a moving dashed outline until it is pasted (a
   * cut), Escape is pressed, or the sheet is edited. Null otherwise.
   */
  readonly marked: SheetMarked | null;
}

export interface SheetMarked {
  readonly sheet: number;
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  readonly cut: boolean;
}

/** What Paste special pastes; see `SheetRanges.pasteCopied`. */
export type SheetPasteMode = 'all' | 'values' | 'formats' | 'transposed';

/**
 * Files on their way in and out.
 *
 * A download is a request one way and an answer the other, on the
 * shape `SheetClipboard` already uses: the render worker asks for the
 * CSV, the application worker builds it and publishes it here, and
 * the render worker hands it to the shell — the only thread that can
 * put a file in front of somebody. `serial` makes exporting the same
 * sheet twice a change, for the reason the clipboard's does.
 *
 * `report` is what the last file opened turned into, as a sentence,
 * or the empty string: it answers "did that work", and travels here
 * because it is only ever read beside the download it is not.
 */
/**
 * Which document the tab is showing, and whether it has a file.
 *
 * Its own key because it changes at its own rate — when a document is
 * opened, saved, or first edited after either — and the title, the
 * route and the File menu all read it.
 */
export interface SheetDocumentView {
  /** The library's id for it; the route is `/d/<id>`. Empty until one is open. */
  readonly id: string;
  readonly name: string;
  /** The file it was saved to or opened from, or null. */
  readonly file: { readonly handle: number | null; readonly name: string } | null;
  /** Changed since it was opened or last saved to its file. */
  readonly edited: boolean;
  /**
   * Open in another tab, which holds it: this tab shows it and does not
   * keep it, because two copies kept at once would each write over the
   * other's edits.
   */
  readonly elsewhere: boolean;
}

export interface SheetTransfer {
  readonly download: SheetDownload | null;
  readonly report: string;
}

export interface SheetDownload {
  readonly serial: number;
  /** What it is, which decides what the picker offers and what a save means. */
  readonly kind: 'workbook' | 'csv' | 'xlsx';
  /**
   * The shell's handle to write to, for a Save; null for a Save As or
   * an export, which ask the shell where.
   */
  readonly handle: number | null;
  /** A file name to suggest, extension included. */
  readonly name: string;
  readonly mediaType: string;
  readonly text: string;
  /**
   * The file's bytes, for one that is not text — an `.xlsx` is a zip —
   * as base64, because a command and a view carry plain data.
   */
  readonly base64?: string;
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
/**
 * What a cell is not allowed to hold, and what it may.
 *
 * The marks are keyed like the format window and for the same
 * reason: nested records diff structurally, so a cell that starts or
 * stops breaking its rule is one patch and the rest of the window is
 * silent.
 *
 * Only the cells *in view* are marked, which is the whole shape of
 * this phase — a rule over a million cells is asked about the ones
 * somebody can see.
 */
export interface SheetValidation {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly cells: Readonly<Record<string, Readonly<Record<string, string>>>>;
  /**
   * Why the last commit was refused, or the empty string.
   *
   * Travels with the marks rather than as its own key because it is
   * only ever read beside them: it answers "did that go in", which is
   * a question about this rule at this moment.
   */
  readonly refused: string;
  /**
   * The values the active cell may take, when its rule is a list.
   *
   * Empty for every other kind, which is what the cell editor reads
   * to decide whether there is a dropdown to offer.
   */
  readonly list: readonly string[];
}

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
/**
 * What a cell being typed into could be finished as: AutoComplete.
 *
 * A question and an answer, like `autofit`, because the render thread
 * holds a window of the column and the words are in all of it. The
 * answer names the cell and the prefix it was asked about, so one that
 * arrives after the typing has moved on is recognised and dropped.
 * `text` is the whole word, spelled as it was first typed, or empty.
 */
/**
 * The notes on the cells in view, by row and then column.
 *
 * Proportional to the window like `formats`: a sheet with no notes on
 * screen publishes an empty object, and the differ says nothing about
 * it after the first time.
 */
/** The zooms View offers, smallest first. */
export const ZOOMS: readonly number[] = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2];
/** The next zoom along from `now`, or 100%. */
export function zoomStep(now: number, id: 'zoomIn' | 'zoomOut' | 'zoomReset'): number {
  if (id === 'zoomReset') {
    return 1;
  }
  if (id === 'zoomIn') {
    return ZOOMS.find(level => level > now + 1e-9) ?? ZOOMS[ZOOMS.length - 1];
  }
  return [...ZOOMS].reverse().find(level => level < now - 1e-9) ?? ZOOMS[0];
}

export interface SheetNotes {
  readonly cells: Readonly<Record<string, Readonly<Record<string, string>>>>;
}

export interface SheetCompletion {
  readonly serial: number;
  readonly row: number;
  readonly column: number;
  readonly prefix: string;
  readonly text: string;
}

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

/**
 * The rows whose height depends on what they hold, and what that is.
 *
 * The same split as autofit, in the other direction: this side knows
 * which cells wrap and how wide their columns are, and only the render
 * worker can say how many lines a string takes at a width. A row is
 * sent with the cells that could make it taller than one line — the
 * wrapped ones and the ones in a large font — and a row sent with none
 * is one that should go back to the default.
 *
 * `serial` for the reason `SheetAutofit` has one; the answer comes
 * back as `fitRows` with it, so an answer to a question since replaced
 * can be told apart and dropped.
 */
export interface SheetRowFit {
  readonly serial: number;
  readonly rows: readonly SheetRowFitRow[];
}

export interface SheetRowFitRow {
  readonly row: number;
  readonly cells: readonly SheetRowFitCell[];
}

export interface SheetRowFitCell {
  readonly text: string;
  /** The column's width, which is what wrapped text breaks at. */
  readonly width: number;
  /** In pixels; zero for the default. */
  readonly fontSize: number;
  readonly bold: boolean;
  readonly wrap: boolean;
}

export interface SheetStatus {
  /** Cells whose value is still out of date. Zero when settled. */
  readonly pending: number;
  /** Cells the application thread has evaluated, ever. */
  readonly evaluated: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /**
   * What Undo and Redo would do, as the menu and the tooltip say it:
   * `sort`, `typing in B4`. Empty when there is nothing to do.
   */
  readonly undoLabel: string;
  readonly redoLabel: string;
  /** Whether circular formulas are gone round rather than refused. */
  readonly iterating: boolean;
  /** Whether the sheet is showing its formulas instead of their answers. */
  readonly showingFormulas: boolean;
}

/**
 * A conditional rule as it crosses the barrier.
 *
 * Plain data, because a `postMessage` carries no closures — which is
 * also why the tests in `Conditional.ts` are a tagged union rather
 * than a predicate. The range is the *selection's*, filled in on the
 * other side, so the command says what the rule is and not where.
 */
export interface SheetConditionalRule {
  readonly test: ConditionalTest | null;
  readonly paint?: ConditionalPaint;
  readonly scale?: ColourScale;
}

export type SheetValidationRule = ValidationRule;

/**
 * The rules the active cell is under, for the rules bar to open on.
 *
 * One cell and not the selection, for the reason `activeFormat` is:
 * the bar has one set of controls and a selection can straddle three
 * rules. The conditional is the *last* one covering the cell, because
 * later rules paint over earlier ones and the one on top is the one
 * somebody is looking at; the validation is the first, because that
 * is the one `validationAt` enforces.
 */
export interface SheetActiveRules {
  readonly conditional: SheetConditionalRule | null;
  readonly validation: { readonly rule: SheetValidationRule; readonly strict: boolean } | null;
}

/**
 * A chart on the sheet: what it is, and where it floats.
 *
 * The placement is in the sheet's own pixels, so the render worker
 * subtracts the scroll and draws. The *range* travels as text rather
 * than as a `RangeRef` because it is here to be shown — under the
 * title, and in whatever names the chart in a list — and the render
 * worker has no business resolving a reference.
 */
export interface SheetChart {
  readonly id: number;
  readonly kind: ChartKind;
  readonly title: string;
  readonly range: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly legend: boolean;
}

export interface SheetCharts {
  readonly entries: readonly SheetChart[];
  /** The selected chart's id, or zero when none is. */
  readonly selected: number;
}

/**
 * The numbers behind the charts, at the resolution they can draw.
 *
 * **Its own key, and that is the whole of why there are two.** What a
 * chart *is* changes when somebody drags it; what a chart *shows*
 * changes when a cell it reads is edited, which on a recalculating
 * sheet is every frame. One key carrying both would republish four
 * hundred points because a title was renamed, and republish a title
 * because a number moved — and the differ would send every byte of it
 * either way, since a list of points is not a structure it can look
 * inside.
 *
 * Keyed by the chart's id as a string, because that is what a record
 * key is once it has crossed a `postMessage`.
 */
export interface SheetSeriesView {
  readonly charts: Readonly<Record<string, SheetChartSeries>>;
}

export interface SheetChartSeries {
  readonly categories: readonly string[];
  readonly series: readonly Series[];
  /**
   * How many points were read before thinning, so a chart can say so.
   *
   * The honest label for a downsampled line, and the number the phase
   * is measured on: `read` in the thousands beside a `points` length
   * bounded by the chart's width is the exit criterion, visible.
   */
  readonly read: number;
  /**
   * Where the chart's numbers come from, in the three parts it reads
   * them as, so the grid can outline them while the chart is selected —
   * as Excel does: the series' names, the categories along the axis,
   * and the values. Null for a chart whose sheet is gone.
   */
  readonly source: SheetChartSource | null;
}

export interface SheetChartSource {
  /** The sheet the cells are on, by index, which is the tabs' `active` when it is this one. */
  readonly sheet: number;
  readonly names: SheetRect | null;
  readonly categories: SheetRect | null;
  readonly values: SheetRect | null;
}

export interface SheetRect {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

/**
 * The workbook's scripts, and what the last run came to.
 *
 * Its own key, since it changes when somebody saves a script or runs
 * one, which is rarely, and the editor is the only reader.
 */
export interface SheetScripts {
  readonly entries: readonly SheetScript[];
  /** The script running now, by name, or the empty string. */
  readonly running: string;
  /** What the last attempt to save a script said, or the empty string. */
  readonly refused: string;
  /** How the last run ended, or null before there has been one. */
  readonly last: SheetScriptRun | null;
}

export interface SheetScript {
  readonly name: string;
  readonly source: string;
  /** The empty string for a script written here, or the file it came with. */
  readonly from: string;
  /** A script somebody runs, or one whose functions formulas call. */
  readonly kind: 'run' | 'functions';
  /** For functions: the names a formula can call, as the script spells them. */
  readonly defines: readonly string[];
  /** For functions: why one defines nothing, or less than it says. Empty when all is well. */
  readonly problem: string;
  /** For functions from a file: whether they are turned on. Always true for any other script. */
  readonly on: boolean;
}

export interface SheetScriptRun {
  /** Which run, so the same outcome twice is still news. */
  readonly serial: number;
  readonly name: string;
  readonly outcome: 'done' | 'failed' | 'timeout' | 'stopped' | 'refused';
  /** The sentence the editor shows: what it did, or why it did not. */
  readonly text: string;
  /** What the script logged. */
  readonly log: readonly string[];
}

export interface SheetCommands {
  /**
   * The range the render worker has mounted, on the sheet it is
   * mounted over.
   *
   * The viewport **names its sheet** — the one the render worker last
   * heard was showing — so that none of the forty commands below has
   * to. It does not *choose* the sheet: that is `activateSheet`, and a
   * viewport over any other sheet is one that crossed a tab change in
   * flight, answered for the sheet that is showing. Letting it switch
   * made two quick changes a ping-pong neither side could leave.
   */
  setViewport(sheet: number, firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void;
  /**
   * Adds a conditional format over the selection.
   *
   * The rule crosses as data rather than as a closure, because a
   * closure cannot cross a `postMessage` — which is also why the
   * tests are a tagged union rather than a predicate.
   */
  addConditional(rule: SheetConditionalRule): void;
  removeConditional(at: number): void;
  /** Adds a validation over the selection. */
  addValidation(rule: SheetValidationRule, strict: boolean, message: string): void;
  removeValidation(at: number): void;
  /** Takes every rule off the sheet, which is the only bulk one. */
  clearRules(): void;
  /**
   * Opens a CSV as a sheet of its own, named after its file, and
   * shows it.
   *
   * The text crosses rather than the bytes because decoding is the
   * render worker's to do once, where the file arrived; what it means
   * — which separator, which cells are numbers, whether `=1+2` is a
   * formula — is this side's, and it is not.
   */
  importCsv(fileName: string, text: string): void;
  /** Builds the sheet in view as a CSV, and publishes it on `transfer`. */
  exportCsv(): void;
  /** Builds the workbook as an `.xlsx`, and publishes it on `transfer`. */
  exportXlsx(): void;
  /**
   * Shows a document: `''` for the last one used, `'new'` for a blank
   * one, or an id from the route. Opening the one already open does
   * nothing, so the render worker sends this on every route change.
   */
  openDocument(id: string): void;
  /** Opens what a file held — a workbook as a document, a CSV as a sheet. */
  openFile(fileName: string, text: string, handle: number | null): void;
  /**
   * Opens an Excel workbook as a document of its own, from its bytes as
   * base64 — a command carries plain data, and a buffer is not.
   */
  importXlsx(fileName: string, base64: string): void;
  /** Builds the workbook for the shell to save; `asNew` is Save As. */
  saveDocument(asNew: boolean): void;
  /** Where the shell put a download, so a workbook can remember its file. */
  fileSaved(kind: 'workbook' | 'csv' | 'xlsx', name: string, handle: number | null, via: 'file' | 'download'): void;
  /** A sentence about a file that did not go where it was sent, for the status line. */
  reportFile(text: string): void;
  /**
   * Puts a chart over the selection, and selects it.
   *
   * The range is the selection's, like every other command that acts
   * on one: the render worker says what kind of chart, and where the
   * cells are is already a thing both sides agree on.
   */
  insertChart(kind: ChartKind): void;
  /** Which chart the handles are drawn around, or zero for none. */
  selectChart(id: number): void;
  /**
   * Where a chart sits and how big it is, in the sheet's own pixels.
   *
   * One command for both, because a resize from a corner handle moves
   * the origin as well as the size, and two commands would publish a
   * frame in which the chart had the new size at the old position.
   * The drag itself stays on the render thread and this arrives when
   * the pointer is let go, which is the same trade a column resize
   * makes.
   */
  placeChart(id: number, x: number, y: number, width: number, height: number): void;
  setChartKind(id: number, kind: ChartKind): void;
  setChartTitle(id: number, title: string): void;
  /**
   * The cells a chart reads, changed: its outline dragged on the sheet.
   * On the sheet the chart already reads, which a drag cannot leave.
   */
  setChartRange(id: number, firstRow: number, firstColumn: number, lastRow: number, lastColumn: number): void;
  removeChart(id: number): void;
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
  /**
   * The active cell, the anchor, and — when the range's far end is not
   * the active cell — the corner; see `SheetSelection`.
   */
  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number, cornerRow?: number, cornerColumn?: number): void;
  /**
   * Ctrl+Arrow: moves the cursor to the edge of the data in a
   * direction, keeping the anchor with `extend`.
   *
   * The selection it starts from travels with it, because the render
   * thread's selection leads this side's by a frame and a key pressed
   * inside that frame would otherwise jump from where the cursor was
   * rather than where it is. The answer comes back on `selection`.
   */
  jumpToEdge(
    row: number,
    column: number,
    anchorRow: number,
    anchorColumn: number,
    rows: -1 | 0 | 1,
    columns: -1 | 0 | 1,
    extend: boolean
  ): void;
  /**
   * Ctrl+Enter: what was typed, into every cell of the selection, its
   * references moved as a fill would move them from the cursor's cell.
   * One step of undo, and the selection stays where it is.
   */
  writeSelection(input: string): void;
  /**
   * Asks what typing `prefix` into a cell would complete to; the answer
   * comes back on `completion` with the same serial.
   */
  complete(row: number, column: number, prefix: string, serial: number): void;
  /** Writes a cell's note, or takes it away with empty text. One step of undo. */
  setNote(row: number, column: number, text: string): void;
  undo(): void;
  redo(): void;
  /**
   * Puts the selection on the clipboard. A cut is only marked: the
   * cells move when it is pasted, and the references into them move
   * with them.
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
  paste(text: string, mode?: SheetPasteMode): void;
  /**
   * Pastes what this sheet last copied, in one of Paste special's ways,
   * without the system clipboard — which a menu cannot read.
   */
  pasteSpecial(mode: SheetPasteMode): void;
  /**
   * The selection dragged by its border and let go with its top-left
   * corner at a cell: a move, as a cut pasted there would be, or a copy
   * with Ctrl held. The clipboard is not touched.
   */
  moveRange(row: number, column: number, copy: boolean): void;
  /** Escape over a marked copy: the outline goes, and a cut is called off. */
  unmark(): void;
  /** The format painter picks up the selection's formats, without touching the clipboard. */
  pickFormats(): void;
  /** And puts them on the selection, repeated across it when it is larger. */
  paintFormats(): void;
  /** Empties every cell in the selection. */
  clearRange(): void;
  /** Extends the selection over a cell, repeating it with its formulas moved. */
  fill(toRow: number, toColumn: number): void;
  /** The fill handle's double-click: down as far as the data beside the selection goes. */
  fillToData(): void;
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
   * A row was dragged to a new height, which it keeps whatever it
   * holds. Sent when the drag ends, as a column's width is. A height
   * of zero or less means "fit it to what it holds" again.
   */
  setRowHeight(row: number, height: number): void;
  /**
   * Forgets the heights set by hand in these rows, so they fit what
   * they hold again — the way back from a drag.
   */
  fitRowsToContents(first: number, last: number): void;
  /**
   * Turns Excel's iterative calculation on or off for the workbook:
   * circular formulas gone round a hundred times, or until nothing
   * moves by a thousandth, rather than refused as `#CIRC!`.
   */
  setIteration(on: boolean): void;
  /**
   * Draws every formula instead of its answer, or goes back. A way of
   * looking and not an edit: it is not on the undo stack and is not
   * saved, as it is not in Excel's own undo.
   */
  showFormulas(on: boolean): void;
  /**
   * Draws the sheet larger or smaller: 1 is 100%, held between the
   * zooms `ZOOMS` offers. Saved with the sheet, and not an edit.
   */
  setZoom(zoom: number): void;
  /**
   * The heights the rows in a `rowFit` request need, measured.
   *
   * Only rows nobody has set a height for take one; a row set by hand
   * has been told what it is.
   */
  fitRows(serial: number, heights: readonly (readonly [number, number])[]): void;
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
  /** The proof page's other instrument; see `SheetService.chartStress`. */
  chartStress(points: number): void;
  /** And the third; see `SheetService.scriptStress`. */
  scriptStress(): void;
  /** And the fourth; see `SheetService.functionStress`. */
  functionStress(calls: number): void;
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
  /**
   * Gives a name a formula to hold: `=0.2`, or `=LAMBDA(x, x*2)`, which
   * makes a function called by that name. Answered on the names view
   * as `defineName` is.
   */
  defineFormulaName(name: string, formula: string): void;
  /**
   * The Names dialog's save: `was` is the name being edited, empty for
   * a new one; `refersTo` a range or a formula, as written. Refused on
   * the names view, as the name box's gesture is.
   */
  saveName(was: string, name: string, refersTo: string): void;
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
  /**
   * Saves a script: a new one when `was` is empty, or the one called
   * `was`, renamed to `name` if that differs. A script keeps where it
   * came from when it is edited, so a file's script does not become
   * the person's by having a character changed.
   */
  saveScript(was: string, name: string, source: string, kind?: 'run' | 'functions'): void;
  removeScript(name: string): void;
  /**
   * Runs a script. `confirmed` is the person having said yes to a
   * file's script this time; a typed one runs without it.
   */
  runScript(name: string, confirmed: boolean): void;
  stopScript(): void;
  /**
   * Turns on, or off, the functions this workbook's file brought. The
   * decision is kept in this browser's library for this document,
   * never in the file.
   */
  setFunctionsOn(on: boolean): void;
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
  /** CSV out, and what the last CSV in became; see `SheetTransfer`. */
  readonly transfer: SheetTransfer;
  /** Which document this is; see `SheetDocumentView`. */
  readonly document: SheetDocumentView;
  /** Sum, average and count over the selection. */
  readonly stats: SheetStats;
  readonly find: SheetFindView;
  readonly formats: SheetFormatWindow;
  readonly palette: SheetPalette;
  /** Markers for the cells in view that break a rule; see `SheetValidation`. */
  readonly validation: SheetValidation;
  /**
   * The active cell's own format, for the toolbar to show itself
   * pressed with.
   *
   * One cell and not the selection's, because a toolbar has one Bold
   * button and a selection can hold both. Every spreadsheet answers
   * this from the active cell and this one does too.
   */
  readonly activeFormat: SheetActiveFormat;
  /** The rules over the active cell; see `SheetActiveRules`. */
  readonly activeRules: SheetActiveRules;
  readonly autofit: SheetAutofit;
  /** What the cell being typed into could be finished as; see `SheetCompletion`. */
  readonly completion: SheetCompletion;
  /** The notes on the cells in view; see `SheetNotes`. */
  readonly notes: SheetNotes;
  /** The rows whose height the render worker should work out; see `SheetRowFit`. */
  readonly rowFit: SheetRowFit;
  /** The charts floating over this sheet; see `SheetCharts`. */
  readonly charts: SheetCharts;
  /** What those charts draw; see `SheetSeriesView` for why it is apart. */
  readonly series: SheetSeriesView;
  readonly scripts: SheetScripts;
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
    rowHeights: [],
    frozenRows: 0,
    frozenColumns: 0,
    zoom: 1,
    merges: []
  },
  selection: { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 },
  editor: { row: 0, column: 0, input: '', explain: null, spilledFrom: null, note: '' },
  names: { entries: [], formulas: [], refused: '' },
  status: { pending: 0, evaluated: 0, canUndo: false, canRedo: false, undoLabel: '', redoLabel: '', iterating: false, showingFormulas: false },
  clipboard: { text: '', serial: 0, marked: null },
  transfer: { download: null, report: '' },
  document: { id: '', name: '', file: null, edited: false, elsewhere: false },
  stats: NO_STATS,
  find: NO_FIND,
  formats: EMPTY_FORMATS,
  palette: { entries: [PLAIN_PAINT] },
  validation: { firstRow: 0, lastRow: -1, cells: {}, refused: '', list: [] },
  activeFormat: { paint: PLAIN_PAINT, number: { kind: 'general' } },
  activeRules: { conditional: null, validation: null },
  autofit: { serial: 0, columns: [] },
  completion: { serial: 0, row: 0, column: 0, prefix: '', text: '' },
  notes: { cells: {} },
  rowFit: { serial: 0, rows: [] },
  charts: { entries: [], selected: 0 },
  series: { charts: {} },
  scripts: { entries: [], running: '', refused: '', last: null }
});
