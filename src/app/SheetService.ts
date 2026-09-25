import { BehaviorSubject, type Observable } from 'rxjs';

import { relativeRef } from '../sheet/A1';
import { addressOf, explainCell } from '../sheet/Explain';
import { nameProblemText } from '../sheet/Names';
import { aggregateOf } from './Aggregate';
import { ConditionalPainter } from './ConditionalPaint';
import type { CellPaint } from '../sheet/Format';
import { validate } from '../sheet/Validation';
import { ROW_HEIGHT, COLUMN_WIDTH, MIN_COLUMN_WIDTH } from './dimensions';
import {
  EMPTY_FORMATS,
  EMPTY_WINDOW,
  NO_FIND,
  PLAIN_PAINT,
  type SheetClipboard,
  type SheetEditor,
  type SheetNames,
  type SheetExplain,
  type SheetFindView,
  type SheetGeometry,
  type SheetSelection,
  type BorderPattern,
  type SheetActiveFormat,
  type SheetAutofit,
  type SheetEdge,
  type SheetFormatChange,
  type SheetFormatWindow,
  type SheetPalette,
  type SheetStatus,
  type SheetTabs,
  type SheetValidation,
  type SheetConditionalRule,
  type SheetValidationRule,
  type SheetWindow
} from './SheetContract';
import { DEFAULT_FORMAT, withPlaces, type CellFormat } from '../sheet/Format';
import type { Shift } from '../sheet/Shift';
import { at, findMatches, replaceIn, stepBack, stepTo, type FindOptions } from './SheetFind';
import { sortRect } from './SheetSort';
import { NO_STATS, type SheetStats } from './Statistics';
import type { SheetDocument } from './SheetDocument';
import { cellKey, columnName } from '../sheet/A1';
import { snapshotOf, applySnapshot, type SheetSnapshot } from './SheetFile';
import type { SheetRepository } from './SheetRepository';
import {
  clearRect,
  copyRect,
  currentRegion,
  fillRect,
  fillTarget,
  looksLikeHeader,
  pasteBlock,
  rectOf,
  type CopyOrigin,
  type Rect
} from './SheetRanges';

/**
 * Runs a continuation later, as a task rather than a microtask.
 *
 * It has to be a *task*. A microtask runs before the thread returns to
 * its message queue, so a recalc sliced across microtasks never lets a
 * `setViewport` command in and is exactly the uninterrupted recalc the
 * budget exists to avoid — the sheet would still go blank, and the
 * slicing would look like it was working.
 */
export type Schedule = (run: () => void) => void;

const defaultSchedule: Schedule = run => {
  setTimeout(run, 0);
};

export interface SheetServiceOptions {
  /** Cells evaluated per slice before the thread is handed back. */
  readonly budget?: number;
  readonly schedule?: Schedule;
  readonly rowCount?: number;
  readonly columnCount?: number;
  /** Where the sheet is kept. Without one it is kept nowhere. */
  readonly repository?: SheetRepository;
}

/**
 * The application worker's view of the sheet, shaped for the wire.
 *
 * This is the last layer before the barrier and the only one that has
 * to care that plain data is all that crosses. Everything below it —
 * the document, the sheet, the parser, the graph — is unaware there is
 * a barrier at all.
 *
 * **The recalc pump is the point of this phase.** An edit marks work
 * and the pump does it a slice at a time, publishing the window after
 * each slice and handing the thread back in between. Phase 0 measured
 * what happens without it: thirty milliseconds of uninterrupted
 * application thread leaves the sheet blank in 89% of frames, because
 * the render worker goes on scrolling at 60fps and asking for windows
 * that nobody is free to serve. A `setViewport` arriving mid-recalc is
 * answered on the spot, ahead of the arithmetic.
 */
export class SheetService {
  readonly window: Observable<SheetWindow>;
  readonly sheets: Observable<SheetTabs>;
  readonly validation: Observable<SheetValidation>;
  readonly geometry: Observable<SheetGeometry>;
  readonly selection: Observable<SheetSelection>;
  readonly editor: Observable<SheetEditor>;
  readonly names: Observable<SheetNames>;
  readonly status: Observable<SheetStatus>;
  readonly clipboard: Observable<SheetClipboard>;
  readonly selectionStats: Observable<SheetStats>;
  readonly findView: Observable<SheetFindView>;
  readonly formats: Observable<SheetFormatWindow>;
  readonly palette: Observable<SheetPalette>;
  readonly activeFormat: Observable<SheetActiveFormat>;
  readonly autofit: Observable<SheetAutofit>;

  /** Slices run, for a spec that wants to know the pump ran at all. */
  readonly stats = { slices: 0, publishes: 0 };

  /** What the conditional formats cost, for the budget spec and nothing else. */
  get painterStats(): { scans: number; scanned: number; evaluations: number } {
    return this.painter.stats;
  }

  private readonly windowSubject = new BehaviorSubject<SheetWindow>(EMPTY_WINDOW);
  private readonly sheetsSubject: BehaviorSubject<SheetTabs>;
  private readonly validationSubject = new BehaviorSubject<SheetValidation>({
    firstRow: 0,
    lastRow: -1,
    cells: {},
    refused: '',
    list: []
  });
  /** Why the last commit was refused, until the next one. */
  private refusal = '';
  /**
   * The conditional formats of the sheet in view, resolved for its
   * window; see `ConditionalPainter`.
   */
  private readonly painter = new ConditionalPainter(() => this.document.sheet);
  /**
   * Paints a rule asked for that the document's palette does not
   * hold, appended after it and never reordered.
   */
  private readonly extraPaints: CellPaint[] = [];
  private readonly extraIds = new Map<string, number>();
  /** The document palette's size when the palette was last sent. */
  private publishedBase = 0;
  private readonly geometrySubject: BehaviorSubject<SheetGeometry>;
  private readonly selectionSubject: BehaviorSubject<SheetSelection>;
  private readonly editorSubject: BehaviorSubject<SheetEditor>;
  private readonly namesSubject: BehaviorSubject<SheetNames>;
  private readonly statusSubject: BehaviorSubject<SheetStatus>;
  private readonly clipboardSubject = new BehaviorSubject<SheetClipboard>({ text: '', serial: 0 });
  private readonly statsSubject = new BehaviorSubject<SheetStats>(NO_STATS);
  private readonly findSubject = new BehaviorSubject<SheetFindView>(NO_FIND);
  private readonly formatsSubject = new BehaviorSubject<SheetFormatWindow>(EMPTY_FORMATS);
  private readonly paletteSubject = new BehaviorSubject<SheetPalette>({ entries: [PLAIN_PAINT] });
  private readonly activeFormatSubject = new BehaviorSubject<SheetActiveFormat>({
    paint: PLAIN_PAINT,
    number: { kind: 'general' }
  });
  private readonly autofitSubject = new BehaviorSubject<SheetAutofit>({ serial: 0, columns: [] });
  private autofitSerial = 0;
  /**
   * The cells the current search matched, as keys, in reading order.
   *
   * Held here and never published. The render worker shows "3 of 412"
   * and lets this side do the moving, so the list — which can be
   * tens of thousands of keys — stays on the side that has a use for
   * it. Recomputed after any edit, because an edit can create a match
   * or destroy one and a stale list steps somebody to a cell that no
   * longer says what they searched for.
   */
  private found: number[] = [];
  /** How wide a hidden column was, so showing it puts that back. */
  private readonly hiddenWidths = new Map<number, number>();
  /**
   * What this sheet last copied, and from where.
   *
   * Kept so that pasting it back knows how far it moved. Text from
   * anywhere else has no origin and is written as it arrived: a block
   * out of Excel means what it says, and moving references that were
   * never relative to this sheet would be inventing an intent.
   */
  private copied: CopyOrigin | null = null;
  private serial = 0;

  private viewport = { firstRow: 0, lastRow: -1, firstColumn: 0, lastColumn: -1 };
  private readonly budget: number;
  private readonly schedule: Schedule;
  private readonly repository: SheetRepository | undefined;
  private pumping = false;
  /**
   * Whether a load has finished.
   *
   * Saving before it has would write an empty sheet over a real one:
   * the seed runs first, the file arrives second, and between them the
   * document is neither.
   */
  private restored = false;
  /** How long the stress chain is, so it is built once. */
  private stressCells = 0;
  private stressRuns = 0;

  constructor(
    private readonly document: SheetDocument,
    options: SheetServiceOptions = {}
  ) {
    this.budget = options.budget ?? 2_000;
    this.schedule = options.schedule ?? defaultSchedule;
    const columnCount = options.columnCount ?? 100;
    this.repository = options.repository;
    document.columnWidths = Array.from({ length: columnCount }, () => COLUMN_WIDTH);
    this.geometrySubject = new BehaviorSubject<SheetGeometry>({
      rowCount: options.rowCount ?? 10_000,
      columnCount,
      rowHeight: ROW_HEIGHT,
      columnWidth: COLUMN_WIDTH,
      columnWidths: document.columnWidths,
      hiddenRows: [],
      frozenRows: 0,
      frozenColumns: 0,
      merges: []
    });
    this.selectionSubject = new BehaviorSubject<SheetSelection>(document.selection);
    this.editorSubject = new BehaviorSubject<SheetEditor>({
      row: document.selection.row,
      column: document.selection.column,
      input: document.activeInput,
      explain: null
    });
    this.sheetsSubject = new BehaviorSubject<SheetTabs>(this.tabsNow());
    this.namesSubject = new BehaviorSubject<SheetNames>({ entries: [], refused: '' });
    // Seeded from the document, because a sheet loaded from a file
    // arrives with its names already in it and nothing else would
    // ever tell the other thread they exist.
    this.publishNames('');
    this.statusSubject = new BehaviorSubject<SheetStatus>(this.statusNow());

    this.window = this.windowSubject;
    this.sheets = this.sheetsSubject;
    this.validation = this.validationSubject;
    this.geometry = this.geometrySubject;
    this.selection = this.selectionSubject;
    this.editor = this.editorSubject;
    this.names = this.namesSubject;
    this.status = this.statusSubject;
    this.clipboard = this.clipboardSubject;
    this.selectionStats = this.statsSubject;
    this.findView = this.findSubject;
    this.formats = this.formatsSubject;
    this.palette = this.paletteSubject;
    this.activeFormat = this.activeFormatSubject;
    this.autofit = this.autofitSubject;
    this.publishStats();
    this.publishActiveFormat();
  }

  // ---------------------------------------------------------------------
  // Commands
  // ---------------------------------------------------------------------

  /**
   * The render worker's window moved.
   *
   * Published at once, whatever else is happening. This is the line
   * that keeps a scroll answerable during a recalc: the cells for the
   * new range are read straight out of the store, which holds correct
   * values for everything the recalc has already reached and stale
   * ones for what it has not — and a stale value on screen for two
   * frames is not a blank sheet for two seconds.
   */
  setViewport(sheet: number, firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void {
    this.viewport = { firstRow, lastRow, firstColumn, lastColumn };
    // The sheet arrives with the window, so a scroll and a tab change
    // are the same message. Published in full when it moved, because
    // everything below the window is the active sheet's too.
    if (this.document.activate(sheet)) {
      this.publishSheet();
      return;
    }
    this.publishWindow();
    this.publishFormats();
    this.publishValidation();
  }

  // ---------------------------------------------------------------------
  // The sheets
  // ---------------------------------------------------------------------

  activateSheet(sheet: number): void {
    if (this.document.activate(sheet)) {
      this.publishSheet();
    }
  }

  addSheet(): void {
    this.document.addSheet();
    // A new sheet's columns are as wide as the first sheet's were
    // when it was made. The constructor does this for sheet one and
    // nothing did it for the rest, so an added sheet reached the file
    // with no widths at all — which drew correctly, because the grid
    // falls back, and reloaded as a sheet whose widths were a shorter
    // array than the sheet is wide.
    this.document.columnWidths = this.defaultWidths();
    this.publishSheet();
    this.persist();
  }

  private defaultWidths(): number[] {
    return Array.from({ length: this.geometrySubject.value.columnCount }, () => COLUMN_WIDTH);
  }

  renameSheet(sheet: number, name: string): void {
    if (this.document.renameSheet(sheet, name)) {
      this.publishSheet();
      this.persist();
      this.pump();
    }
  }

  removeSheet(sheet: number): void {
    if (this.document.removeSheet(sheet)) {
      this.publishSheet();
      this.persist();
      this.pump();
    }
  }

  moveSheet(from: number, to: number): void {
    if (this.document.moveSheet(from, to)) {
      this.publishSheet();
      this.persist();
    }
  }

  duplicateSheet(sheet: number): void {
    if (this.document.duplicateSheet(sheet) !== -1) {
      this.publishSheet();
      this.persist();
      this.pump();
    }
  }

  setSheetColour(sheet: number, colour: string | null): void {
    if (this.document.setSheetColour(sheet, colour)) {
      this.publishTabs();
      this.persist();
    }
  }

  // ---------------------------------------------------------------------
  // Formats that think, and what a cell is allowed to hold
  // ---------------------------------------------------------------------

  addConditional(rule: SheetConditionalRule): void {
    const rect = rectOf(this.document.selection);
    this.document.addConditional({
      range: {
        start: relativeRef(rect.firstRow, rect.firstColumn),
        end: relativeRef(rect.lastRow, rect.lastColumn)
      },
      test: rule.test,
      paint: rule.paint,
      scale: rule.scale
    });
    this.rulesChanged();
  }

  removeConditional(at: number): void {
    if (this.document.removeConditional(at)) {
      this.rulesChanged();
    }
  }

  addValidation(rule: SheetValidationRule, strict: boolean, message: string): void {
    const rect = rectOf(this.document.selection);
    this.document.addValidation({
      range: {
        start: relativeRef(rect.firstRow, rect.firstColumn),
        end: relativeRef(rect.lastRow, rect.lastColumn)
      },
      rule,
      strict,
      message: message === '' ? undefined : message
    });
    this.rulesChanged();
  }

  removeValidation(at: number): void {
    if (this.document.removeValidation(at)) {
      this.rulesChanged();
    }
  }

  /**
   * Every rule off the sheet, as one step.
   *
   * The only bulk command here, and it exists because the bar has no
   * list of rules to pick from: rules are a fact about ranges that
   * overlap, and a list of them is a screen of its own. Clearing and
   * starting again is the honest small version, and it is one press
   * of ctrl-Z away from being taken back.
   */
  clearRules(): void {
    this.document.clearRules();
    this.rulesChanged();
  }

  /**
   * The rules changed, so everything they decide has to be asked
   * again.
   *
   * The palette is republished because a rule that has gone leaves
   * entries nothing points at — harmless, since the palette only
   * grows — but a rule that has *arrived* needs its colours sent
   * before the indices naming them are.
   */
  private rulesChanged(): void {
    this.painter.setRules(this.document.conditional);
    this.publishFormats();
    this.publishValidation();
    this.publishStatus();
    this.persist();
  }

  /**
   * The cells in view that break a rule, and what the active cell may
   * hold.
   *
   * Only the window, which is the whole shape of this phase: a rule
   * over a million cells is asked about the ones somebody can see.
   */
  private publishValidation(): void {
    const { firstRow, lastRow } = this.viewport;
    const rules = this.document.validations;
    const at = this.document.selection;
    const held = this.document.validationAt(at.row, at.column);
    const list = held !== null && held.rule.kind === 'list' ? held.rule.values : [];
    if (rules.length === 0) {
      // Nothing to say and nothing to walk. A sheet with no
      // validations costs this exactly one comparison per publish.
      const current = this.validationSubject.value;
      if (current.refused !== this.refusal || Object.keys(current.cells).length > 0 || current.list.length > 0) {
        this.validationSubject.next({ firstRow, lastRow, cells: {}, refused: this.refusal, list: [] });
      }
      return;
    }
    const columns = this.columnsInView();
    const cells: Record<string, Record<string, string>> = {};
    for (const row of this.rowsInView()) {
      const line: Record<string, string> = {};
      for (const column of columns) {
        const rule = this.document.validationAt(row, column);
        if (rule === null) {
          continue;
        }
        const complaint = validate(rule.rule, this.document.sheet.value(row, column));
        if (complaint !== null) {
          line[column] = rule.message ?? complaint;
        }
      }
      if (Object.keys(line).length > 0) {
        cells[row] = line;
      }
    }
    this.validationSubject.next({ firstRow, lastRow, cells, refused: this.refusal, list: [...list] });
  }

  private tabsNow(): SheetTabs {
    return { entries: this.document.sheets(), active: this.document.active };
  }

  private publishTabs(): void {
    this.sheetsSubject.next(this.tabsNow());
  }

  /**
   * Everything that is about *a* sheet, because the sheet changed.
   *
   * Every key here is the active sheet's, which is the other half of
   * the claim this phase defends: the window, the formats and the
   * geometry describe one sheet, so a sheet nobody is looking at
   * cannot publish anything — there is no key for it to publish on.
   */
  private publishSheet(): void {
    // Each sheet has its own rules, so the painter is bound to the
    // one in view rather than to the document.
    this.painter.setRules(this.document.conditional);
    this.publishTabs();
    this.publishWindow();
    this.publishFormats();
    this.publishPalette();
    this.publishGeometry();
    this.selectionSubject.next(this.document.selection);
    this.publishEditor();
    this.publishActiveFormat();
    this.publishValidation();
    this.publishStats();
    this.publishStatus();
  }

  setCell(row: number, column: number, input: string): void {
    // A changed cell can move what a colour scale spreads between.
    this.painter.invalidate();
    const refused = this.document.setCell(row, column, input);
    this.refusal = refused ?? '';
    if (refused !== null) {
      // A rule refused the write, so nothing changed but what has to
      // be said about it.
      this.publishValidation();
      return;
    }
    // The edited cell's own value is settled already — a literal is
    // itself and a formula is queued — so the window can go out before
    // any arithmetic, which is what makes typing feel immediate.
    this.publishWindow();
    this.repaintIfRuled();
    this.publishValidation();
    this.publishEditor();
    this.publishStatus();
    this.persist();
    this.pump();
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void {
    this.document.setSelection(row, column, anchorRow, anchorColumn);
    this.selectionSubject.next(this.document.selection);
    this.publishEditor();
    this.publishStats();
    this.publishActiveFormat();
    // Which match the selection is on, not which matches there are.
    // Moving off a match with the find bar open has to stop saying
    // "3 of 412", and the only thing that changed is where we are.
    this.publishFindPosition();
  }

  undo(): void {
    if (this.document.undo()) {
      this.afterHistory();
    }
  }

  redo(): void {
    if (this.document.redo()) {
      this.afterHistory();
    }
  }

  copy(cut: boolean): void {
    const rect = rectOf(this.document.selection);
    const text = copyRect(this.document, rect);
    this.copied = { text, row: rect.firstRow, column: rect.firstColumn };
    this.serial++;
    this.clipboardSubject.next({ text, serial: this.serial });
    if (cut) {
      clearRect(this.document, rect);
      this.afterEdit();
    }
  }

  paste(text: string): void {
    const at = rectOf(this.document.selection);
    const written = pasteBlock(this.document, text, { row: at.firstRow, column: at.firstColumn }, this.copied);
    this.document.setSelection(written.firstRow, written.firstColumn, written.lastRow, written.lastColumn);
    this.selectionSubject.next(this.document.selection);
    this.afterEdit();
  }

  clearRange(): void {
    clearRect(this.document, rectOf(this.document.selection));
    this.afterEdit();
  }

  fill(toRow: number, toColumn: number): void {
    const source = rectOf(this.document.selection);
    const target = fillTarget(source, toRow, toColumn);
    fillRect(this.document, source, target);
    this.document.setSelection(target.firstRow, target.firstColumn, target.lastRow, target.lastColumn);
    this.selectionSubject.next(this.document.selection);
    this.afterEdit();
  }

  setColumnWidth(column: number, width: number): void {
    const geometry = this.geometrySubject.value;
    if (column < 0 || column >= geometry.columnCount) {
      return;
    }
    const columnWidths = [...this.document.columnWidths];
    columnWidths[column] = Math.max(MIN_COLUMN_WIDTH, Math.round(width));
    this.document.columnWidths = columnWidths;
    this.publishGeometry();
    this.persist();
  }

  /** The geometry, with the widths and hidden rows the document holds. */
  private publishGeometry(): void {
    this.geometrySubject.next({
      ...this.geometrySubject.value,
      columnWidths: this.document.columnWidths,
      hiddenRows: [...new Set([...this.document.hiddenRows, ...this.document.filteredRows])].sort((a, b) => a - b),
      frozenRows: this.document.frozenRows,
      frozenColumns: this.document.frozenColumns,
      merges: this.document.merges.all.map(rect => ({ ...rect }))
    });
  }

  /**
   * A chain of `cells` cells, each reading the one before it.
   *
   * Laid out across the sheet in rows so that a chain longer than the
   * sheet is tall still fits, and written through the model rather
   * than the document so that two hundred thousand cells are not two
   * hundred thousand undo entries. Built once; every call after moves
   * the head, which is what makes the whole chain out of date.
   *
   * It starts one row past the end of the sheet, so every cell in it
   * is somewhere nobody can scroll to, select, or type in. Two reasons,
   * and the second was found by reloading the page rather than by
   * reading the code: the claim is that recalculating cells nobody can
   * see costs the scroll nothing, so the cells have to be ones nobody
   * can see — and a cell outside the sheet is not part of the document,
   * which is what keeps a quarter of a million formulas from being
   * saved to somebody's file the first time they press the button. See
   * `snapshotOf`.
   */
  stress(cells: number): void {
    const sheet = this.document.sheet;
    const { columnCount, rowCount } = this.geometrySubject.value;
    if (this.stressCells !== cells) {
      const first = rowCount;
      sheet.setCell(first, 0, '1');
      // `cells` formulas, plus the literal head they hang from, so the
      // number on the button is the number of cells that go out of
      // date rather than the number of cells written.
      for (let at = 1; at <= cells; at++) {
        const row = first + Math.floor(at / columnCount);
        const column = at % columnCount;
        const fromRow = first + Math.floor((at - 1) / columnCount);
        const fromColumn = (at - 1) % columnCount;
        sheet.setCell(row, column, `=${columnName(fromColumn)}${fromRow + 1}+1`);
      }
      this.stressCells = cells;
    }
    this.stressRuns++;
    sheet.setCell(rowCount, 0, String(this.stressRuns));
    this.publishWindow();
    this.publishStatus();
    this.pump();
  }


  // ---------------------------------------------------------------------
  // Phase 8: fill, find and replace
  // ---------------------------------------------------------------------

  /**
   * Ctrl+D, and Ctrl+R for the other axis.
   *
   * A selection more than one cell tall repeats its own top row down
   * itself. A selection one cell tall takes from the cell *above*
   * instead, which is what every spreadsheet does and what makes the
   * key usable without selecting a range first — the alternative is a
   * key that silently does nothing nine times out of ten.
   */
  fillDown(): void {
    const rect = rectOf(this.document.selection);
    if (rect.lastRow > rect.firstRow) {
      this.fillWithin({ ...rect, lastRow: rect.firstRow }, rect);
      return;
    }
    if (rect.firstRow === 0) {
      return;
    }
    this.fillWithin({ ...rect, firstRow: rect.firstRow - 1, lastRow: rect.firstRow - 1 }, rect);
  }

  fillRight(): void {
    const rect = rectOf(this.document.selection);
    if (rect.lastColumn > rect.firstColumn) {
      this.fillWithin({ ...rect, lastColumn: rect.firstColumn }, rect);
      return;
    }
    if (rect.firstColumn === 0) {
      return;
    }
    this.fillWithin({ ...rect, firstColumn: rect.firstColumn - 1, lastColumn: rect.firstColumn - 1 }, rect);
  }

  /**
   * The same machinery the fill handle uses, so the two cannot
   * disagree about what a relative reference does when it moves.
   */
  private fillWithin(source: Rect, target: Rect): void {
    fillRect(this.document, source, fillTarget(source, target.lastRow, target.lastColumn));
    this.afterEdit();
  }

  find(query: string, matchCase: boolean, wholeCell: boolean, inFormulas: boolean): void {
    const options: FindOptions = { matchCase, wholeCell, inFormulas };
    this.search(query, options);
    // Offer the cell the selection is already on. Somebody who
    // selected a cell and then searched for what is in it should not
    // be thrown to the next one.
    this.goToMatch(stepTo(this.found, this.currentKey(), false));
  }

  findStep(forward: boolean): void {
    const from = this.currentKey();
    this.goToMatch(forward ? stepTo(this.found, from, true) : stepBack(this.found, from));
  }

  /**
   * Replaces the match the selection is on and moves to the next.
   *
   * "The match it is on" and not "the first match": a person watching
   * the highlight move expects Replace to act on what they can see.
   * When the selection is not on a match this steps to one and
   * replaces nothing, which is what the button does everywhere.
   */
  replaceOne(replacement: string): void {
    const view = this.findSubject.value;
    const key = this.currentKey();
    if (!this.found.includes(key)) {
      this.findStep(true);
      return;
    }
    const where = at(key);
    const options = optionsOf(view);
    const before = this.document.sheet.input(where.row, where.column);
    this.document.setCell(where.row, where.column, replaceIn(before, view.query, replacement, options));
    this.afterEdit();
    this.search(view.query, options);
    this.goToMatch(stepTo(this.found, key, true));
  }

  /**
   * Every match, as one step on the undo stack.
   *
   * One step and not one per cell, for the reason `transact` exists:
   * replacing four hundred cells and then needing four hundred
   * presses of ctrl-Z to take it back is how people stop trusting
   * undo.
   */
  replaceAll(replacement: string): void {
    const view = this.findSubject.value;
    if (view.query === '' || this.found.length === 0) {
      return;
    }
    const options = optionsOf(view);
    const targets = [...this.found];
    this.document.transact(() => {
      for (const key of targets) {
        const where = at(key);
        const before = this.document.sheet.input(where.row, where.column);
        this.document.setCell(where.row, where.column, replaceIn(before, view.query, replacement, options));
      }
    });
    this.afterEdit();
    this.search(view.query, options);
  }

  clearFind(): void {
    this.found = [];
    this.findSubject.next(NO_FIND);
  }

  /** Runs the search and publishes what it found. */
  private search(query: string, options: FindOptions): void {
    const { rowCount, columnCount } = this.geometrySubject.value;
    this.found = findMatches(this.document, query, options, rowCount, columnCount);
    this.findSubject.next({
      query,
      ...options,
      matches: this.found.length,
      active: this.activeMatch()
    });
  }

  private goToMatch(key: number): void {
    if (key === -1) {
      this.publishFindPosition();
      return;
    }
    const where = at(key);
    this.setSelection(where.row, where.column, where.row, where.column);
  }

  /** The cell the selection's active corner is on, as a key. */
  private currentKey(): number {
    const { row, column } = this.document.selection;
    return cellKey(row, column);
  }

  /** Which match the selection is on, from one, or zero for none. */
  private activeMatch(): number {
    return this.found.indexOf(this.currentKey()) + 1;
  }

  /**
   * The count is unchanged and only the position moved.
   *
   * Split out from `search` because moving the selection must not
   * re-walk the store: arrow keys move the selection, and a find bar
   * left open would otherwise make every arrow key a full search.
   */
  private publishFindPosition(): void {
    const view = this.findSubject.value;
    if (view.query === '') {
      return;
    }
    const active = this.activeMatch();
    if (active !== view.active) {
      this.findSubject.next({ ...view, active });
    }
  }

  // ---------------------------------------------------------------------
  // Phase 9: formatting
  // ---------------------------------------------------------------------

  /**
   * Applies a change to every cell in the selection.
   *
   * Cell by cell, and that is not laziness: a change is relative to
   * what each cell already has, so a selection holding one bold cell
   * and one plain one, told `{ italic: true }`, ends up bold-italic
   * and italic rather than both the same. Making a format out of the
   * change once and stamping it over the range is the bug every
   * naive formatting model has, and it is the reason `format` takes
   * a change rather than a format.
   *
   * One step on the undo stack however many cells it touched.
   */
  format(change: SheetFormatChange): void {
    this.applyToSelection(format => applyChange(format, change));
  }

  clearFormat(): void {
    this.applyToSelection(() => DEFAULT_FORMAT);
  }

  /**
   * Borders over the selection, cell by cell.
   *
   * Always cell by cell, even when the selection is a whole column,
   * because a border pattern is *about* where each cell sits in the
   * block: `outline` puts an edge on the outer rim and nothing in the
   * middle, so two cells in the same column get different answers and
   * a region format — which by definition cannot vary inside itself —
   * is the wrong shape for it.
   *
   * The cost is a cell entry per bordered cell, which is what the
   * person asked for: they can see the border, so they can see what
   * it cost. A whole-sheet outline is four edges, not a million.
   */
  setBorders(pattern: BorderPattern, width: number, color: string): void {
    const rect = rectOf(this.document.selection);
    const { rowCount, columnCount } = this.geometrySubject.value;
    const lastRow = Math.min(rect.lastRow, rowCount - 1);
    const lastColumn = Math.min(rect.lastColumn, columnCount - 1);
    const edge: SheetEdge = { width, color };
    const off: SheetEdge = { width: 0, color: '' };

    this.document.transact(() => {
      for (let row = rect.firstRow; row <= lastRow; row++) {
        for (let column = rect.firstColumn; column <= lastColumn; column++) {
          const top = row === rect.firstRow;
          const bottom = row === lastRow;
          const left = column === rect.firstColumn;
          const right = column === lastColumn;
          const borders =
            pattern === 'none'
              ? { top: off, right: off, bottom: off, left: off }
              : pattern === 'all'
                ? { top: edge, right: edge, bottom: edge, left: edge }
                : pattern === 'outline'
                  ? {
                      top: top ? edge : off,
                      right: right ? edge : off,
                      bottom: bottom ? edge : off,
                      left: left ? edge : off
                    }
                  : pattern === 'top'
                    ? { top: top ? edge : off }
                    : { bottom: bottom ? edge : off };
          const held = this.document.formatAt(row, column);
          this.document.setFormat(row, column, applyChange(held, { borders }));
        }
      }
    });
    this.afterFormat();
  }

  /**
   * Runs a formatting change over the selection, as regions where it
   * can and cell by cell where it cannot.
   *
   * **The region case is not an optimisation.** Cell by cell, ctrl-A
   * followed by ctrl-B wrote a million cell entries, a million-entry
   * undo step, and a thirty-megabyte file that was then read back on
   * every load — found by pressing two keys in a browser, and by no
   * spec at all. A selection that covers the whole sheet, a whole
   * column or a whole row *is* a region, and storing it as one is the
   * same insight as the palette turned on its side.
   *
   * The change still has to be computed per region from what that
   * region already had, so that Bold does not undo Currency. What a
   * region cannot do is vary per cell inside it, which is exactly
   * what a region means.
   */
  private applyToSelection(change: (format: CellFormat) => CellFormat): void {
    const rect = rectOf(this.document.selection);
    const { rowCount, columnCount } = this.geometrySubject.value;
    const lastRow = Math.min(rect.lastRow, rowCount - 1);
    const lastColumn = Math.min(rect.lastColumn, columnCount - 1);
    const allRows = rect.firstRow === 0 && lastRow >= rowCount - 1;
    const allColumns = rect.firstColumn === 0 && lastColumn >= columnCount - 1;

    this.document.transact(() => {
      if (allRows && allColumns) {
        this.document.formatRegion('sheet', 0, change);
        return;
      }
      if (allRows) {
        for (let column = rect.firstColumn; column <= lastColumn; column++) {
          this.document.formatRegion('column', column, change);
        }
        return;
      }
      if (allColumns) {
        for (let row = rect.firstRow; row <= lastRow; row++) {
          this.document.formatRegion('row', row, change);
        }
        return;
      }
      for (let row = rect.firstRow; row <= lastRow; row++) {
        for (let column = rect.firstColumn; column <= lastColumn; column++) {
          this.document.setFormat(row, column, change(this.document.formatAt(row, column)));
        }
      }
    });
    this.afterFormat();
  }

  /**
   * What a format change has to republish.
   *
   * The window as well as the indices, because a number format
   * changes the *string* a cell shows — that is what a number format
   * is — and the string is what the window carries. Bounded by the
   * viewport, which is the whole answer to whether that is
   * affordable.
   */
  private afterFormat(): void {
    this.publishPalette();
    this.publishFormats();
    this.publishWindow();
    this.publishStatus();
    this.publishActiveFormat();
    this.persist();
    this.pump();
  }

  // ---------------------------------------------------------------------
  // Phase 10: rows and columns
  // ---------------------------------------------------------------------

  /**
   * Sorts the selection, or the block around it.
   *
   * `widen` is what a selection of one cell means: sort the table I
   * am standing in. Where that table *stops* is this side's question
   * — the render worker holds the rows it has mounted and the block
   * may be bigger or smaller — so the widening is here, and so is the
   * guess about whether the first row is a heading.
   */
  sortRange(column: number, ascending: boolean, widen: boolean): void {
    const { rowCount, columnCount } = this.geometrySubject.value;
    const at = this.document.selection;
    const selected = rectOf(at);
    const rect = widen
      ? currentRegion(this.document, at.row, at.column, rowCount, columnCount)
      : { ...selected, lastRow: Math.min(selected.lastRow, rowCount - 1) };

    sortRect(this.document, rect, {
      column,
      ascending,
      hasHeader: widen && looksLikeHeader(this.document, rect)
    });
    // The block that was sorted is what is now selected, so it is
    // plain what moved — and so a second sort does not have to guess
    // again.
    this.document.setSelection(rect.firstRow, rect.firstColumn, rect.lastRow, rect.lastColumn);
    this.selectionSubject.next(this.document.selection);
    this.afterEdit();
  }

  /**
   * A hidden column is one of width zero.
   *
   * That is the whole implementation, and it works because the widths
   * are already an array and the offsets already a prefix sum — Phase
   * 3 paid for both. The width it had is remembered so that showing
   * it again does not make it the default width instead of the one
   * somebody dragged.
   */
  hideColumns(first: number, last: number): void {
    const widths = [...this.document.columnWidths];
    for (let column = first; column <= last && column < widths.length; column++) {
      if (widths[column] > 0) {
        this.hiddenWidths.set(column, widths[column]);
        widths[column] = 0;
      }
    }
    this.document.columnWidths = widths;
    this.publishGeometry();
    this.persist();
  }

  /**
   * A hidden row, which the engine draws as one of height zero.
   *
   * It could not be done at all until the virtual sheet took a row
   * height per row: `columnWidth` was a number *or an array* and
   * `rowHeight` was only ever a number, so a column could be hidden
   * and a row could not. The heights are sparse rather than an array,
   * because a sheet is ten thousand rows tall and all but a handful
   * of them are the same.
   */
  hideRows(first: number, last: number): void {
    const { rowCount } = this.geometrySubject.value;
    for (let row = first; row <= last && row < rowCount; row++) {
      this.document.hiddenRows.add(row);
    }
    this.publishGeometry();
    this.persist();
  }

  showRows(first: number, last: number): void {
    // Widened by one on each side, so that selecting the rows either
    // side of a hidden one and asking to show it works — which is the
    // only way to select a row you cannot see.
    for (let row = Math.max(0, first - 1); row <= last + 1; row++) {
      this.document.hiddenRows.delete(row);
    }
    this.publishGeometry();
    this.persist();
  }

  /**
   * Freezes a pane, or unfreezes one with zeroes.
   *
   * Clamped to leave something scrolling: a sheet frozen all the way
   * down is a sheet that cannot be scrolled, and the person who did
   * it by accident has no way back except the menu they have just
   * learned not to trust.
   */
  freeze(rows: number, columns: number): void {
    const { rowCount, columnCount } = this.geometrySubject.value;
    this.document.frozenRows = Math.min(Math.max(0, Math.floor(rows)), Math.max(0, rowCount - 1));
    this.document.frozenColumns = Math.min(Math.max(0, Math.floor(columns)), Math.max(0, columnCount - 1));
    this.publishGeometry();
    // The pane is part of what is on screen, so the cells in it have
    // to be sent — a frozen column outside the scrolled window is a
    // column of blanks otherwise.
    this.publishWindow();
    this.publishFormats();
    this.persist();
  }

  /**
   * Merges the selection, emptying everything but its top-left cell.
   *
   * Destructive on purpose and in one step of undo, which is the
   * promise a warning dialog makes and this keeps: the cells a merge
   * covers have nowhere to show what they held, so they are cleared
   * — and ctrl-Z puts every one of them back.
   */
  mergeCells(): void {
    const rect = rectOf(this.document.selection);
    const { rowCount, columnCount } = this.geometrySubject.value;
    const merged = {
      firstRow: rect.firstRow,
      lastRow: Math.min(rect.lastRow, rowCount - 1),
      firstColumn: rect.firstColumn,
      lastColumn: Math.min(rect.lastColumn, columnCount - 1)
    };
    if (merged.lastRow === merged.firstRow && merged.lastColumn === merged.firstColumn) {
      return;
    }
    this.document.transact(() => {
      for (let row = merged.firstRow; row <= merged.lastRow; row++) {
        for (let column = merged.firstColumn; column <= merged.lastColumn; column++) {
          if (row !== merged.firstRow || column !== merged.firstColumn) {
            this.document.setCell(row, column, '');
          }
        }
      }
    });
    this.document.merges.add(merged);
    this.publishGeometry();
    this.afterEdit();
  }

  /**
   * Gives the selection a name, or publishes why it cannot have one.
   *
   * The refusal travels with the list rather than coming back as a
   * return value, because the render worker asked over a channel and
   * a channel command answers by the view changing. A name box that
   * argued in a dialog would be worse than one that shows a line of
   * text under itself.
   */
  defineName(name: string): void {
    const rect = rectOf(this.document.selection);
    const range = {
      start: relativeRef(rect.firstRow, rect.firstColumn),
      end: relativeRef(rect.lastRow, rect.lastColumn)
    };
    const problem = this.document.defineName(name, range);
    if (problem !== null) {
      this.publishNames(nameProblemText(problem));
      return;
    }
    this.publishNames('');
    this.afterEdit();
  }

  removeName(name: string): void {
    if (this.document.removeName(name)) {
      this.publishNames('');
      this.afterEdit();
    }
  }

  private publishNames(refused: string): void {
    this.namesSubject.next({
      entries: this.document.sheet.names.all().map(entry => ({
        name: entry.name,
        firstRow: Math.min(entry.range.start.row, entry.range.end.row),
        firstColumn: Math.min(entry.range.start.column, entry.range.end.column),
        lastRow: Math.max(entry.range.start.row, entry.range.end.row),
        lastColumn: Math.max(entry.range.start.column, entry.range.end.column)
      })),
      refused
    });
  }

  /**
   * The longest strings in each column, for the thread that can
   * measure them.
   *
   * Character count picks the shortlist and not the winner: in a
   * proportional font `WWW` is wider than `lllllll`, so the longest
   * string is often not the widest one. Sending several and letting
   * the render worker measure all of them is what makes the answer
   * right without this side ever learning what a font is.
   *
   * Bounded by the store, not the sheet: a column of eight values in
   * a ten-thousand-row sheet costs eight.
   */
  measureColumns(first: number, last: number): void {
    const { rowCount } = this.geometrySubject.value;
    const wanted = new Set<number>();
    for (let column = first; column <= last; column++) {
      wanted.add(column);
    }
    const best = new Map<number, { text: string; bold: boolean }[]>();
    for (const cell of this.document.sheet.entries()) {
      if (cell.row >= rowCount || !wanted.has(cell.column)) {
        continue;
      }
      const text = this.document.display(cell.row, cell.column);
      if (text === '') {
        continue;
      }
      const list = best.get(cell.column) ?? [];
      list.push({ text, bold: this.document.formatAt(cell.row, cell.column).paint.bold });
      best.set(cell.column, list);
    }

    this.autofitSerial++;
    this.autofitSubject.next({
      serial: this.autofitSerial,
      columns: [...wanted].sort((a, b) => a - b).map(column => {
        const list = (best.get(column) ?? [])
          .sort((a, b) => [...b.text].length - [...a.text].length)
          .slice(0, AUTOFIT_SAMPLES);
        return { column, samples: list.map(entry => entry.text), bold: list.map(entry => entry.bold) };
      })
    });
  }

  unmergeCells(): void {
    if (!this.document.merges.remove(rectOf(this.document.selection))) {
      return;
    }
    this.publishGeometry();
    this.publishWindow();
  }

  /**
   * Hides every row in the block whose cell in this column is not the
   * one the selection is on.
   *
   * Filtering by the value under the cursor, which is the filter
   * people actually use and the one that needs no dialog: stand on
   * `North` and ask for it, and the sheet is the North rows. The
   * block is the current region, on the same reasoning as a sort —
   * where the table stops is a question only this side can answer.
   *
   * It is a *snapshot*, not a rule. Editing a cell afterwards does
   * not re-run it, which is what every spreadsheet does and what
   * keeps an edit from making rows vanish under somebody's hands.
   */
  filterToSelection(): void {
    const { rowCount, columnCount } = this.geometrySubject.value;
    const at = this.document.selection;
    const rect = currentRegion(this.document, at.row, at.column, rowCount, columnCount);
    const wanted = this.document.display(at.row, at.column);
    const header = looksLikeHeader(this.document, rect) ? rect.firstRow : -1;

    this.document.filteredRows.clear();
    for (let row = rect.firstRow; row <= rect.lastRow; row++) {
      if (row === header || row === at.row) {
        continue;
      }
      if (this.document.display(row, at.column) !== wanted) {
        this.document.filteredRows.add(row);
      }
    }
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.persist();
  }

  clearFilter(): void {
    if (this.document.filteredRows.size === 0) {
      return;
    }
    this.document.filteredRows.clear();
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.persist();
  }

  showColumns(first: number, last: number): void {
    const widths = [...this.document.columnWidths];
    // Widened by one on each side, so that selecting the columns
    // either side of a hidden one and asking to show it works — which
    // is the only way to select a column you cannot see.
    for (let column = Math.max(0, first - 1); column <= last + 1 && column < widths.length; column++) {
      if (widths[column] === 0) {
        widths[column] = this.hiddenWidths.get(column) ?? COLUMN_WIDTH;
        this.hiddenWidths.delete(column);
      }
    }
    this.document.columnWidths = widths;
    this.publishGeometry();
    this.persist();
  }

  insertRows(at: number, count: number): void {
    this.structural({ axis: 'row', at, by: Math.max(1, count) });
  }

  deleteRows(at: number, count: number): void {
    this.structural({ axis: 'row', at, by: -Math.max(1, count) });
  }

  insertColumns(at: number, count: number): void {
    this.structural({ axis: 'column', at, by: Math.max(1, count) });
  }

  deleteColumns(at: number, count: number): void {
    this.structural({ axis: 'column', at, by: -Math.max(1, count) });
  }

  /**
   * A structural change, and everything that has to follow it.
   *
   * Nearly every published key moves: the window because cells moved,
   * the formats because they moved with them, the geometry because a
   * column insert moves the widths, and the editor because the cell
   * the formula bar is showing may now be a different one.
   *
   * It is *not* sliced. An insert rewrites the formulas that mention
   * the line and rebuilds the graph, which on a normal sheet is
   * instant and on the fifty-thousand-formula chain is not — and the
   * recalculation it causes goes through the pump as every other edit
   * does, so what is left unsliced is the rewrite itself. `pnpm proof`
   * is what says whether that is affordable; see the phase's notes.
   */
  private structural(shift: Shift): void {
    this.document.applyShift(shift);
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.publishPalette();
    this.publishEditor();
    this.publishStatus();
    this.publishStats();
    this.publishActiveFormat();
    // A search's matches are cell keys, and every one of them past
    // the line is now the wrong cell.
    if (this.findSubject.value.query !== '') {
      this.search(this.findSubject.value.query, optionsOf(this.findSubject.value));
    }
    this.persist();
    this.pump();
  }

  // ---------------------------------------------------------------------
  // Keeping it
  // ---------------------------------------------------------------------

  /**
   * Loads what was stored, or seeds a sheet that has never been opened.
   *
   * Async, and called after `serveChannels` rather than before it: a
   * channel served late misses the handshake, and the render worker is
   * perfectly able to draw an empty grid for the frame it takes to
   * read a file. The seed is written straight back, so what is on disk
   * from the second run onwards is a file this build wrote.
   */
  async restore(seed?: (document: SheetDocument) => void): Promise<void> {
    const stored = (await this.repository?.load()) ?? null;
    if (stored === null) {
      seed?.(this.document);
      this.document.sheet.recalculate();
    } else {
      applySnapshot(this.document, stored);
      this.publishSheet();
    }
    this.restored = true;
    this.selectionSubject.next(this.document.selection);
    this.publishWindow();
    this.publishFormats();
    // The palette, which nothing else publishes on this path. Without
    // it a loaded sheet's cells all point at entries the render
    // worker has never been sent, so every one of them falls back to
    // plain — the numbers come out formatted, because that happens on
    // this thread, and not one cell is bold. Which is exactly how it
    // looked in a browser.
    this.publishPalette();
    this.publishEditor();
    this.publishStatus();
    this.publishStats();
    this.publishActiveFormat();
    if (stored === null) {
      this.persist();
    }
  }

  /** The snapshot as it stands, for a spec or a worker shutting down. */
  snapshot(): SheetSnapshot {
    return snapshotOf(this.document, this.geometrySubject.value.rowCount);
  }

  /** Writes anything outstanding now. */
  flush(): Promise<void> {
    return this.repository?.flush() ?? Promise.resolve();
  }

  private persist(): void {
    if (!this.restored) {
      return;
    }
    this.repository?.save(this.snapshot());
  }

  /** What every edit that is not a single keystroke has to do afterwards. */
  /**
   * The formats again, but only when a rule could have changed them.
   *
   * Before this phase a cell's format could not change because its
   * *value* did, so typing never republished them. A conditional
   * format is exactly that, so it has to — and a sheet with no rules
   * goes on paying nothing, which is what keeps `pnpm proof`
   * measuring the same thing it always did.
   */
  private repaintIfRuled(): void {
    if (!this.painter.isEmpty) {
      this.publishFormats();
    }
  }

  private afterEdit(): void {
    // The rules read the cells, so a changed cell can change what a
    // scale spreads between.
    this.painter.invalidate();
    this.publishWindow();
    this.publishFormats();
    this.publishValidation();
    this.publishEditor();
    this.publishStatus();
    this.publishStats();
    this.persist();
    this.pump();
  }

  private afterHistory(): void {
    this.selectionSubject.next(this.document.selection);
    // An undone column insert puts the widths back where they were,
    // and the geometry is where the render worker reads them.
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.publishPalette();
    this.publishEditor();
    this.publishStatus();
    this.publishStats();
    // An undo is an edit as far as the file is concerned. Left out,
    // taking something back and closing the tab would bring it back on
    // the next open, which is the opposite of what undo promises.
    this.persist();
    this.pump();
  }

  // ---------------------------------------------------------------------
  // The pump
  // ---------------------------------------------------------------------

  /** True while a recalc is still in slices. */
  get recalculating(): boolean {
    return this.pumping;
  }

  private pump(): void {
    if (this.pumping || this.document.sheet.pending === 0) {
      return;
    }
    this.pumping = true;
    this.step();
  }

  private step(): void {
    const result = this.document.sheet.recalculate(this.budget);
    this.stats.slices++;
    // A formula settling is a value changing, which a scale's extent
    // is computed from — so a recalculation invalidates it the same
    // way an edit does.
    if (result.evaluated > 0) {
      this.painter.invalidate();
    }
    this.publishWindow();
    this.repaintIfRuled();
    this.publishStatus();
    this.publishStats();
    if (result.done) {
      this.pumping = false;
      return;
    }
    this.schedule(() => this.step());
  }

  // ---------------------------------------------------------------------
  // Publishing
  // ---------------------------------------------------------------------

  /**
   * The window, rebuilt from the store.
   *
   * Rebuilt whole rather than patched, because the differ is what
   * decides the wire: two structurally equal windows produce no
   * patches at all and `provide` sends nothing, so publishing after
   * every slice costs a walk of the visible cells and nothing else.
   * The walk is a few hundred string comparisons; the alternative is
   * bookkeeping that has to be right about which cells a recalc
   * touched, which is the same information the differ already has.
   */
  /**
   * The rows and columns somebody can see.
   *
   * The scrolled window *and the frozen pane*, which is not the same
   * rectangle: a sheet frozen at column A and scrolled to column E
   * shows A and E through N, and nothing between. The frozen cells
   * were drawn empty until this existed — correctly placed, correctly
   * stuck, and holding nothing, because the window they would have
   * come from had scrolled past them.
   *
   * Listed rather than bounded, so the gap in the middle costs
   * nothing. The frozen pane is a handful of rows and columns; asking
   * for everything from row 0 to the window instead would fetch five
   * thousand rows to show one.
   */
  private rowsInView(): number[] {
    const { firstRow, lastRow } = this.viewport;
    const rows: number[] = [];
    for (let row = 0; row < this.document.frozenRows && row < firstRow; row++) {
      rows.push(row);
    }
    for (let row = firstRow; row <= lastRow; row++) {
      rows.push(row);
    }
    return rows;
  }

  private columnsInView(): number[] {
    const { firstColumn, lastColumn } = this.viewport;
    const columns: number[] = [];
    for (let column = 0; column < this.document.frozenColumns && column < firstColumn; column++) {
      columns.push(column);
    }
    for (let column = firstColumn; column <= lastColumn; column++) {
      columns.push(column);
    }
    return columns;
  }

  private publishWindow(): void {
    const { firstRow, lastRow, firstColumn, lastColumn } = this.viewport;
    if (lastRow < firstRow || lastColumn < firstColumn) {
      this.windowSubject.next(EMPTY_WINDOW);
      return;
    }
    const columns = this.columnsInView();
    const cells: Record<string, Record<string, string>> = {};
    for (const row of this.rowsInView()) {
      const line: Record<string, string> = {};
      for (const column of columns) {
        line[column] = this.document.display(row, column);
      }
      cells[row] = line;
    }
    this.stats.publishes++;
    this.windowSubject.next({ firstRow, lastRow, firstColumn, lastColumn, cells });
  }

  /**
   * The palette index for each visible cell.
   *
   * A cell with the default format is left out rather than sent as a
   * zero, so an unformatted sheet publishes an object of empty
   * objects — and, once the differ has seen it, nothing at all on
   * every publish after. Scrolling a sheet nobody has formatted
   * costs four patches here and no more, the same as the window.
   */
  private publishFormats(): void {
    const { firstRow, lastRow, firstColumn, lastColumn } = this.viewport;
    if (lastRow < firstRow || lastColumn < firstColumn) {
      this.formatsSubject.next(EMPTY_FORMATS);
      return;
    }
    const columns = this.columnsInView();
    const cells: Record<string, Record<string, number>> = {};
    const grew = this.extraPaints.length;
    for (const row of this.rowsInView()) {
      const line: Record<string, number> = {};
      for (const column of columns) {
        const id = this.paintedId(row, column);
        if (id !== 0) {
          line[column] = id;
        }
      }
      cells[row] = line;
    }
    this.formatsSubject.next({ firstRow, lastRow, firstColumn, lastColumn, cells });
    // A rule that asked for a colour nothing has used yet has just
    // put it in the palette, and an index into a palette the other
    // side has not been sent is an index it cannot draw.
    if (this.extraPaints.length !== grew || this.paletteBase() !== this.publishedBase) {
      this.publishPalette();
    }
  }

  /**
   * A cell's palette index, with its conditional formats folded in.
   *
   * The index is the document's own unless a rule paints over it, and
   * then it is an entry in a second palette that is appended to the
   * first. **Interned and never reordered**, so the same rule
   * produces the same index on every publish and a scroll costs the
   * rows that moved rather than the whole window.
   *
   * A sheet with no rules never reaches past the first line, which is
   * why this costs `pnpm proof` nothing.
   */
  private paintedId(row: number, column: number): number {
    const base = this.document.formats.idAt(row, column);
    if (this.painter.isEmpty) {
      return base;
    }
    const over = this.painter.paintFor(row, column, this.document.sheet.value(row, column));
    if (over === null) {
      return base;
    }
    const painted: CellPaint = {
      ...this.document.formats.byId(base).paint,
      ...(over.fill === undefined ? {} : { fill: over.fill }),
      ...(over.color === undefined ? {} : { color: over.color }),
      ...(over.bold === undefined ? {} : { bold: over.bold }),
      ...(over.italic === undefined ? {} : { italic: over.italic })
    };
    // A *position* in the extras, resolved against the document's
    // palette at publish time — so the document growing a format
    // moves every extra index and the palette goes out with it.
    return this.paletteBase() + this.internPaint(painted);
  }

  /**
   * A painted cell's index, appended to the palette the first time it
   * is seen.
   *
   * The table only grows, and what bounds it is the colour scale's
   * quantisation: a scale can ask for at most `SCALE_STEPS` colours
   * however many cells it covers, so scrolling a million-cell rule
   * reuses entries rather than making them.
   */
  /**
   * Where the conditional entries start in the published palette.
   *
   * The document's palette length, which is **not** `Formats.size` —
   * that counts the cells somebody has formatted. Using it put every
   * painted cell's index at zero on an unformatted sheet, which reads
   * as the default format and paints nothing at all.
   */
  private paletteBase(): number {
    return this.document.formats.entries.length;
  }

  private internPaint(paint: CellPaint): number {
    const key = JSON.stringify(paint);
    const held = this.extraIds.get(key);
    if (held !== undefined) {
      return held;
    }
    const at = this.extraPaints.length;
    this.extraIds.set(key, at);
    this.extraPaints.push(paint);
    return at;
  }

  /**
   * The paint half of every palette entry.
   *
   * Only ever appended to, so the differ sees one new element and
   * emits one patch however long the palette has grown. The number
   * format is left behind on this thread on purpose: what crosses is
   * the formatted string, and the render worker never learns a
   * locale.
   */
  private publishPalette(): void {
    this.publishedBase = this.paletteBase();
    this.paletteSubject.next({
      entries: [...this.document.formats.entries.map(format => format.paint), ...this.extraPaints]
    });
  }

  private publishActiveFormat(): void {
    const { row, column } = this.document.selection;
    const format = this.document.formatAt(row, column);
    this.activeFormatSubject.next({ paint: format.paint, number: format.number });
  }

  private publishEditor(): void {
    const { row, column } = this.document.selection;
    this.editorSubject.next({
      row,
      column,
      input: this.document.activeInput,
      explain: this.explainAt(row, column)
    });
  }

  /**
   * Why the active cell is broken, as a sentence and an address.
   *
   * Only for the one cell the selection is on: the walk is cheap for
   * one cell and would not be for a window of them, and the question
   * "why is *this* showing an error" is one somebody asks about the
   * cell they are pointing at.
   */
  private explainAt(row: number, column: number): SheetExplain | null {
    const found = explainCell(this.document.sheet, row, column);
    if (found === null) {
      return null;
    }
    return {
      code: found.code,
      meaning: found.meaning,
      blame: found.blame === null ? null : addressOf(found.blame.row, found.blame.column)
    };
  }

  private publishStatus(): void {
    this.statusSubject.next(this.statusNow());
  }

  /**
   * Sum, average and count over the selection.
   *
   * Recomputed after an edit as well as after a move, because a cell
   * that changed inside the selection changes the total — and the
   * cost is bounded by `aggregateOf` to whichever is smaller, the
   * selection or the store.
   */
  private publishStats(): void {
    const { rowCount } = this.geometrySubject.value;
    this.statsSubject.next(aggregateOf(this.document.sheet, rectOf(this.document.selection), rowCount));
  }

  private statusNow(): SheetStatus {
    return {
      pending: this.document.sheet.pending,
      evaluated: this.document.sheet.stats.evaluated,
      canUndo: this.document.canUndo,
      canRedo: this.document.canRedo
    };
  }
}

/** A published find view read back as the options that produced it. */
function optionsOf(view: SheetFindView): FindOptions {
  return { matchCase: view.matchCase, wholeCell: view.wholeCell, inFormulas: view.inFormulas };
}

/**
 * One cell's format with a change applied on top of it.
 *
 * Every absent field means "leave it alone", which is what makes a
 * toolbar of independent buttons possible: pressing Italic must not
 * undo what Bold did, and pressing Bold must not undo the currency
 * symbol somebody chose.
 */
function applyChange(format: CellFormat, change: SheetFormatChange): CellFormat {
  const number =
    change.number !== undefined
      ? (change.number as CellFormat['number'])
      : change.places !== undefined
        ? withPlaces(format.number, change.places)
        : format.number;
  return {
    number,
    paint: {
      bold: change.bold ?? format.paint.bold,
      italic: change.italic ?? format.paint.italic,
      underline: change.underline ?? format.paint.underline,
      fontSize: change.fontSize ?? format.paint.fontSize,
      color: change.color ?? format.paint.color,
      fill: change.fill ?? format.paint.fill,
      align: change.align ?? format.paint.align,
      wrap: change.wrap ?? format.paint.wrap,
      borders:
        change.borders === undefined
          ? format.paint.borders
          : {
              top: change.borders.top ?? format.paint.borders.top,
              right: change.borders.right ?? format.paint.borders.right,
              bottom: change.borders.bottom ?? format.paint.borders.bottom,
              left: change.borders.left ?? format.paint.borders.left
            }
    }
  };
}

/**
 * How many candidates a column sends.
 *
 * Enough that the widest is among them whichever way a proportional
 * font falls, few enough that measuring them is nothing. Twelve is
 * the number at which a column of mixed text stops changing its
 * answer when you add more.
 */
const AUTOFIT_SAMPLES = 12;
