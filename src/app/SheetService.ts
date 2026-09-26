import { BehaviorSubject, type Observable } from 'rxjs';

import { formatRange, relativeRef } from '../sheet/A1';
import {
  DEFAULT_CHART_HEIGHT,
  DEFAULT_CHART_WIDTH,
  MIN_CHART_HEIGHT,
  MIN_CHART_WIDTH,
  type Chart,
  type ChartKind
} from '../sheet/Chart';
import { headersIn, orientationOf, seriesFrom } from '../sheet/Series';
import { STRESS_CELLS } from './SheetCommands';
import type { CellValue } from '../sheet/Values';
import { addressOf, explainCell } from '../sheet/Explain';
import { nameProblemText } from '../sheet/Names';
import { aggregateOf } from './Aggregate';
import { ConditionalPainter } from './ConditionalPaint';
import type { CellPaint } from '../sheet/Format';
import { validate } from '../sheet/Validation';
import { ROW_HEIGHT, COLUMN_WIDTH, MIN_COLUMN_WIDTH, CELL_FONT_SIZE, MAX_ROW_HEIGHT, MIN_ROW_HEIGHT } from './dimensions';
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
  type SheetActiveRules,
  type SheetTransfer,
  type SheetDocumentView,
  type SheetDownload,
  type SheetAutofit,
  type SheetRowFit,
  type SheetRowFitRow,
  type SheetEdge,
  type SheetFormatChange,
  type SheetFormatWindow,
  type SheetPalette,
  type SheetStatus,
  type SheetTabs,
  type SheetValidation,
  type SheetChartSeries,
  type SheetCharts,
  type SheetConditionalRule,
  type SheetSeriesView,
  type SheetValidationRule,
  type SheetWindow
} from './SheetContract';
import { DEFAULT_FORMAT, withPlaces, type CellFormat } from '../sheet/Format';
import type { Shift } from '../sheet/Shift';
import { at, findMatches, replaceIn, stepBack, stepTo, type FindOptions } from './SheetFind';
import { sortRect } from './SheetSort';
import { NO_STATS, type SheetStats } from './Statistics';
import { SheetDocument } from './SheetDocument';
import { cellKey, columnName } from '../sheet/A1';
import { snapshotOf, applySnapshot, parseSnapshot, type SheetSnapshot } from './SheetFile';
import { FIRST_DOCUMENT, type DocumentEntry, type SheetLibrary } from './SheetLibrary';
import type { SheetRepository } from './SheetRepository';
import { exportCsv, importCsv } from './SheetCsv';
import { guessOf, placeOf } from './alignment';
import { platformInflate, reportOfXlsx, snapshotOfXlsx } from './SheetXlsx';
import { base64OfBytes, bytesOfBase64 } from './base64';
import { platformDeflate, xlsxOfDocument } from './SheetXlsxOut';
import { writeXlsx } from '../sheet/XlsxWrite';
import { openXlsx, XlsxError } from '../sheet/Xlsx';
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
  /**
   * Every workbook this browser keeps, for a service that opens them
   * by id; see `openDocument`. With a library the repository is the
   * open document's, and changes when another is opened.
   */
  readonly library?: SheetLibrary;
  /** What the very first document starts with, before anybody types. */
  readonly seed?: (document: SheetDocument) => void;
  /** The clock a document's `used` is read from. */
  readonly now?: () => number;
}

/** The view of a document before one has been opened. */
const NO_DOCUMENT: SheetDocumentView = { id: '', name: '', file: null, edited: false, elsewhere: false };

/** `Q3 sales.gsheet` is called `Q3 sales`. */
/** `a`, `a and b`, `a, b and c`. */
function listOf(items: readonly string[]): string {
  return items.length < 2 ? (items[0] ?? '') : `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`;
}

function baseName(fileName: string): string {
  const base = fileName.replace(/\.[^.]*$/, '').trim();
  return base === '' ? 'Untitled' : base;
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
  readonly transfer: Observable<SheetTransfer>;
  readonly documentView: Observable<SheetDocumentView>;
  readonly selectionStats: Observable<SheetStats>;
  readonly findView: Observable<SheetFindView>;
  readonly formats: Observable<SheetFormatWindow>;
  readonly palette: Observable<SheetPalette>;
  readonly activeFormat: Observable<SheetActiveFormat>;
  readonly activeRules: Observable<SheetActiveRules>;
  readonly autofit: Observable<SheetAutofit>;
  readonly rowFit: Observable<SheetRowFit>;
  readonly charts: Observable<SheetCharts>;
  readonly chartSeries: Observable<SheetSeriesView>;

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
  private readonly transferSubject = new BehaviorSubject<SheetTransfer>({ download: null, report: '' });
  private readonly documentSubject = new BehaviorSubject<SheetDocumentView>(NO_DOCUMENT);
  /** The open document's entry in the library, or null without one. */
  private entry: DocumentEntry | null = null;
  /** Whether anything has changed since the document was opened or saved to its file. */
  private edited = false;
  /** Lets go of the open document's claim; see `SheetLibrary.claim`. */
  private release: (() => void) | null = null;
  /** Whether another tab holds the open document, so this one must not keep it. */
  private elsewhere = false;
  /**
   * Opens, one after another. A document is loaded asynchronously, and
   * two opens that overlapped would each swap a document in under the
   * other.
   */
  private opening: Promise<void> = Promise.resolve();
  private readonly library: SheetLibrary | undefined;
  private readonly seed: ((document: SheetDocument) => void) | undefined;
  private readonly now: () => number;
  private downloadSerial = 0;
  private readonly statsSubject = new BehaviorSubject<SheetStats>(NO_STATS);
  private readonly findSubject = new BehaviorSubject<SheetFindView>(NO_FIND);
  private readonly formatsSubject = new BehaviorSubject<SheetFormatWindow>(EMPTY_FORMATS);
  private readonly paletteSubject = new BehaviorSubject<SheetPalette>({ entries: [PLAIN_PAINT] });
  private readonly activeFormatSubject = new BehaviorSubject<SheetActiveFormat>({
    paint: PLAIN_PAINT,
    number: { kind: 'general' }
  });
  private readonly activeRulesSubject = new BehaviorSubject<SheetActiveRules>({ conditional: null, validation: null });
  private readonly autofitSubject = new BehaviorSubject<SheetAutofit>({ serial: 0, columns: [] });
  private autofitSerial = 0;
  private readonly rowFitSubject = new BehaviorSubject<SheetRowFit>({ serial: 0, rows: [] });
  /**
   * The rows the last fit asked about, until it is answered. A newer
   * question replaces an older one, so it has to ask about both.
   */
  private rowFitPending: Set<number> | 'all' | null = null;
  private readonly chartsSubject = new BehaviorSubject<SheetCharts>({ entries: [], selected: 0 });
  private readonly seriesSubject = new BehaviorSubject<SheetSeriesView>({ charts: {} });
  /**
   * Which chart has the handles round it.
   *
   * On the service rather than in the document, because selecting a
   * chart is not an edit: it does not belong on the undo stack and it
   * does not belong in the file. The same argument the cell selection
   * loses — that one *is* in the file, because reopening a sheet on
   * the cell you left is worth a line of JSON — and a chart does not
   * win it, because there is no such thing as the chart you left.
   */
  private selectedChart = 0;
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
  /** Whether the last formats publish had a formatted cell in view; see `repaintIfRuled`. */
  private formattedInView = false;
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
  private repository: SheetRepository | undefined;
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
  /** How long the proof chart's series is, so it is built once. */
  private chartPoints = 0;
  private stressChart = 0;

  /**
   * The workbook in front of somebody. Not readonly, since Phase 16:
   * opening another document swaps a fresh one in rather than
   * rewriting this one, because a snapshot applied over a workbook
   * writes *over* what is there and leaves what it does not mention.
   */
  private document: SheetDocument;

  constructor(document: SheetDocument, options: SheetServiceOptions = {}) {
    this.document = document;
    this.library = options.library;
    this.seed = options.seed;
    this.now = options.now ?? Date.now;
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
      rowHeights: [],
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
    this.transfer = this.transferSubject;
    this.documentView = this.documentSubject;
    this.selectionStats = this.statsSubject;
    this.findView = this.findSubject;
    this.formats = this.formatsSubject;
    this.palette = this.paletteSubject;
    this.activeFormat = this.activeFormatSubject;
    this.activeRules = this.activeRulesSubject;
    this.autofit = this.autofitSubject;
    this.rowFit = this.rowFitSubject;
    this.charts = this.chartsSubject;
    this.chartSeries = this.seriesSubject;
    this.publishStats();
    this.publishActiveFormat();
    this.publishActiveRules();
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
    // A viewport over a sheet that is not the one showing is *stale*,
    // not a request to switch: the render worker names the sheet its
    // last tabs patch said was active, and that patch can be one
    // behind. Switching on it was a ping-pong — two sheets added in
    // one turn, and each side went on answering the other's previous
    // message, swapping the sheet back and forth for ever. Switching
    // is `activateSheet`'s, which the tabs and the menu send; this
    // answers for the sheet that is showing, and the render worker's
    // next tabs patch brings it into line.
    if (sheet !== this.document.active) {
      this.publishTabs();
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

  // ---------------------------------------------------------------------
  // Files in and out
  // ---------------------------------------------------------------------

  importCsv(fileName: string, text: string): void {
    if (!/\.(csv|tsv|txt)$/i.test(fileName)) {
      this.transferSubject.next({
        ...this.transferSubject.value,
        report: `${fileName} was not opened: this opens workbooks (.gsheet and .xlsx) and CSV files.`
      });
      return;
    }
    const { rowCount, columnCount } = this.geometrySubject.value;
    const opened = importCsv(this.document, fileName, text, { rows: rowCount, columns: columnCount });
    if (opened.sheet !== -1) {
      this.document.columnWidths = this.defaultWidths();
      this.document.setSelection(0, 0, 0, 0);
      this.publishSheet();
      this.persist();
      this.pump();
    }
    this.transferSubject.next({ ...this.transferSubject.value, report: opened.report });
  }

  /**
   * The whole workbook as an `.xlsx`, built here and handed to the
   * render worker as base64 for the shell to write.
   *
   * Asynchronous because the platform's deflate is, and queued like a
   * save so that it writes the workbook as it stood when it was asked.
   */
  exportXlsx(): void {
    const { rowCount } = this.geometrySubject.value;
    const { book, leftOut } = xlsxOfDocument(this.document, rowCount);
    this.xlsxLeftOut =
      leftOut.length === 0 ? '' : ` Its ${listOf(leftOut)} are not in it: an .xlsx from here carries cells, formats and layout.`;
    const name = `${this.entry?.name ?? 'Untitled'}.xlsx`;
    this.enqueue(async () => {
      const bytes = await writeXlsx(book, platformDeflate);
      this.publishDownload({
        kind: 'xlsx',
        name,
        mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        text: '',
        base64: base64OfBytes(bytes),
        handle: null
      });
    });
  }

  /** What the last `.xlsx` left out, said when it is saved. */
  private xlsxLeftOut = '';

  exportCsv(): void {
    this.publishDownload({
      kind: 'csv',
      name: `${this.document.sheet.name}.csv`,
      mediaType: 'text/csv',
      text: exportCsv(this.document),
      handle: null
    });
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
    this.publishActiveRules();
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
    // And its own charts, which the last sheet's selection does not
    // survive: an id from another page would draw handles round
    // nothing.
    this.selectedChart = 0;
    this.publishCharts();
    this.publishSeries();
    this.publishTabs();
    this.publishWindow();
    this.publishFormats();
    this.publishPalette();
    this.publishGeometry();
    this.selectionSubject.next(this.document.selection);
    this.publishEditor();
    this.publishActiveFormat();
    this.publishActiveRules();
    this.publishValidation();
    this.publishStats();
    this.publishStatus();
    this.fitRowsLater('all');
  }

  setCell(row: number, column: number, input: string): void {
    // What the last dropped file became is news until somebody starts
    // working, and then it is clutter.
    if (this.transferSubject.value.report !== '') {
      this.transferSubject.next({ ...this.transferSubject.value, report: '' });
    }
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
    this.redrawCharts();
    this.fitRowsLater([row]);
    this.persist();
    this.pump();
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void {
    this.document.setSelection(row, column, anchorRow, anchorColumn);
    this.selectionSubject.next(this.document.selection);
    this.publishEditor();
    this.publishStats();
    this.publishActiveFormat();
    this.publishActiveRules();
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
    // Wrapped text breaks at the column's width.
    this.fitRowsLater('all');
    this.publishGeometry();
    this.persist();
  }

  /** The geometry, with the widths and hidden rows the document holds. */
  private publishGeometry(): void {
    this.geometrySubject.next({
      ...this.geometrySubject.value,
      columnWidths: this.document.columnWidths,
      hiddenRows: [...new Set([...this.document.hiddenRows, ...this.document.filteredRows])].sort((a, b) => a - b),
      rowHeights: this.rowHeightsNow(),
      frozenRows: this.document.frozenRows,
      frozenColumns: this.document.frozenColumns,
      merges: this.document.merges.all.map(rect => ({ ...rect }))
    });
  }

  /** Every row that is not the default height: set by hand first, then fitted. */
  private rowHeightsNow(): [number, number][] {
    const heights = new Map(this.document.fittedRows);
    for (const [row, height] of this.document.rowHeights) {
      heights.set(row, height);
    }
    return [...heights].sort((a, b) => a[0] - b[0]);
  }

  setRowHeight(row: number, height: number): void {
    if (row < 0 || row >= this.geometrySubject.value.rowCount) {
      return;
    }
    if (height <= 0) {
      // Back to fitting what it holds, which only the render worker
      // can say, so it is asked.
      this.document.rowHeights.delete(row);
      this.fitRowsLater([row]);
    } else {
      this.document.rowHeights.set(row, Math.min(Math.max(Math.round(height), MIN_ROW_HEIGHT), MAX_ROW_HEIGHT));
    }
    this.publishGeometry();
    this.persist();
  }

  setIteration(on: boolean): void {
    if (on === (this.document.book.iteration !== null)) {
      return;
    }
    // Excel's defaults, which is also what a file that turns it on
    // without saying how far means.
    this.document.book.iteration = on ? { count: 100, delta: 0.001 } : null;
    // Every circle has to be worked out again, one way or the other;
    // re-reading every formula is what finds them.
    this.document.book.namesChanged();
    this.publishStatus();
    this.afterEdit();
  }

  fitRowsToContents(first: number, last: number): void {
    const cleared: number[] = [];
    for (const row of [...this.document.rowHeights.keys()]) {
      if (row >= first && row <= last) {
        this.document.rowHeights.delete(row);
        cleared.push(row);
      }
    }
    if (cleared.length === 0) {
      return;
    }
    this.fitRowsLater(cleared);
    this.publishGeometry();
    this.persist();
  }

  fitRows(serial: number, heights: readonly (readonly [number, number])[]): void {
    // An answer to a question since replaced — the sheet switched, or
    // something changed again — would fit rows to what they held then.
    if (serial !== this.rowFitSubject.value.serial) {
      return;
    }
    this.rowFitPending = null;
    const fitted = this.document.fittedRows;
    let changed = false;
    for (const [row, height] of heights) {
      const next = Math.min(Math.round(height), MAX_ROW_HEIGHT);
      if (next <= ROW_HEIGHT) {
        changed = fitted.delete(row) || changed;
      } else if (fitted.get(row) !== next) {
        fitted.set(row, next);
        changed = true;
      }
    }
    if (changed) {
      this.publishGeometry();
      this.persist();
    }
  }

  /**
   * Asks the render worker how tall some rows need to be.
   *
   * A row is a candidate when a cell in it wraps or is set larger than
   * the default, and has something in it. `'all'` looks at the whole
   * sheet, and at every row fitted before, so a row whose last wrapped
   * cell was emptied is sent with nothing and goes back to the default.
   * A sheet with no wrapping and no large fonts anywhere in its palette
   * and no fitted rows — nearly every sheet — costs one look at the
   * palette.
   */
  private fitRowsLater(asked: readonly number[] | 'all'): void {
    const pending = this.rowFitPending;
    const rows: readonly number[] | 'all' =
      asked === 'all' || pending === 'all' ? 'all' : pending === null ? asked : [...pending, ...asked];
    const palette = this.document.formats.entries;
    const tall = (format: CellFormat): boolean =>
      format.paint.wrap || (format.paint.fontSize !== 0 && format.paint.fontSize > CELL_FONT_SIZE);
    if (!palette.some(tall) && this.document.fittedRows.size === 0) {
      return;
    }
    const wanted = rows === 'all' ? null : new Set(rows);
    const { rowCount } = this.geometrySubject.value;
    const found = new Map<number, SheetRowFitRow['cells'][number][]>();
    if (wanted !== null) {
      for (const row of wanted) {
        found.set(row, []);
      }
    } else {
      for (const row of this.document.fittedRows.keys()) {
        found.set(row, []);
      }
    }
    for (const cell of this.document.sheet.entries()) {
      if (cell.row >= rowCount || (wanted !== null && !wanted.has(cell.row))) {
        continue;
      }
      const format = this.document.formatAt(cell.row, cell.column);
      // A merged cell is left to its own size, as Excel leaves it.
      if (!tall(format) || this.document.merges.at(cell.row, cell.column) !== null) {
        continue;
      }
      const text = this.document.display(cell.row, cell.column);
      if (text === '') {
        continue;
      }
      const list = found.get(cell.row) ?? [];
      list.push({
        text,
        width: this.document.columnWidths[cell.column] ?? COLUMN_WIDTH,
        fontSize: format.paint.fontSize,
        bold: format.paint.bold,
        wrap: format.paint.wrap
      });
      found.set(cell.row, list);
    }
    if (found.size === 0) {
      return;
    }
    this.rowFitPending = rows === 'all' ? 'all' : new Set(rows);
    this.rowFitSubject.next({
      serial: this.rowFitSubject.value.serial + 1,
      rows: [...found].sort((a, b) => a[0] - b[0]).map(([row, cells]) => ({ row, cells }))
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
    this.fitRowsLater('all');
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
    this.rowsShownChanged();
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
    this.rowsShownChanged();
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
    this.rowsShownChanged();
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.persist();
  }

  /**
   * Which rows show changed, which no cell did — so the `SUBTOTAL`s,
   * the one kind of formula that reads it, are told and the pump
   * picks them up.
   */
  private rowsShownChanged(): void {
    this.document.book.visibilityChanged();
    this.pump();
  }

  clearFilter(): void {
    if (this.document.filteredRows.size === 0) {
      return;
    }
    this.document.filteredRows.clear();
    this.rowsShownChanged();
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
    this.publishActiveRules();
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
    let stored: SheetSnapshot | null;
    try {
      stored = (await this.repository?.load()) ?? null;
    } catch (error) {
      // Not restored, so nothing is persisted over what could not be
      // read; see `SheetRepository.load`.
      console.warn('[sheet] could not read the workbook; it was left as it was.', error);
      this.publishSheet();
      return;
    }
    if (stored === null) {
      seed?.(this.document);
      this.document.sheet.recalculate();
    } else {
      applySnapshot(this.document, stored);
    }
    this.restored = true;
    /**
     * The same publish both ways round, which is the whole of the fix
     * this used to need.
     *
     * A loaded file went through `publishSheet` and a seeded one went
     * through a hand-written list of publishes that looked complete
     * and was not: it left out the tabs, the geometry and the
     * painter's rules. A seed of one plain sheet has no tabs worth
     * drawing, no merges and no rules, so for six phases the gap cost
     * nothing and stayed invisible — and the first seed that had all
     * three opened with its banner clipped to one column, its freezes
     * gone, its conditional formats absent and a tab strip claiming
     * one sheet over a workbook of three.
     *
     * Two paths that have to publish the same thing are one path.
     */
    this.publishSheet();
    if (stored === null) {
      this.persist();
    }
  }

  // ---------------------------------------------------------------------
  // Documents
  // ---------------------------------------------------------------------

  /**
   * Shows a document from the library, by id.
   *
   * `''` is "whichever was used last", which is what a tab opened at
   * `/` wants, and `'new'` is a blank one. An id the library does not
   * know falls back to the last used rather than to nothing: a url
   * from another browser is somebody's mistake, and the sheet is the
   * right thing to show them. The very first document of all starts
   * from the seed; every document after it starts empty.
   *
   * Opening the document already open does nothing, which is what
   * lets the render worker send this whenever its route changes
   * without caring whether the change was its own.
   */
  openDocument(id: string): void {
    this.enqueue(() => this.open(id));
  }

  /**
   * Opens what a file held: a workbook as a document of its own, and
   * anything else through `importCsv`.
   *
   * A workbook is matched to the document it was last opened as by its
   * handle, so reopening a file from the recent list shows the same
   * document rather than a second copy of it — with the file's
   * contents, since the file is what somebody chose to open.
   */
  openFile(fileName: string, text: string, handle: number | null): void {
    if (!/\.(gsheet|json)$/i.test(fileName)) {
      this.importCsv(fileName, text);
      return;
    }
    this.enqueue(async () => {
      const snapshot = parseSnapshot(text, this.geometrySubject.value.columnCount);
      if (snapshot === null) {
        this.report(`${fileName} was not opened: it is not a workbook this can read.`);
        return;
      }
      const entries = (await this.library?.entries()) ?? [];
      const known = handle === null ? undefined : entries.find(each => each.file?.handle === handle);
      const entry: DocumentEntry = {
        id: known?.id ?? this.library?.newId() ?? 'file',
        name: baseName(fileName),
        used: this.now(),
        file: { handle, name: fileName }
      };
      await this.swapTo(entry, snapshot);
      // Written through at once: the file is the source, and a copy
      // that waited for the first edit to be kept would be a document
      // that vanished from the library if the tab closed first.
      this.persist();
      this.edited = false;
      this.publishDocument();
      this.report(`Opened ${fileName}.`);
    });
  }

  /**
   * Opens an Excel workbook as a document of its own.
   *
   * A document with no file, and that is deliberate: this application
   * reads `.xlsx` and does not write it, so a Save afterwards asks
   * where to put a `.gsheet` rather than overwriting the workbook it
   * came from with a different format. The bytes cross as base64,
   * because a command carries plain data and a buffer is not.
   */
  importXlsx(fileName: string, base64: string): void {
    this.enqueue(async () => {
      const { rowCount, columnCount } = this.geometrySubject.value;
      let book;
      try {
        book = await openXlsx(bytesOfBase64(base64), platformInflate, { rows: rowCount, columns: columnCount });
      } catch (error) {
        const why = error instanceof XlsxError ? error.message : 'it could not be read.';
        this.report(`${fileName} was not opened: ${why.charAt(0).toLowerCase()}${why.slice(1)}`);
        return;
      }
      const entry: DocumentEntry = {
        id: this.library?.newId() ?? 'xlsx',
        name: baseName(fileName),
        used: this.now(),
        file: null
      };
      await this.swapTo(entry, snapshotOfXlsx(book, columnCount));
      this.persist();
      this.edited = false;
      this.publishDocument();
      this.report(reportOfXlsx(fileName, book));
    });
  }

  /**
   * Builds the workbook as a `.gsheet` and publishes it on `transfer`
   * for the render worker to hand to the shell.
   *
   * With the file it was last saved to, unless `asNew` — which is the
   * difference between Save and Save As, and the render worker's
   * `saveFile` call turns it into a write or a picker.
   */
  saveDocument(asNew: boolean): void {
    const file = this.entry?.file ?? null;
    const handle = !asNew && file !== null ? file.handle : null;
    this.publishDownload({
      kind: 'workbook',
      name: file?.name ?? `${this.entry?.name ?? 'Untitled'}.gsheet`,
      mediaType: 'application/json',
      text: JSON.stringify(this.snapshot()),
      handle
    });
  }

  /**
   * What became of a download: where the shell wrote it, or that it
   * did not.
   *
   * A workbook saved to a file makes that file the document's, and the
   * document stops being edited. A download is recorded with no
   * handle, because there is nothing to save back to — the next Save
   * asks where again, which is the honest answer in a browser that
   * cannot write to a file it did not open.
   */
  fileSaved(kind: 'workbook' | 'csv' | 'xlsx', name: string, handle: number | null, via: 'file' | 'download'): void {
    const verb = via === 'file' ? 'Saved' : 'Downloaded';
    // An export is a copy: the document goes on being the one it was,
    // saved where it was saved, and edited if it was.
    if (kind === 'csv' || kind === 'xlsx') {
      this.report(`${verb} ${name}.${kind === 'xlsx' ? this.xlsxLeftOut : ''}`);
      return;
    }
    this.enqueue(async () => {
      if (this.entry !== null) {
        this.entry = { ...this.entry, name: baseName(name), file: { handle, name } };
        await this.library?.put(this.entry);
      }
      this.edited = false;
      this.publishDocument();
      this.report(`${verb} ${name}.`);
    });
  }

  reportFile(text: string): void {
    this.report(text);
  }

  /** Settles when every open asked for so far has. For specs, and the proof. */
  get settled(): Promise<void> {
    return this.opening;
  }

  private enqueue(run: () => Promise<void>): void {
    this.opening = this.opening.then(run).catch(error => {
      console.warn('[sheet] could not open the document.', error);
      this.report('That could not be opened, and nothing was changed. Reload to try again.');
    });
  }

  private async open(id: string): Promise<void> {
    const library = this.library;
    if (library === undefined || (id !== 'new' && id !== '' && id === this.entry?.id)) {
      return;
    }
    let entries: DocumentEntry[];
    try {
      entries = await library.entries();
    } catch (error) {
      // Without the list there is no telling which documents exist,
      // and guessing "none" would seed a first one and index it alone.
      console.warn('[sheet] could not read the document index.', error);
      this.report('The list of workbooks could not be read. Reload to try again.');
      return;
    }
    const latest = [...entries].sort((a, b) => b.used - a.used)[0];
    const found = id === 'new' ? undefined : id === '' ? latest : (entries.find(each => each.id === id) ?? latest);
    if (found !== undefined && found.id === this.entry?.id) {
      return;
    }
    const first = found === undefined && id !== 'new' && entries.length === 0;
    const entry: DocumentEntry = found ?? {
      id: first ? FIRST_DOCUMENT : library.newId(),
      name: 'Untitled',
      used: this.now(),
      file: null
    };
    let stored: SheetSnapshot | null;
    try {
      stored = await library.repository(entry.id).load();
    } catch (error) {
      // There is a workbook there and it could not be read. Opening an
      // empty one in its place would save that empty one over it, so
      // nothing is opened and nothing is saved; the tab stays on what
      // it had, which on a first open is a blank that is kept nowhere.
      console.warn('[sheet] could not read the document; it was left as it was.', error);
      this.report(`${entry.name} could not be read, so it was left as it was. Reload to try again.`);
      return;
    }
    await this.swapTo({ ...entry, used: this.now() }, stored, first ? this.seed : undefined);
    this.edited = false;
    this.publishDocument();
    if (this.elsewhere) {
      // Not taken over when the other tab closes, and on purpose: this
      // copy may be older than what that tab last saved, and keeping it
      // then would write over those edits — the loss the claim is for.
      // The status line already says it is open elsewhere; this says
      // what to do about it, and short enough to stay on one line.
      this.report('Close the other tab and reload this one to edit here.');
    }
  }

  /**
   * Puts a fresh workbook in front of somebody, from a snapshot or a
   * seed, and makes `entry` the open document.
   *
   * Everything this service remembers *about* the last workbook goes
   * with it: the paints a rule interned, the chart with handles round
   * it, the matches of a search, the block last copied and where it
   * came from. Each of those is an index or a position into a
   * workbook that is no longer the one on screen.
   */
  private async swapTo(
    entry: DocumentEntry,
    stored: SheetSnapshot | null,
    seed?: (document: SheetDocument) => void
  ): Promise<void> {
    await this.repository?.flush();
    this.release?.();
    this.release = null;
    this.elsewhere = false;
    if (this.library !== undefined) {
      this.release = await this.library.claim(entry.id);
      this.elsewhere = this.release === null;
    }
    this.restored = false;
    const document = new SheetDocument();
    document.columnWidths = this.defaultWidths();
    this.document = document;
    this.repository = this.library?.repository(entry.id) ?? this.repository;
    this.refusal = '';
    this.extraPaints.length = 0;
    this.extraIds.clear();
    this.publishedBase = 0;
    this.selectedChart = 0;
    this.hiddenWidths.clear();
    this.copied = null;
    this.clearFind();
    if (stored === null) {
      seed?.(document);
      document.sheet.recalculate();
    } else {
      applySnapshot(document, stored);
    }
    this.restored = true;
    this.entry = entry;
    await this.library?.put(entry);
    this.publishSheet();
    this.publishNames('');
    this.publishPalette();
    if (stored === null) {
      this.persist();
    }
    this.pump();
  }

  private publishDocument(): void {
    this.documentSubject.next({
      id: this.entry?.id ?? '',
      name: this.entry?.name ?? '',
      file: this.entry?.file ?? null,
      edited: this.edited,
      elsewhere: this.elsewhere
    });
  }

  private publishDownload(download: Omit<SheetDownload, 'serial'>): void {
    this.downloadSerial++;
    this.transferSubject.next({ ...this.transferSubject.value, download: { ...download, serial: this.downloadSerial } });
  }

  private report(report: string): void {
    this.transferSubject.next({ ...this.transferSubject.value, report });
  }

  // ---------------------------------------------------------------------
  // Charts
  // ---------------------------------------------------------------------

  /**
   * A chart over the selection, at a default size beside it.
   *
   * Placed to the right of the range it reads rather than on top of
   * it, which is what every spreadsheet does and is the only position
   * that does not hide the numbers the chart is of.
   */
  insertChart(kind: ChartKind): void {
    const rect = rectOf(this.document.selection);
    const geometry = this.geometrySubject.value;
    const right = offsetOfColumn(geometry, rect.lastColumn + 1) + 16;
    const top = rect.firstRow * geometry.rowHeight;
    const id = this.document.addChart({
      kind,
      title: '',
      range: {
        start: relativeRef(rect.firstRow, rect.firstColumn),
        end: relativeRef(rect.lastRow, rect.lastColumn)
      },
      place: { x: right, y: top, width: DEFAULT_CHART_WIDTH, height: DEFAULT_CHART_HEIGHT },
      legend: true
    });
    this.selectedChart = id;
    this.publishCharts();
    this.publishSeries();
    this.persist();
    this.publishStatus();
  }

  selectChart(id: number): void {
    if (this.selectedChart === id) {
      return;
    }
    this.selectedChart = id;
    this.publishCharts();
  }

  placeChart(id: number, x: number, y: number, width: number, height: number): void {
    const place = {
      x: Math.max(0, Math.round(x)),
      y: Math.max(0, Math.round(y)),
      width: Math.max(MIN_CHART_WIDTH, Math.round(width)),
      height: Math.max(MIN_CHART_HEIGHT, Math.round(height))
    };
    if (!this.document.changeChart(id, chart => ({ ...chart, place }))) {
      return;
    }
    this.publishCharts();
    // A chart that changed width can draw more points, or fewer, and
    // the series it is sent is cut to its width. This is the only
    // place a *move* has to republish data, and it is why the two
    // keys are worth having: a drag that only moved it republishes
    // the placement and finds the series structurally equal.
    this.publishSeries();
    this.persist();
    this.publishStatus();
  }

  setChartKind(id: number, kind: ChartKind): void {
    if (!this.document.changeChart(id, chart => ({ ...chart, kind }))) {
      return;
    }
    this.publishCharts();
    this.persist();
    this.publishStatus();
  }

  setChartTitle(id: number, title: string): void {
    if (!this.document.changeChart(id, chart => ({ ...chart, title }))) {
      return;
    }
    this.publishCharts();
    this.persist();
    this.publishStatus();
  }

  removeChart(id: number): void {
    if (!this.document.removeChart(id)) {
      return;
    }
    if (this.selectedChart === id) {
      this.selectedChart = 0;
    }
    this.publishCharts();
    this.publishSeries();
    this.persist();
    this.publishStatus();
  }

  private publishCharts(): void {
    this.chartsSubject.next({
      entries: this.document.charts.map(chart => ({
        id: chart.id,
        kind: chart.kind,
        title: chart.title,
        range: formatRange(chart.range),
        x: chart.place.x,
        y: chart.place.y,
        width: chart.place.width,
        height: chart.place.height,
        legend: chart.legend
      })),
      selected: this.selectedChart
    });
  }

  /**
   * What every chart on this sheet draws, cut to what it can draw.
   *
   * **This is where the phase's exit criterion is actually met.** The
   * limit handed to `seriesFrom` is the chart's own width in pixels,
   * so a chart four hundred pixels wide is sent at most four hundred
   * points however many cells its range covers — and the cells are
   * read here, on the thread that has them, rather than crossing.
   *
   * The whole set is rebuilt rather than the one chart that changed.
   * A sheet holds a handful of charts, each bounded by its own width,
   * so the work is bounded by pixels on screen twice over; and the
   * differ compares what comes out, so a rebuild that produced the
   * same numbers puts nothing on the wire.
   */
  private publishSeries(): void {
    const charts = this.document.charts;
    if (charts.length === 0) {
      // Structurally equal to the last empty one, so this costs a
      // comparison and no patch on every sheet that has no charts —
      // which is every sheet, on the page `pnpm proof` measures.
      this.seriesSubject.next({ charts: {} });
      return;
    }
    const built: Record<string, SheetChartSeries> = {};
    for (const chart of charts) {
      built[String(chart.id)] = this.seriesFor(chart);
    }
    this.seriesSubject.next({ charts: built });
  }

  private seriesFor(chart: Chart): SheetChartSeries {
    const rect = {
      firstRow: Math.min(chart.range.start.row, chart.range.end.row),
      lastRow: Math.max(chart.range.start.row, chart.range.end.row),
      firstColumn: Math.min(chart.range.start.column, chart.range.end.column),
      lastColumn: Math.max(chart.range.start.column, chart.range.end.column)
    };
    const grid: CellValue[][] = [];
    for (let row = rect.firstRow; row <= rect.lastRow; row++) {
      const line: CellValue[] = [];
      for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
        line.push(this.document.sheet.value(row, column));
      }
      grid.push(line);
    }
    const rows = rect.lastRow - rect.firstRow + 1;
    const columns = rect.lastColumn - rect.firstColumn + 1;
    const { byColumn } = orientationOf(rows, columns);
    const lines = byColumn ? grid : transposed(grid, rows, columns);
    const headers = headersIn(lines[0] ?? [], lines[1]);
    const labels = (lines[headers ? 1 : 0] ?? []).some(value => typeof value === 'string');
    const read = seriesFrom(grid, {
      byColumn,
      headers,
      labels,
      // The chart's own width, which is the most points a line drawn
      // in it can be told apart at. A pie has no width to speak of
      // and reads one series, so it is cut the same way and never
      // notices.
      limit: chart.place.width
    });
    return { categories: read.categories, series: read.series, read: read.read };
  }

  /**
   * Fifty thousand readings, and a chart of them.
   *
   * The proof page's second instrument, beside the recalculation
   * chain and built the same way: the cells go *past the end of the
   * sheet*, where nobody can scroll to them, select them or type in
   * them, and where `snapshotOf` will not write them to anybody's
   * file. A chart reads by key and does not care that the keys are
   * out there, which is the one place that property is useful.
   *
   * Past the chain as well as past the sheet, so the two instruments
   * do not write over each other.
   *
   * Built once. Calling it again re-selects the chart rather than
   * making a second one, so a script that presses the button twice
   * measures one chart.
   */
  chartStress(points: number): void {
    const sheet = this.document.sheet;
    const geometry = this.geometrySubject.value;
    if (this.chartPoints !== points) {
      const first = geometry.rowCount + STRESS_CELLS + 16;
      for (let at = 0; at < points; at++) {
        // A sine rather than a ramp, so the thinning has extremes to
        // keep and the picture shows whether it kept them.
        sheet.setCell(first + at, 0, String(Math.round(Math.sin(at / 400) * 1000)));
      }
      this.chartPoints = points;
      this.stressChart = this.document.addChart({
        kind: 'line',
        title: `${points.toLocaleString('en-US')} readings`,
        range: { start: relativeRef(first, 0), end: relativeRef(first + points - 1, 0) },
        place: {
          x: offsetOfColumn(geometry, 3),
          y: geometry.rowHeight * 2,
          width: DEFAULT_CHART_WIDTH,
          height: DEFAULT_CHART_HEIGHT
        },
        legend: false
      });
    }
    this.selectedChart = this.stressChart;
    this.publishCharts();
    this.publishSeries();
    this.publishStatus();
    this.pump();
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
    if (!this.restored || this.elsewhere) {
      return;
    }
    this.repository?.save(this.snapshot());
    if (!this.edited && this.entry !== null) {
      this.edited = true;
      this.publishDocument();
    }
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
    // A formatted cell's alignment can follow its value — a currency
    // cell whose formula starts returning an error moves from right to
    // left — so the formats go out again while any are in view. A view
    // of nothing but default cells still pays nothing.
    if (!this.painter.isEmpty || this.formattedInView) {
      this.publishFormats();
    }
  }

  /**
   * The series again, but only when there is a chart to draw them.
   *
   * The guard is the whole reason this is a method. A chart's range
   * is read cell by cell on every publish, and a sheet with no charts
   * must not pay a single read for the feature — which is what the
   * page `pnpm proof` measures, and what its budgets would notice.
   */
  private redrawCharts(): void {
    if (this.document.charts.length > 0) {
      this.publishSeries();
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
    this.redrawCharts();
    this.fitRowsLater('all');
    this.persist();
    this.pump();
  }

  private afterHistory(): void {
    this.selectionSubject.next(this.document.selection);
    // Undo can take back the insert of a chart as well as the cells
    // one reads, so both keys go out rather than only the series.
    this.publishCharts();
    this.publishSeries();
    // An undone column insert puts the widths back where they were,
    // and the geometry is where the render worker reads them.
    this.publishGeometry();
    this.publishWindow();
    this.publishFormats();
    this.publishPalette();
    this.publishEditor();
    this.publishActiveRules();
    this.publishStatus();
    this.publishStats();
    this.fitRowsLater('all');
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
    // A formula settling changes what a chart of it draws, and a
    // chart is the one thing on screen that can be looking at cells
    // the window is not.
    this.redrawCharts();
    if (result.done) {
      this.pumping = false;
      // A formula's new value is new text, and wrapped text may now
      // take more lines or fewer.
      this.fitRowsLater('all');
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
    let formatted = false;
    for (const row of this.rowsInView()) {
      const line: Record<string, number> = {};
      for (const column of columns) {
        const id = this.paintedId(row, column);
        if (id !== 0) {
          line[column] = id;
          formatted = true;
        }
      }
      cells[row] = line;
    }
    this.formattedInView = formatted;
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
    if (base === 0 && this.painter.isEmpty) {
      // An unformatted cell shows a number as a number and text as
      // text, so the grid's reading of the string is already right.
      return base;
    }
    const value = this.document.sheet.value(row, column);
    const over = this.painter.isEmpty ? null : this.painter.paintFor(row, column, value);
    const own = this.document.formats.byId(base).paint;
    const painted: CellPaint =
      over === null
        ? own
        : {
            ...own,
            ...(over.fill === undefined ? {} : { fill: over.fill }),
            ...(over.color === undefined ? {} : { color: over.color }),
            ...(over.bold === undefined ? {} : { bold: over.bold }),
            ...(over.italic === undefined ? {} : { italic: over.italic })
          };
    const aligned = this.alignedFor(painted, row, column, value);
    if (over === null && aligned === painted) {
      return base;
    }
    // A *position* in the extras, resolved against the document's
    // palette at publish time — so the document growing a format
    // moves every extra index and the palette goes out with it.
    return this.paletteBase() + this.internPaint(aligned);
  }

  /**
   * The paint with its `auto` alignment made explicit, where the grid
   * would otherwise get it wrong.
   *
   * The grid places an `auto` cell by whether its *string* reads as a
   * number, because the string is all it is sent. A format is exactly
   * what makes the string lie: `$4.50`, `9.6%` and `2026-09-24` are
   * numbers that do not read as one, and `007` in a Text cell is text
   * that does. For those, and only those, the alignment the value
   * wants is written into the paint — so a correct guess costs nothing
   * on the wire, and a wrong one costs one palette entry per format.
   */
  private alignedFor(paint: CellPaint, row: number, column: number, value: CellValue): CellPaint {
    if (paint.align !== 'auto' || value === null || value === '') {
      return paint;
    }
    const wants = placeOf(value);
    if (guessOf(this.document.display(row, column)) === wants) {
      return paint;
    }
    return { ...paint, align: wants };
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

  private publishActiveRules(): void {
    const { row, column } = this.document.selection;
    const conditional = this.document.conditionalAt(row, column);
    const validation = this.document.validationAt(row, column);
    this.activeRulesSubject.next({
      conditional:
        conditional === null
          ? null
          : {
              test: conditional.test,
              ...(conditional.paint === undefined ? {} : { paint: conditional.paint }),
              ...(conditional.scale === undefined ? {} : { scale: conditional.scale })
            },
      validation: validation === null ? null : { rule: validation.rule, strict: validation.strict === true }
    });
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
      canRedo: this.document.canRedo,
      iterating: this.document.book.iteration !== null
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

/**
 * How far along the sheet a column starts, in its own pixels.
 *
 * The render worker asks `UiVirtualSheet` for this and gets it from a
 * running total it already keeps; this side has only the list, and a
 * chart is placed once rather than per frame. Columns past the end of
 * the list take the default width, which is what the grid draws them
 * at.
 */
function offsetOfColumn(geometry: SheetGeometry, column: number): number {
  let offset = 0;
  for (let at = 0; at < column; at++) {
    offset += geometry.columnWidths[at] ?? geometry.columnWidth;
  }
  return offset;
}

/** A grid of values with its rows and columns swapped. */
function transposed(grid: readonly (readonly CellValue[])[], rows: number, columns: number): CellValue[][] {
  const out: CellValue[][] = [];
  for (let column = 0; column < columns; column++) {
    const line: CellValue[] = [];
    for (let row = 0; row < rows; row++) {
      line.push(grid[row]?.[column] ?? null);
    }
    out.push(line);
  }
  return out;
}
