import { parseTypedDate } from '../sheet/Dates';
import { COLUMN_WIDTH, ROW_HEIGHT } from './dimensions';
import { formatWith, type CellFormat, type NumberFormat } from '../sheet/Format';
import { Formats, type FormatPlacement } from '../sheet/Formats';
import { Merges } from '../sheet/Merges';
import type { Chart } from '../sheet/Chart';
import { columnName, MAX_COLUMNS as MAX_COLUMNS_HERE, MAX_ROWS as MAX_ROWS_HERE, type RangeRef } from '../sheet/A1';
import { Notes, type Note } from '../sheet/Notes';
import { ScenarioInputs, type Scenario, type ScenarioInput } from '../sheet/ScenarioInputs';
import type { ConditionalRule } from '../sheet/Conditional';
import type { Validation } from '../sheet/Validation';
import { isNamedRange, nameProblem, type DefinedName, type NameProblem } from '../sheet/Names';
import { FormulaSyntaxError, parseFormula } from '../sheet/Parser';
import { Sheet } from '../sheet/Sheet';
import type { SheetSelection } from './SheetContract';
import { literalOf, Workbook } from '../sheet/Workbook';
import { validate } from '../sheet/Validation';
import { shiftIndex, shiftRange, type Shift } from '../sheet/Shift';
import type { Script } from '../script/ScriptHost';

/**
 * One change to one cell, and what it replaced.
 *
 * Two kinds rather than one since Phase 9, because formatting a range
 * and typing into it are both things a person undoes and neither is
 * the other. Keeping them in the same list is what makes a paste that
 * carried formats one press of ctrl-Z rather than two.
 */
type Edit = TextEdit | FormatEdit | RegionEdit | StructureEdit | NamesEdit | RulesEdit | ChartsEdit | NoteEdit | ScenarioInputEdit;

/**
 * Where an edit was made, which every kind of edit has to say.
 *
 * Undo is the workbook's rather than the sheet's — ctrl-Z takes back
 * the last thing you did, wherever you did it — so taking a step back
 * means going to the sheet it was made on first. A stack per sheet
 * would be a ctrl-Z whose meaning depended on which tab happened to
 * be showing.
 */
interface OnASheet {
  readonly sheet: number;
}

/**
 * A name defined, redefined or removed.
 *
 * The whole table either side rather than the one entry that changed,
 * which is the cheap answer here and not a lazy one: a sheet holds a
 * handful of names, the table is a list of short strings and four
 * numbers each, and an edit that carries it whole cannot get the
 * inverse wrong. The row and column are the range's corner, so
 * undoing a definition puts the selection back on what was named.
 */
interface NamesEdit extends OnASheet {
  readonly kind: 'names';
  readonly row: number;
  readonly column: number;
  readonly before: readonly DefinedName[];
  readonly after: readonly DefinedName[];
}

/**
 * The conditional formats and validations of a sheet, before and
 * after.
 *
 * Both lists whole rather than the entry that changed, on the same
 * rule the names go by: a sheet holds a handful of rules, an edit
 * that carries them whole cannot get its own inverse wrong, and the
 * saving from being cleverer is a few hundred bytes.
 */
interface RulesEdit extends OnASheet {
  readonly kind: 'rules';
  readonly row: number;
  readonly column: number;
  readonly before: { readonly conditional: readonly ConditionalRule[]; readonly validations: readonly Validation[] };
  readonly after: { readonly conditional: readonly ConditionalRule[]; readonly validations: readonly Validation[] };
}

/**
 * The charts of a sheet, before and after.
 *
 * The list whole, on the same rule the names and the rules go by: a
 * sheet holds a handful of charts, each is a title and eight numbers,
 * and an edit that carries the list whole cannot get its own inverse
 * wrong. The row and column are the chart's own range corner, so
 * undoing an insert puts the selection back on the cells it was
 * drawn from.
 */
interface ChartsEdit extends OnASheet {
  readonly kind: 'charts';
  readonly row: number;
  readonly column: number;
  readonly before: readonly Chart[];
  readonly after: readonly Chart[];
}

/** A cell's note, before and after; empty is no note. */
interface NoteEdit extends OnASheet {
  readonly kind: 'note';
  readonly row: number;
  readonly column: number;
  readonly before: string;
  readonly after: string;
}

/**
 * A cell a scenario types differently, before and after; null is the
 * base's own input. Typing while a scenario is shown is an edit like
 * any other, and ctrl-Z takes it back like any other.
 */
interface ScenarioInputEdit extends OnASheet {
  readonly kind: 'scenarioInput';
  readonly scenario: string;
  readonly row: number;
  readonly column: number;
  readonly before: string | null;
  readonly after: string | null;
}

interface TextEdit extends OnASheet {
  readonly kind: 'text';
  readonly row: number;
  readonly column: number;
  readonly before: string;
  readonly after: string;
}

/** A cell's palette id, before and after. Ids, not formats: the
 * palette only grows, so an id undoes to an entry that is still
 * there, and a step stays four numbers however large the format is. */
interface FormatEdit extends OnASheet {
  readonly kind: 'format';
  readonly row: number;
  readonly column: number;
  readonly before: number;
  readonly after: number;
}

/**
 * A row or column inserted or deleted.
 *
 * Recorded as the shift plus what the shift destroyed, because a
 * shift is not reversible on its own: deleting a row takes the cells
 * in it away and turns every reference to them into `#REF!`, and
 * neither comes back from applying the opposite shift. So the cells
 * that were removed and the formulas that were rewritten are kept, at
 * the positions they had *before* — which is where the inverse shift
 * puts everything back.
 *
 * `rewritten` is proportional to the formulas that actually mentioned
 * the line, not to the sheet. On a normal sheet that is tens; on a
 * column of fifty thousand chained formulas it is fifty thousand,
 * which is how much information the edit really changed and the price
 * of being able to take it back.
 */
interface StructureEdit extends OnASheet {
  readonly kind: 'structure';
  readonly row: number;
  readonly column: number;
  readonly shift: Shift;
  /** Cells the shift removed, at the positions they had. */
  readonly removed: readonly { readonly row: number; readonly column: number; readonly input: string }[];
  /** Formulas the shift changed, as they read before it. */
  readonly rewritten: readonly { readonly row: number; readonly column: number; readonly input: string }[];
  /** Column widths as they were, so undo puts them back. */
  readonly widths: readonly number[];
  /** And the rows' heights and which rows were hidden, for the same reason. */
  readonly rows: RowsHeld;
  /** The sheet's notes as they were, all of them: a deleted row takes its notes with it. */
  readonly notes: readonly Note[];
  /**
   * Every sheet's scenario inputs as they were. Those on the sheet that
   * changed shape move with its cells, a formula among them on any sheet
   * may point at cells that moved, and one on a deleted row is gone.
   */
  readonly scenarioInputs: readonly (readonly ScenarioInput[])[];
  /**
   * Where the formats were, for a deletion only: a deleted row takes
   * its formats with it, and shifting back cannot bring them. An insert
   * destroys nothing and keeps nothing.
   */
  readonly formats: FormatPlacement | null;
  /**
   * Every sheet's rules and charts as they were. The rules of the sheet
   * that changed shape move with its cells, and so do the charts on any
   * sheet that read those cells, so an undo has more than one page to
   * put back — and a rule over rows that were deleted is gone, which
   * shifting back cannot undo either.
   */
  readonly ranges: readonly PageRanges[];
}

/** One page's rules and charts, held whole for an undo. */
interface PageRanges {
  readonly conditional: readonly ConditionalRule[];
  readonly validations: readonly Validation[];
  readonly charts: readonly Chart[];
}

/** What a row shift moves besides the cells, held so undo can put it back. */
interface RowsHeld {
  readonly heights: readonly (readonly [number, number])[];
  readonly fitted: readonly (readonly [number, number])[];
  readonly hidden: readonly number[];
  readonly filtered: readonly number[];
}

/**
 * A whole row, column or sheet formatted at once.
 *
 * One entry however many cells it covers, which is the point: a
 * million cell edits for one press of ctrl-B is a step nobody can
 * afford to keep and a file nobody asked for.
 *
 * `overrides` is the cells inside the region that had a format of
 * their own, before and after — the change is applied to those too,
 * so that making a sheet bold keeps the currency in column C. It is
 * bounded by how many cells anybody has actually formatted.
 */
interface RegionEdit extends OnASheet {
  readonly kind: 'region';
  readonly scope: 'sheet' | 'row' | 'column';
  readonly index: number;
  readonly row: number;
  readonly column: number;
  readonly before: number;
  readonly after: number;
  readonly overrides: readonly { readonly key: number; readonly before: number; readonly after: number }[];
}

/**
 * What one press of ctrl-Z takes back.
 *
 * A step rather than an edit, because a paste is one action and fifty
 * cells: undoing it a cell at a time would be fifty presses to
 * reverse one, which is the kind of thing that makes people stop
 * trusting undo.
 */
type Step = readonly Edit[];

/**
 * The sheet plus the two things a document has that a model does not:
 * where the selection is, and what can be taken back.
 *
 * Still plain TypeScript — no RxJS, no framework — so it runs in node
 * beside `Sheet` and is the layer a spec drives when it wants to ask
 * about behaviour rather than about the wire.
 *
 * Undo is per commit, not per keystroke. Phase 4 brings coalescing
 * when there is a cell editor to coalesce, and `EditableTextModel`
 * already has its own; what belongs here is the commit, which is the
 * unit a person means when they press ctrl-Z in a spreadsheet.
 */
/**
 * Everything about one sheet that is not its cells.
 *
 * All of it is per sheet and none of it is shared, which is the
 * answer to the question this phase kept asking: column widths,
 * hidden rows, freezes, merges, formats and where the selection sits
 * are facts about *a* sheet. Two things are the workbook's instead —
 * the names, because `=SUM(Sales)` has to mean the same cells
 * wherever it is written, and the undo stack, because ctrl-Z takes
 * back the last thing you did rather than the last thing you did
 * here.
 */
interface Page {
  readonly sheet: Sheet;
  readonly formats: Formats;
  readonly merges: Merges;
  /** The notes somebody left on cells; see `Notes`. */
  readonly notes: Notes;
  /** What each scenario types into this sheet's cells; see `ScenarioInputs`. */
  readonly scenarioInputs: ScenarioInputs;
  /**
   * How wide each column is drawn.
   *
   * Here rather than on the service's geometry since Phase 10, and
   * the reason is undo. A column insert moves the widths along with
   * the columns they describe, and taking that back has to move them
   * back — so the widths have to be somewhere the undo stack can
   * reach.
   */
  columnWidths: number[];
  /**
   * The rows somebody has hidden.
   *
   * A set of exceptions rather than a height per row, for the reason
   * `UiVirtualSheetOptions.rowHeights` is sparse: a sheet is ten
   * thousand rows tall and all but a handful are the same.
   */
  readonly hiddenRows: Set<number>;
  /**
   * Rows somebody made taller or shorter by hand, by index.
   *
   * Sparse for the reason the hidden rows are. A height set by hand is
   * kept whatever the row holds, which is Excel's rule: a row dragged
   * to a height has been told what it is.
   */
  readonly rowHeights: Map<number, number>;
  /**
   * Rows made taller to fit what they hold: wrapped text, a large
   * font. Worked out by the render worker, which is the only thread
   * that can measure, and kept here so the sheet opens at the heights
   * it had rather than growing a frame after it appears. Only rows
   * that are not the default are in it, and a height set by hand wins.
   */
  readonly fittedRows: Map<number, number>;
  /**
   * Rows a filter is hiding, kept apart from the ones somebody hid.
   *
   * Two sets rather than one because they are undone by different
   * things: clearing a filter must not reveal a row that was hidden
   * on purpose, and showing a hidden row must not fight the filter
   * that is hiding it. What the screen sees is the union.
   */
  readonly filteredRows: Set<number>;
  /**
   * The column the filter was chosen by, for the mark in its header;
   * -1 with no filter. Kept as the rows are, for the session.
   */
  filterColumn: number;
  /** How many rows and columns stay put while the rest scrolls. */
  frozenRows: number;
  frozenColumns: number;
  /**
   * How large this sheet is drawn: 1 is 100%. A way of looking, like
   * the frozen panes, so it is saved with the sheet and is not an edit.
   */
  zoom: number;
  selection: SheetSelection;
  /**
   * Formats that think, and what a cell is allowed to hold.
   *
   * Per sheet like everything else drawn over the cells, and held as
   * plain lists rather than indexed by cell: a rule is a fact about
   * a *range*, and a million-cell range indexed per cell would be
   * the million entries this whole design exists to avoid.
   */
  conditional: ConditionalRule[];
  validations: Validation[];
  /**
   * The charts floating over this sheet.
   *
   * Per sheet like everything else drawn over the cells. A chart
   * reads a range on *its own* sheet, so a chart that followed the
   * workbook would be a picture of cells nobody was looking at.
   */
  charts: Chart[];
}

function newPage(sheet: Sheet): Page {
  return {
    sheet,
    formats: new Formats(),
    merges: new Merges(),
    notes: new Notes(),
    scenarioInputs: new ScenarioInputs(),
    columnWidths: [],
    hiddenRows: new Set<number>(),
    rowHeights: new Map<number, number>(),
    fittedRows: new Map<number, number>(),
    filteredRows: new Set<number>(),
    filterColumn: -1,
    frozenRows: 0,
    frozenColumns: 0,
    zoom: 1,
    selection: { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 },
    conditional: [],
    validations: [],
    charts: []
  };
}

export class SheetDocument {
  readonly book = new Workbook();
  private readonly pages: Page[] = [newPage(this.book.sheet(0))];
  /**
   * Which sheet the tabs are showing, and which one every command
   * without a sheet of its own means.
   *
   * One field rather than a sheet argument on forty commands, which
   * is the trade the roadmap called for: the viewport names its sheet
   * and everything else follows it.
   */
  private activeSheet = 0;

  private readonly undoStack: Step[] = [];
  private readonly redoStack: Step[] = [];

  /**
   * The workbook's scripts, in the order they were written.
   *
   * The workbook's rather than a sheet's, as the names are. Changing
   * one is not an edit on the undo stack: undo is for what a sheet
   * holds, and a script is closer to a file kept beside it. What a
   * script's *run* changes is one step of undo like any other.
   */
  scripts: Script[] = [];

  /**
   * The workbook's scenarios: named ways it might have gone, each a few
   * inputs typed differently. The list is the workbook's, as the names
   * are; what each one types is kept on the sheets it types into (see
   * `ScenarioInputs`). Adding, renaming and deleting one is not on the
   * undo stack, for the reason a script is not: it is closer to a file
   * kept beside the sheet than to what the sheet holds. Typing into one
   * is.
   */
  scenarios: Scenario[] = [];

  constructor() {
    // Which rows show is the document's to know and `SUBTOTAL`'s to
    // ask; see `Workbook.rowState`. A filter wins over a hand-hidden
    // row, because 1–11 leave out the filtered and keep the hidden.
    this.book.rowState = (sheet, row) => {
      const page = this.pages[sheet];
      if (page === undefined) {
        return null;
      }
      return page.filteredRows.has(row) ? 'filtered' : page.hiddenRows.has(row) ? 'hidden' : null;
    };
  }
  /** Edits collected by an open `transact`, or null outside one. */
  private collecting: Edit[] | null = null;
  /** What the step being collected is called, once something has said. */
  private collectingLabel: string | null = null;
  /**
   * What each step on the stacks is called, for "Undo sort".
   *
   * Beside the steps rather than in them, because a step is a list of
   * edits and every reader of one — undo, redo, `selectStep` — reads it
   * as that. A step nobody named is described from what it holds.
   */
  private readonly labels = new WeakMap<Step, string>();

  // ---------------------------------------------------------------------
  // The active page, which is what every unqualified call means
  // ---------------------------------------------------------------------

  get active(): number {
    return this.activeSheet;
  }

  get page(): Page {
    return this.pages[this.activeSheet];
  }

  get sheet(): Sheet {
    return this.page.sheet;
  }

  get formats(): Formats {
    return this.page.formats;
  }

  get merges(): Merges {
    return this.page.merges;
  }

  get notes(): Notes {
    return this.page.notes;
  }

  noteAt(row: number, column: number): string {
    return this.page.notes.at(row, column);
  }

  /**
   * The scenario on screen, or null for the base.
   *
   * Showing one does two things here and nothing else. What a cell
   * displays is read from that scenario's fork (see `shown`), and what
   * is typed — by hand, by a paste, a fill, a sort — goes into that
   * scenario rather than the base, through `setCell`, which every one of
   * those passes through. Formats, rules and the sheet's shape are
   * shared: a scenario is a question about values.
   */
  private shownScenario: string | null = null;
  /** That scenario's fork, once it has been made; see `SheetService`. */
  private forked: Workbook | null = null;

  get scenario(): string | null {
    return this.shownScenario;
  }

  /** Shows a scenario, or the base with null. Its fork is the caller's to make and hand back. */
  showScenario(id: string | null): void {
    this.shownScenario = id !== null && this.scenarios.some(entry => entry.id === id) ? id : null;
    this.forked = null;
  }

  get fork(): Workbook | null {
    return this.forked;
  }

  set fork(book: Workbook | null) {
    this.forked = this.shownScenario === null ? null : book;
  }

  /** The active sheet as it is on screen: the scenario's fork when one is shown and made, the base's otherwise. */
  get shown(): Sheet {
    return this.forked?.sheet(this.activeSheet) ?? this.sheet;
  }

  /** What a cell holds as typed, in the scenario on screen: its own input where it has one, the base's elsewhere. */
  inputAt(row: number, column: number): string {
    const scenario = this.shownScenario;
    const own = scenario === null ? null : this.page.scenarioInputs.at(scenario, row, column);
    return own ?? this.sheet.input(row, column);
  }

  /** What `scenario` types into a cell of the active sheet, or null where it leaves the base's. */
  scenarioInputAt(scenario: string, row: number, column: number, sheet = this.activeSheet): string | null {
    return this.pages[sheet]?.scenarioInputs.at(scenario, row, column) ?? null;
  }

  /**
   * Types an input into a cell of the active sheet for `scenario`, or
   * with null gives the cell back to the base. One step of undo.
   */
  setScenarioInput(scenario: string, row: number, column: number, input: string | null): void {
    const before = this.page.scenarioInputs.at(scenario, row, column);
    if (before === input) {
      return;
    }
    this.page.scenarioInputs.set(scenario, row, column, input);
    this.record({ kind: 'scenarioInput', sheet: this.activeSheet, scenario, row, column, before, after: input });
  }

  /** Everything `scenario` types, on every sheet, as the workbook forks it. */
  overridesOf(scenario: string): { sheet: number; row: number; column: number; input: string; asText: boolean }[] {
    return this.pages.flatMap((page, sheet) =>
      page.scenarioInputs.of(scenario).map(entry => ({
        sheet,
        row: entry.row,
        column: entry.column,
        input: entry.input,
        asText: page.formats.formatAt(entry.row, entry.column).number.kind === 'text'
      }))
    );
  }

  /** Every sheet's scenario inputs, by sheet, which is what a file keeps. */
  scenarioInputsBySheet(): ScenarioInput[][] {
    return this.pages.map(page => page.scenarioInputs.all());
  }

  /** Puts every sheet's scenario inputs back, by sheet, for a load. */
  restoreScenarioInputs(bySheet: readonly (readonly ScenarioInput[])[]): void {
    this.pages.forEach((page, sheet) => page.scenarioInputs.restore(bySheet[sheet] ?? []));
  }

  /**
   * A new scenario, empty or typing what `from` types, at the end of the
   * list. Returns its id.
   */
  addScenario(name: string, from: string | null = null): string {
    const id = this.freeScenarioId();
    this.scenarios = [...this.scenarios, { id, name: this.freeScenarioName(name) }];
    if (from !== null) {
      for (const page of this.pages) {
        page.scenarioInputs.duplicate(from, id);
      }
    }
    return id;
  }

  renameScenario(id: string, name: string): void {
    const trimmed = name.trim();
    if (trimmed === '') {
      return;
    }
    this.scenarios = this.scenarios.map(entry =>
      entry.id === id ? { ...entry, name: this.freeScenarioName(trimmed, id) } : entry
    );
  }

  /**
   * Deletes a scenario and what it types. Its inputs leave the undo
   * stack with it: an undo of a cell it typed would be an undo into a
   * scenario that is not there.
   */
  deleteScenario(id: string): void {
    this.scenarios = this.scenarios.filter(entry => entry.id !== id);
    if (this.shownScenario === id) {
      this.showScenario(null);
    }
    for (const page of this.pages) {
      page.scenarioInputs.drop(id);
    }
    const keep = (step: Step): boolean => step.every(edit => edit.kind !== 'scenarioInput' || edit.scenario !== id);
    const undo = this.undoStack.filter(keep);
    const redo = this.redoStack.filter(keep);
    this.undoStack.length = 0;
    this.undoStack.push(...undo);
    this.redoStack.length = 0;
    this.redoStack.push(...redo);
  }

  private freeScenarioId(): string {
    let next = this.scenarios.length + 1;
    while (this.scenarios.some(entry => entry.id === `s${next}`)) {
      next++;
    }
    return `s${next}`;
  }

  /** A name no other scenario has, and never Base, which is what showing none of them is called. */
  private freeScenarioName(wanted: string, except = ''): string {
    const taken = (name: string): boolean =>
      name.toUpperCase() === 'BASE' ||
      this.scenarios.some(entry => entry.id !== except && entry.name.toUpperCase() === name.toUpperCase());
    if (!taken(wanted)) {
      return wanted;
    }
    for (let n = 2; ; n++) {
      if (!taken(`${wanted} ${n}`)) {
        return `${wanted} ${n}`;
      }
    }
  }

  /** Writes a cell's note, or takes it away with an empty one. One step of undo. */
  setNote(row: number, column: number, text: string): void {
    const before = this.page.notes.at(row, column);
    if (before === text) {
      return;
    }
    this.page.notes.set(row, column, text);
    this.record({ kind: 'note', sheet: this.activeSheet, row, column, before, after: text });
  }

  get conditional(): readonly ConditionalRule[] {
    return this.page.conditional;
  }

  get validations(): readonly Validation[] {
    return this.page.validations;
  }

  get hiddenRows(): Set<number> {
    return this.page.hiddenRows;
  }

  get filteredRows(): Set<number> {
    return this.page.filteredRows;
  }

  get filterColumn(): number {
    return this.page.filteredRows.size === 0 ? -1 : this.page.filterColumn;
  }

  set filterColumn(column: number) {
    this.page.filterColumn = column;
  }

  get rowHeights(): Map<number, number> {
    return this.page.rowHeights;
  }

  get fittedRows(): Map<number, number> {
    return this.page.fittedRows;
  }

  get columnWidths(): number[] {
    return this.page.columnWidths;
  }

  /**
   * The widths, and the charts over them with them.
   *
   * A chart sits over a cell, as it does in Excel: widening a column to
   * its left moves it right, and widening the one it starts in stretches
   * where in that column it starts. Every width change comes through
   * here — a drag, an autofit, a reset — so every one keeps the charts
   * beside the numbers they were placed beside. A shift, which moves
   * whole columns, writes the widths itself and moves the charts by the
   * columns; see `applyShift`.
   */
  set columnWidths(widths: number[]) {
    const before = this.page.columnWidths;
    if (this.page.charts.length > 0 && !sameWidths(before, widths)) {
      const was = columnAxis(before);
      const now = columnAxis(widths);
      this.page.charts = this.page.charts.map(chart => {
        const x = now.to(was.from(chart.place.x));
        return x === chart.place.x ? chart : { ...chart, place: { ...chart.place, x } };
      });
    }
    this.page.columnWidths = widths;
  }

  /**
   * The charts on this sheet, moved with a row or column insert or
   * delete: each keeps to the cell its corner was over, which moves
   * with the shift, or to where the deleted lines were.
   */
  private moveChartsWith(shift: Shift, widths: readonly number[], rows: RowsHeld): void {
    if (this.page.charts.length === 0) {
      return;
    }
    const was = shift.axis === 'row' ? rowAxis(rows) : columnAxis(widths);
    const now =
      shift.axis === 'row'
        ? rowAxis({ heights: [...this.rowHeights], fitted: [...this.fittedRows], hidden: [...this.hiddenRows], filtered: [...this.filteredRows] })
        : columnAxis(this.page.columnWidths);
    this.page.charts = this.page.charts.map(chart => {
      const at = was.from(shift.axis === 'row' ? chart.place.y : chart.place.x);
      const moved = shiftIndex(at.index, shift);
      const next = now.to(moved === -1 ? { index: shift.at, share: 0 } : { index: moved, share: at.share });
      const place = shift.axis === 'row' ? { ...chart.place, y: next } : { ...chart.place, x: next };
      return place.x === chart.place.x && place.y === chart.place.y ? chart : { ...chart, place };
    });
  }

  get frozenRows(): number {
    return this.page.frozenRows;
  }

  set frozenRows(rows: number) {
    this.page.frozenRows = rows;
  }

  get frozenColumns(): number {
    return this.page.frozenColumns;
  }

  get zoom(): number {
    return this.page.zoom;
  }

  set zoom(zoom: number) {
    this.page.zoom = zoom;
  }

  set frozenColumns(columns: number) {
    this.page.frozenColumns = columns;
  }

  get selection(): SheetSelection {
    return this.page.selection;
  }

  set selection(at: SheetSelection) {
    this.page.selection = at;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Commits a cell, recording what it replaced.
   *
   * A commit that changes nothing is not an edit: retyping the same
   * text, or pressing Enter on a cell without touching it, must not
   * put an entry on the stack that undoes to itself and looks broken.
   */
  setCell(row: number, column: number, input: string): string | null {
    const before = this.inputAt(row, column);
    if (before === input) {
      return null;
    }
    /**
     * A rule that refuses, refusing.
     *
     * Only what was *typed*, and only when it is not a formula: a
     * formula's value is not known until the recalculation has run,
     * and a commit that waited for it would be a keystroke that
     * blocked on the other thread. Those are marked after the fact
     * instead, which is where the marker in the window comes from.
     */
    const rule = this.validationAt(row, column);
    if (rule?.strict === true && !input.startsWith('=')) {
      const complaint = validate(rule.rule, literalOf(input));
      if (complaint !== null) {
        return rule.message ?? complaint;
      }
    }
    // One step, because typing a date is one action: it writes a
    // serial number and the format that makes the serial legible, and
    // undoing it has to take both back.
    const scenario = this.shownScenario;
    this.transact(() => {
      if (scenario === null) {
        this.writeCell(row, column, input);
        this.record({ kind: 'text', sheet: this.activeSheet, row, column, before, after: input });
      } else {
        // Typed into the scenario that is showing, and not the base.
        // Typing the base's own input back gives the cell back to it.
        this.setScenarioInput(scenario, row, column, input === this.sheet.input(row, column) ? null : input);
      }
      this.formatTypedDate(row, column, input);
      this.wrapTypedBreak(row, column, input);
    });
    return null;
  }

  /**
   * A line break typed into a cell turns wrap on, as Alt+Enter does in
   * Excel: a cell that does not wrap draws its text on one line, and
   * the break somebody typed would be there and not be seen. In the
   * same step as the text, so one undo takes both back.
   */
  private wrapTypedBreak(row: number, column: number, input: string): void {
    if (!input.includes('\n')) {
      return;
    }
    const format = this.formatAt(row, column);
    if (!format.paint.wrap) {
      this.setFormat(row, column, { ...format, paint: { ...format.paint, wrap: true } });
    }
  }

  /**
   * A typed date brings its format with it.
   *
   * The engine turns `2026-09-24` into 46,289 — see `literalValue` —
   * and on its own that is a cell showing forty-six thousand to
   * somebody who typed a date. The format is what finishes the job,
   * and it is applied here because the format axis is the document's
   * and the value is the sheet's.
   *
   * Only over `General`, and that is the whole of the rule. A cell
   * somebody has deliberately formatted has been answered already: a
   * date typed into a currency column shows as currency, which looks
   * odd and is what every spreadsheet does, because the alternative is
   * a format that silently undoes a choice somebody made on purpose.
   * The pattern is the one they typed in, so slashes give slashes
   * back.
   */
  private formatTypedDate(row: number, column: number, input: string): void {
    const current = this.formats.formatAt(row, column);
    if (current.number.kind !== 'general') {
      return;
    }
    const typed = parseTypedDate(input);
    if (typed === null) {
      return;
    }
    const number: NumberFormat =
      typed.date === null
        ? { kind: 'time', pattern: typed.time ?? 'hm' }
        : typed.time === null
          ? { kind: 'date', pattern: typed.date }
          : { kind: 'datetime', date: typed.date, time: typed.time };
    this.setFormat(row, column, { ...current, number });
  }

  /**
   * The write itself, with the cell's own format consulted.
   *
   * Separate from `setCell` because undo and redo write without
   * recording, and all three have to agree about what a Text-
   * formatted cell does with `007`.
   */
  private writeCell(row: number, column: number, input: string): void {
    this.sheet.setCell(row, column, input, this.formats.formatAt(row, column).number.kind === 'text');
  }

  /**
   * Formats a cell, recording the palette id it replaced.
   *
   * A format that changes nothing is not an edit, on the same rule as
   * a commit that changes nothing: formatting a bold cell bold again
   * must not put an entry on the stack that undoes to itself.
   *
   * Re-writing the cell afterwards is not busywork. A format can
   * change what its *value* is — Text is the one that does — so a
   * cell holding the number seven and formatted as Text has to become
   * a cell holding `7` as text, and the only way to get there is
   * through the same path that typing takes.
   */
  setFormat(row: number, column: number, format: CellFormat): void {
    const before = this.formats.idAt(row, column);
    const after = this.formats.idFor(format);
    if (before === after) {
      return;
    }
    this.applyFormat(row, column, after);
    this.record({ kind: 'format', sheet: this.activeSheet, row, column, before, after });
  }

  /**
   * Formats a whole row, column or sheet.
   *
   * The command layer decides when a selection is a region — see
   * `SheetService.format` — and this is what it calls. The saving is
   * not in the writing but in what is *kept*: one undo entry, one
   * number in the file, and one lookup per cell rather than an entry
   * per cell in all three.
   */
  formatRegion(scope: 'sheet' | 'row' | 'column', index: number, change: (format: CellFormat) => CellFormat): void {
    const heldId = scope === 'sheet' ? this.formats.sheetId : scope === 'row' ? this.formats.rowId(index) : this.formats.columnId(index);
    const after = this.formats.idFor(change(this.formats.byId(heldId)));

    // The change applies to the cells inside the region that had a
    // format of their own as well, so that making the sheet bold
    // keeps the currency somebody put in column C.
    const overrides = this.formats.overridesIn(scope, index).map(([key, before]) => ({
      key,
      before,
      after: this.formats.idFor(change(this.formats.byId(before)))
    }));

    if (after === heldId && overrides.every(entry => entry.after === entry.before)) {
      return;
    }
    this.writeRegion(scope, index, after, overrides.map(entry => ({ key: entry.key, id: entry.after })));
    this.record({
      kind: 'region',
      sheet: this.activeSheet,
      scope,
      index,
      row: scope === 'row' ? index : 0,
      column: scope === 'column' ? index : 0,
      before: heldId,
      after,
      overrides
    });
  }

  /** Writes a region and its overrides, and re-reads what it covers. */
  private writeRegion(
    scope: 'sheet' | 'row' | 'column',
    index: number,
    id: number,
    overrides: readonly { key: number; id: number }[]
  ): void {
    this.formats.setRegion(scope, index, id);
    for (const entry of overrides) {
      this.formats.setIdByKey(entry.key, entry.id);
    }
    this.retextRegion(scope, index);
  }

  /**
   * Re-reads every cell a region covers, in case its Text-ness moved.
   *
   * Bounded by the cells that hold something, not by the region: a
   * sheet-wide format walks the store, which is the number of cells
   * somebody has actually typed in.
   */
  private retextRegion(scope: 'sheet' | 'row' | 'column', index: number): void {
    for (const cell of this.sheet.entries()) {
      if (scope === 'row' && cell.row !== index) {
        continue;
      }
      if (scope === 'column' && cell.column !== index) {
        continue;
      }
      if (cell.input !== '') {
        this.writeCell(cell.row, cell.column, cell.input);
      }
    }
  }

  private applyFormat(row: number, column: number, id: number): void {
    const wasText = this.formats.formatAt(row, column).number.kind === 'text';
    this.formats.setId(row, column, id);
    const isText = this.formats.byId(id).number.kind === 'text';
    if (wasText !== isText) {
      const input = this.sheet.input(row, column);
      if (input !== '') {
        this.writeCell(row, column, input);
      }
    }
  }

  /** A cell's format, as the palette holds it. */
  formatAt(row: number, column: number): CellFormat {
    return this.formats.formatAt(row, column);
  }

  /**
   * What the screen shows for a cell: its value under its format.
   *
   * Moved off `Sheet` in Phase 9, because a display string now
   * depends on two things the sheet does not own both of. It stays on
   * *this* thread either way — the render worker receives the string
   * and never the format that made it, which is what keeps a
   * locale-aware number formatter off the frame path.
   */
  display(row: number, column: number): string {
    return formatWith(this.shown.value(row, column), this.formats.formatAt(row, column).number);
  }

  /** What a cell displays in the base, whatever is on screen. */
  baseDisplay(row: number, column: number): string {
    return formatWith(this.sheet.value(row, column), this.formats.formatAt(row, column).number);
  }

  /**
   * Runs several edits as one step.
   *
   * A paste, a cut and a fill are each one action to the person doing
   * them, so they are one entry on the stack however many cells they
   * touched. Nested calls join the step already open rather than
   * opening another, so a fill that pastes is still one press of
   * ctrl-Z.
   */
  transact(run: () => void, label?: string): void {
    if (this.collecting !== null) {
      // The outermost name wins: a sort that writes cells is a sort.
      this.collectingLabel ??= label ?? null;
      run();
      return;
    }
    const step: Edit[] = [];
    this.collecting = step;
    this.collectingLabel = label ?? null;
    try {
      run();
    } finally {
      this.collecting = null;
    }
    if (step.length > 0) {
      this.labels.set(step, this.collectingLabel ?? describeStep(step));
      this.undoStack.push(step);
      this.redoStack.length = 0;
    }
    this.collectingLabel = null;
  }

  /**
   * What Undo would take back, as the menu says it: `sort`, `typing in
   * B4`. Empty when there is nothing to undo.
   */
  get undoLabel(): string {
    return this.labelOf(this.undoStack[this.undoStack.length - 1]);
  }

  /** What Redo would put forward again, on the same terms. */
  get redoLabel(): string {
    return this.labelOf(this.redoStack[this.redoStack.length - 1]);
  }

  private labelOf(step: Step | undefined): string {
    return step === undefined ? '' : (this.labels.get(step) ?? describeStep(step));
  }

  private record(edit: Edit): void {
    if (this.collecting !== null) {
      this.collecting.push(edit);
      return;
    }
    this.undoStack.push([edit]);
    // A new edit is a new future; whatever was undone is unreachable.
    this.redoStack.length = 0;
  }

  undo(): boolean {
    const step = this.undoStack.pop();
    if (step === undefined) {
      return false;
    }
    // Backwards: two edits to one cell in one step have to be undone
    // in the order they were made or the earlier one wins.
    for (let at = step.length - 1; at >= 0; at--) {
      const edit = step[at];
      // The sheet the edit was made on, before anything is written:
      // `this.sheet`, `this.formats` and `this.merges` all mean the
      // active page, and undoing on the wrong one would write the old
      // text into the same address on whatever tab happened to show.
      this.activeSheet = edit.sheet;
      if (edit.kind === 'text') {
        this.writeCell(edit.row, edit.column, edit.before);
      } else if (edit.kind === 'region') {
        this.writeRegion(
          edit.scope,
          edit.index,
          edit.before,
          edit.overrides.map(entry => ({ key: entry.key, id: entry.before }))
        );
      } else if (edit.kind === 'structure') {
        this.undoShift(edit);
      } else if (edit.kind === 'note') {
        this.page.notes.set(edit.row, edit.column, edit.before);
      } else if (edit.kind === 'scenarioInput') {
        this.page.scenarioInputs.set(edit.scenario, edit.row, edit.column, edit.before);
      } else if (edit.kind === 'names') {
        this.restoreNames(edit.before);
      } else if (edit.kind === 'rules') {
        this.restoreRules(edit.before);
      } else if (edit.kind === 'charts') {
        this.restoreCharts(edit.before);
      } else {
        this.applyFormat(edit.row, edit.column, edit.before);
      }
    }
    this.redoStack.push(step);
    this.selectStep(step);
    return true;
  }

  redo(): boolean {
    const step = this.redoStack.pop();
    if (step === undefined) {
      return false;
    }
    for (const edit of step) {
      this.activeSheet = edit.sheet;
      if (edit.kind === 'text') {
        this.writeCell(edit.row, edit.column, edit.after);
      } else if (edit.kind === 'region') {
        this.writeRegion(
          edit.scope,
          edit.index,
          edit.after,
          edit.overrides.map(entry => ({ key: entry.key, id: entry.after }))
        );
      } else if (edit.kind === 'structure') {
        this.sheet.shift(edit.shift);
        this.formats.shift(edit.shift);
        this.page.columnWidths = shiftWidths(edit.widths, edit.shift);
        this.restoreRows(edit.rows, edit.shift);
        this.page.notes.restore(edit.notes);
        this.page.notes.shift(edit.shift);
        this.restoreScenarioInputs(edit.scenarioInputs);
        this.shiftScenarioInputs(edit.shift);
        this.restoreRanges(edit.ranges);
        this.shiftRanges(edit.shift);
        this.moveChartsWith(edit.shift, edit.widths, edit.rows);
      } else if (edit.kind === 'note') {
        this.page.notes.set(edit.row, edit.column, edit.after);
      } else if (edit.kind === 'scenarioInput') {
        this.page.scenarioInputs.set(edit.scenario, edit.row, edit.column, edit.after);
      } else if (edit.kind === 'names') {
        this.restoreNames(edit.after);
      } else if (edit.kind === 'rules') {
        this.restoreRules(edit.after);
      } else if (edit.kind === 'charts') {
        this.restoreCharts(edit.after);
      } else {
        this.applyFormat(edit.row, edit.column, edit.after);
      }
    }
    this.undoStack.push(step);
    this.selectStep(step);
    return true;
  }

  /**
   * Puts a structural change back.
   *
   * The opposite shift first, which restores every position, and then
   * the two things a shift destroys: the cells that were in a deleted
   * row, and the formulas it turned into `#REF!`. Both are written at
   * the positions they had before, which is where the opposite shift
   * has just put everything else.
   */
  private undoShift(edit: StructureEdit): void {
    this.sheet.shift({ ...edit.shift, by: -edit.shift.by });
    this.formats.shift({ ...edit.shift, by: -edit.shift.by });
    if (edit.formats !== null) {
      this.formats.restorePlacement(edit.formats);
    }
    // The charts come back from the step's own copy of them, below.
    this.page.columnWidths = [...edit.widths];
    this.restoreRows(edit.rows, null);
    this.page.notes.restore(edit.notes);
    this.restoreScenarioInputs(edit.scenarioInputs);
    this.restoreRanges(edit.ranges);
    for (const cell of edit.removed) {
      this.writeCell(cell.row, cell.column, cell.input);
    }
    for (const cell of edit.rewritten) {
      this.writeCell(cell.row, cell.column, cell.input);
    }
  }

  /**
   * The rows' heights and hidden sets as a structural step held them,
   * moved by `shift` for a redo or as they were for an undo.
   */
  private restoreRows(held: RowsHeld, shift: Shift | null): void {
    const move = shift === null || shift.axis !== 'row' ? null : shift;
    const heights = move === null ? held.heights : shiftHeights(held.heights, move);
    const fitted = move === null ? held.fitted : shiftHeights(held.fitted, move);
    this.rowHeights.clear();
    for (const [row, height] of heights) {
      this.rowHeights.set(row, height);
    }
    this.fittedRows.clear();
    for (const [row, height] of fitted) {
      this.fittedRows.set(row, height);
    }
    if (move === null) {
      replaceRows(this.hiddenRows, held.hidden);
      replaceRows(this.filteredRows, held.filtered);
    } else {
      shiftRows(this.hiddenRows, held.hidden, move);
      shiftRows(this.filteredRows, held.filtered, move);
    }
  }

  /** Takes the selection to what a step changed, so it is seen. */
  /**
   * Gives a range a name, or says why it cannot have one.
   *
   * The rules live in `Names`; this is the half that records the
   * change so it can be taken back, and tells the sheet to re-read
   * its formulas — a name changes what they *read*, not only what
   * they answer.
   */
  defineName(name: string, range: RangeRef): NameProblem | null {
    const before = this.sheet.names.all();
    const problem = this.sheet.names.define(name, range);
    if (problem !== null) {
      return problem;
    }
    this.record({
      kind: 'names',
      sheet: this.activeSheet,
      row: Math.min(range.start.row, range.end.row),
      column: Math.min(range.start.column, range.end.column),
      before,
      after: this.sheet.names.all()
    });
    this.sheet.namesChanged();
    return null;
  }

  /**
   * Gives a name a formula to hold — `=0.2`, or `=LAMBDA(x, x*2)` — or
   * says why it cannot. Recorded like a range's, and taken back the
   * same way; the step selects the cell somebody was on, because a
   * formula has no range of its own to go to.
   */
  defineFormulaName(name: string, formula: string): NameProblem | null {
    const before = this.sheet.names.all();
    const problem = this.sheet.names.defineFormula(name, formula);
    if (problem !== null) {
      return problem;
    }
    this.record({
      kind: 'names',
      sheet: this.activeSheet,
      row: this.selection.row,
      column: this.selection.column,
      before,
      after: this.sheet.names.all()
    });
    this.sheet.namesChanged();
    return null;
  }

  /**
   * Saves a name from the Names dialog: `was` is the name being edited,
   * empty for a new one, and `refersTo` is what it holds as written —
   * `=Sales!$A$4:$A$27` for a range, anything else for a formula.
   *
   * One step on the undo stack however much it did, so a rename is one
   * Ctrl+Z and not two. Formulas that read the old name are left
   * reading it: they say `#NAME?` until it is back, as in Excel.
   */
  saveName(was: string, name: string, refersTo: string): NameProblem | null {
    const problem = nameProblem(name);
    if (problem !== null) {
      return problem;
    }
    const text = refersTo.trim().replace(/^=/, '');
    let tree;
    try {
      tree = text === '' ? null : parseFormula(text);
    } catch (error) {
      if (!(error instanceof FormulaSyntaxError)) {
        throw error;
      }
      tree = null;
    }
    if (tree === null) {
      return 'formula';
    }
    const names = this.sheet.names;
    const before = names.all();
    if (was.trim() !== '' && was.trim().toUpperCase() !== name.trim().toUpperCase()) {
      names.remove(was);
    }
    const refused =
      tree.kind === 'ref'
        ? names.define(name, { start: tree.ref, end: tree.ref })
        : tree.kind === 'range'
          ? names.define(name, tree.range)
          : names.defineFormula(name, `=${text}`);
    if (refused !== null) {
      names.restore(before);
      return refused;
    }
    this.record({
      kind: 'names',
      sheet: this.activeSheet,
      row: this.selection.row,
      column: this.selection.column,
      before,
      after: names.all()
    });
    this.sheet.namesChanged();
    return null;
  }

  /** Takes a name away. False when there was no such name. */
  removeName(name: string): boolean {
    const entry = this.sheet.names.get(name);
    if (entry === undefined) {
      return false;
    }
    const before = this.sheet.names.all();
    this.sheet.names.remove(name);
    this.record({
      kind: 'names',
      sheet: this.activeSheet,
      row: isNamedRange(entry) ? Math.min(entry.range.start.row, entry.range.end.row) : this.selection.row,
      column: isNamedRange(entry) ? Math.min(entry.range.start.column, entry.range.end.column) : this.selection.column,
      before,
      after: this.sheet.names.all()
    });
    this.sheet.namesChanged();
    return true;
  }

  // ---------------------------------------------------------------------
  // Formats that think, and what a cell is allowed to hold
  // ---------------------------------------------------------------------

  /**
   * Adds a conditional format, as one step.
   *
   * The selection is left where it is: unlike a name or a cell edit,
   * a rule is about a range somebody has already chosen, and moving
   * the selection to announce it would take them away from what they
   * were looking at.
   */
  addConditional(rule: ConditionalRule): void {
    this.changeRules(page => page.conditional.push(rule), rule.range.start.row, rule.range.start.column);
  }

  removeConditional(at: number): boolean {
    if (this.page.conditional[at] === undefined) {
      return false;
    }
    const { row, column } = this.page.conditional[at].range.start;
    this.changeRules(page => page.conditional.splice(at, 1), row, column);
    return true;
  }

  /** A rule changed in place, keeping its place in the order they paint in. */
  replaceConditional(at: number, rule: ConditionalRule): boolean {
    if (this.page.conditional[at] === undefined) {
      return false;
    }
    this.changeRules(page => page.conditional.splice(at, 1, rule), rule.range.start.row, rule.range.start.column);
    return true;
  }

  replaceValidation(at: number, validation: Validation): boolean {
    if (this.page.validations[at] === undefined) {
      return false;
    }
    this.changeRules(
      page => page.validations.splice(at, 1, validation),
      validation.range.start.row,
      validation.range.start.column
    );
    return true;
  }

  addValidation(validation: Validation): void {
    this.changeRules(
      page => page.validations.push(validation),
      validation.range.start.row,
      validation.range.start.column
    );
  }

  removeValidation(at: number): boolean {
    if (this.page.validations[at] === undefined) {
      return false;
    }
    const { row, column } = this.page.validations[at].range.start;
    this.changeRules(page => page.validations.splice(at, 1), row, column);
    return true;
  }

  /** Every rule off this sheet, as one step. */
  clearRules(): void {
    if (this.page.conditional.length === 0 && this.page.validations.length === 0) {
      return;
    }
    this.changeRules(page => {
      page.conditional.length = 0;
      page.validations.length = 0;
    }, this.page.selection.row, this.page.selection.column);
  }

  // ---------------------------------------------------------------------
  // Charts
  // ---------------------------------------------------------------------

  get charts(): readonly Chart[] {
    return this.page.charts;
  }

  /**
   * The next id, counted across the whole workbook and never reused.
   *
   * Across the workbook rather than per sheet because a chart can be
   * carried to another one by a duplicate, and two charts with the
   * same id on one page would be one chart as far as selection and
   * dragging are concerned. Counted from what exists rather than kept
   * in a field, so a load does not have to restore a counter and a
   * file written by hand cannot produce a collision on the first
   * insert.
   */
  private nextChartId(): number {
    let highest = 0;
    for (const page of this.pages) {
      for (const chart of page.charts) {
        highest = Math.max(highest, chart.id);
      }
    }
    return highest + 1;
  }

  /** Adds a chart over a range and returns its id. */
  addChart(chart: Omit<Chart, 'id'>): number {
    const id = this.nextChartId();
    this.changeCharts(
      page => page.charts.push({ ...chart, id }),
      chart.range.start.row,
      chart.range.start.column
    );
    return id;
  }

  /**
   * Rewrites one chart, leaving the rest alone.
   *
   * Every change to a chart comes through here — a move, a resize, a
   * retitle, a different kind — so there is one place that records an
   * undo step and one shape of edit to take back.
   */
  changeChart(id: number, change: (chart: Chart) => Chart): boolean {
    const at = this.page.charts.findIndex(chart => chart.id === id);
    if (at < 0) {
      return false;
    }
    const chart = this.page.charts[at];
    const next = change(chart);
    this.changeCharts(page => (page.charts[at] = next), next.range.start.row, next.range.start.column);
    return true;
  }

  removeChart(id: number): boolean {
    const at = this.page.charts.findIndex(chart => chart.id === id);
    if (at < 0) {
      return false;
    }
    const { row, column } = this.page.charts[at].range.start;
    this.changeCharts(page => page.charts.splice(at, 1), row, column);
    return true;
  }

  chart(id: number): Chart | null {
    return this.page.charts.find(entry => entry.id === id) ?? null;
  }

  private changeCharts(change: (page: Page) => void, row: number, column: number): void {
    const page = this.page;
    const before = [...page.charts];
    change(page);
    this.record({
      kind: 'charts',
      sheet: this.activeSheet,
      row,
      column,
      before,
      after: [...page.charts]
    });
  }

  private restoreCharts(held: readonly Chart[]): void {
    const page = this.page;
    page.charts.length = 0;
    page.charts.push(...held);
  }

  /** The validation over a cell, or null. The first one wins. */
  validationAt(row: number, column: number): Validation | null {
    for (const validation of this.page.validations) {
      if (coversCell(validation.range, row, column)) {
        return validation;
      }
    }
    return null;
  }

  /** The conditional rule on top at a cell: the last one covering it, since later rules paint over earlier ones. */
  conditionalAt(row: number, column: number): ConditionalRule | null {
    for (let at = this.page.conditional.length - 1; at >= 0; at--) {
      if (coversCell(this.page.conditional[at].range, row, column)) {
        return this.page.conditional[at];
      }
    }
    return null;
  }

  /** Every conditional rule covering a cell, as indexes into `conditional`, in the order they paint. */
  conditionalsAt(row: number, column: number): number[] {
    const at: number[] = [];
    this.page.conditional.forEach((rule, index) => {
      if (coversCell(rule.range, row, column)) {
        at.push(index);
      }
    });
    return at;
  }

  private changeRules(change: (page: Page) => void, row: number, column: number): void {
    const page = this.page;
    const before = {
      conditional: [...page.conditional],
      validations: [...page.validations]
    };
    change(page);
    this.record({
      kind: 'rules',
      sheet: this.activeSheet,
      row,
      column,
      before,
      after: { conditional: [...page.conditional], validations: [...page.validations] }
    });
  }

  private restoreRules(held: {
    readonly conditional: readonly ConditionalRule[];
    readonly validations: readonly Validation[];
  }): void {
    const page = this.page;
    page.conditional.length = 0;
    page.conditional.push(...held.conditional);
    page.validations.length = 0;
    page.validations.push(...held.validations);
  }

  private restoreNames(entries: readonly DefinedName[]): void {
    this.sheet.names.restore(entries);
    this.sheet.namesChanged();
  }

  private selectStep(step: Step): void {
    const edit = step[0];
    if (edit !== undefined) {
      this.activeSheet = edit.sheet;
    }
    let firstRow = Number.POSITIVE_INFINITY;
    let lastRow = Number.NEGATIVE_INFINITY;
    let firstColumn = Number.POSITIVE_INFINITY;
    let lastColumn = Number.NEGATIVE_INFINITY;
    for (const edit of step) {
      firstRow = Math.min(firstRow, edit.row);
      lastRow = Math.max(lastRow, edit.row);
      firstColumn = Math.min(firstColumn, edit.column);
      lastColumn = Math.max(lastColumn, edit.column);
    }
    this.selection = { row: firstRow, column: firstColumn, anchorRow: lastRow, anchorColumn: lastColumn };
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number, cornerRow?: number, cornerColumn?: number): void {
    // A corner that is the active cell is not written down, so a
    // selection made the old way is the same object shape it always was.
    const corner = cornerRow === undefined || cornerColumn === undefined || (cornerRow === row && cornerColumn === column);
    this.selection = corner ? { row, column, anchorRow, anchorColumn } : { row, column, anchorRow, anchorColumn, cornerRow, cornerColumn };
  }

  /**
   * Inserts or deletes rows or columns, as one step.
   *
   * The whole of it is one entry on the undo stack, because it is one
   * action to the person doing it however far it reached — the same
   * rule `transact` exists for.
   *
   * The column widths move with the columns they describe: inserting
   * a column in front of a wide one and leaving the widths alone
   * makes the wrong column wide.
   */
  applyShift(shift: Shift): void {
    const removed: { row: number; column: number; input: string }[] = [];
    const rewritten: { row: number; column: number; input: string }[] = [];
    for (const cell of this.sheet.entries()) {
      const index = shift.axis === 'row' ? cell.row : cell.column;
      if (shiftIndex(index, shift) === -1) {
        removed.push({ ...cell });
      } else if (cell.input.startsWith('=')) {
        // Kept whether or not it changes. Deciding here would mean
        // shifting every formula twice, and the list is dropped by
        // `record` if the step turns out to be empty anyway.
        rewritten.push({ ...cell });
      }
    }

    const widths = [...this.columnWidths];
    const hidden = [...this.hiddenRows];
    const filtered = [...this.filteredRows];
    const rows: RowsHeld = {
      heights: [...this.rowHeights],
      fitted: [...this.fittedRows],
      hidden,
      filtered
    };
    const notes = this.page.notes.all();
    const scenarioInputs = this.scenarioInputsBySheet();
    const formats = shift.by < 0 ? this.formats.placement() : null;
    const ranges = this.pages.map(page => ({
      conditional: [...page.conditional],
      validations: [...page.validations],
      charts: [...page.charts]
    }));
    this.shiftRanges(shift);
    this.sheet.shift(shift);
    this.formats.shift(shift);
    this.merges.shift(shift);
    this.page.notes.shift(shift);
    this.shiftScenarioInputs(shift);
    // Written past the setter: the columns moved, and the charts move
    // with the columns below, not with the widths.
    this.page.columnWidths = shiftWidths(this.columnWidths, shift);
    // A hidden row is hidden by index, so it moves with the rows it
    // was among — an insert above a hidden row must not reveal it and
    // hide its neighbour instead.
    if (shift.axis === 'row') {
      this.restoreRows(rows, shift);
    }
    this.moveChartsWith(shift, widths, rows);

    this.record({
      kind: 'structure',
      sheet: this.activeSheet,
      row: shift.axis === 'row' ? shift.at : 0,
      column: shift.axis === 'column' ? shift.at : 0,
      shift,
      removed,
      rewritten,
      widths,
      rows,
      notes,
      scenarioInputs,
      formats,
      ranges
    });
  }

  /** Moves every sheet's scenario inputs by a shift of the active sheet; see `ScenarioInputs.shift`. */
  private shiftScenarioInputs(shift: Shift): void {
    const named = { ...shift, sheet: shift.sheet ?? this.book.nameOf(this.activeSheet) };
    this.pages.forEach((page, sheet) => page.scenarioInputs.shift(named, this.book.nameOf(sheet)));
  }

  // ---------------------------------------------------------------------
  // The sheets
  // ---------------------------------------------------------------------

  get sheetCount(): number {
    return this.pages.length;
  }

  /** Every sheet's name and colour, in tab order. */
  sheets(): { name: string; colour: string | null }[] {
    return this.pages.map(page => ({ name: page.sheet.name, colour: page.sheet.colour }));
  }

  pageAt(index: number): Page | undefined {
    return this.pages[index];
  }

  /** Shows a sheet. Not an edit: looking at something is not a change. */
  activate(index: number): boolean {
    if (this.pages[index] === undefined || index === this.activeSheet) {
      return false;
    }
    this.activeSheet = index;
    return true;
  }

  /**
   * Adds a sheet at the end and shows it.
   *
   * **Not undoable, and nor are the other four.** An undo entry for a
   * deleted sheet would have to carry every cell, format, merge,
   * width and freeze on it, and — worse — the entries already on the
   * stack name their sheet by index, which removing or moving one
   * renumbers. So the stack is dropped instead, which is what Excel
   * does with a sheet delete and for the same reason. Saying it is
   * gone is better than a ctrl-Z that puts text back on the wrong
   * tab.
   */
  addSheet(name = `Sheet${this.pages.length + 1}`): number {
    const index = this.book.addSheet(name);
    this.pages.push(newPage(this.book.sheet(index)));
    this.forgetHistory();
    this.activeSheet = index;
    return index;
  }

  /**
   * The rules and charts, moved with the cells a shift moved.
   *
   * The active sheet's conditional formats and validations move with
   * its cells; one whose every row or column was deleted goes with them,
   * as a merge does. A chart moves when its range is on the sheet that
   * changed shape — its own, or another it names — and keeps its place
   * over nothing when all its cells are deleted, as a formula keeps
   * `#REF!`: a chart that vanished with its data would be worse.
   */
  private shiftRanges(shift: Shift): void {
    const name = this.sheet.name.toUpperCase();
    const gone = (range: RangeRef) =>
      shift.axis === 'row'
        ? Math.min(range.start.row, range.end.row) >= MAX_ROWS_HERE
        : Math.min(range.start.column, range.end.column) >= MAX_COLUMNS_HERE;
    const page = this.page;
    const conditional = page.conditional.map(rule => ({ ...rule, range: shiftRange(rule.range, shift) })).filter(rule => !gone(rule.range));
    page.conditional.length = 0;
    page.conditional.push(...conditional);
    const validations = page.validations.map(rule => ({ ...rule, range: shiftRange(rule.range, shift) })).filter(rule => !gone(rule.range));
    page.validations.length = 0;
    page.validations.push(...validations);
    for (const other of this.pages) {
      for (const [at, chart] of other.charts.entries()) {
        const named = chart.range.start.sheet;
        const reads = named === undefined ? other === page : named.toUpperCase() === name;
        if (reads) {
          other.charts[at] = { ...chart, range: shiftRange(chart.range, shift) };
        }
      }
    }
  }

  private restoreRanges(held: readonly PageRanges[]): void {
    held.forEach((ranges, index) => {
      const page = this.pages[index];
      if (page === undefined) {
        return;
      }
      page.conditional.length = 0;
      page.conditional.push(...ranges.conditional);
      page.validations.length = 0;
      page.validations.push(...ranges.validations);
      page.charts.length = 0;
      page.charts.push(...ranges.charts);
    });
  }

  renameSheet(index: number, to: string): boolean {
    if (this.pages[index] === undefined) {
      return false;
    }
    // A chart that reads this sheet from another names it, as a formula
    // does, and follows the rename the formulas do.
    const from = this.book.nameOf(index);
    for (const page of this.pages) {
      for (const [at, chart] of page.charts.entries()) {
        if (chart.range.start.sheet !== undefined && chart.range.start.sheet.toUpperCase() === from.toUpperCase()) {
          page.charts[at] = { ...chart, range: { ...chart.range, start: { ...chart.range.start, sheet: to } } };
        }
      }
    }
    // The name is in the text of every formula that reads across, so
    // the rename rewrites them — and the undo stack holds the text
    // they had before, under the old name.
    this.book.renameSheet(index, to);
    this.forgetHistory();
    return true;
  }

  removeSheet(index: number): boolean {
    if (!this.book.removeSheet(index)) {
      return false;
    }
    this.pages.splice(index, 1);
    this.repage();
    this.activeSheet = Math.min(this.activeSheet, this.pages.length - 1);
    this.forgetHistory();
    return true;
  }

  moveSheet(from: number, to: number): boolean {
    if (!this.book.moveSheet(from, to)) {
      return false;
    }
    const [page] = this.pages.splice(from, 1);
    this.pages.splice(to, 0, page);
    this.repage();
    this.activeSheet = to;
    this.forgetHistory();
    return true;
  }

  /**
   * Copies a sheet, its cells and everything drawn over them.
   *
   * The cells are the workbook's to copy; the formats, merges,
   * widths, freezes and hidden rows are this layer's, and a duplicate
   * that brought the cells without them would be a copy that did not
   * look like the thing it copied.
   */
  duplicateSheet(index: number): number {
    const from = this.pages[index];
    if (from === undefined) {
      return -1;
    }
    const at = this.book.duplicateSheet(index);
    this.pages.push({
      ...newPage(this.book.sheet(at)),
      formats: from.formats.copy(),
      merges: from.merges.copy(),
      notes: from.notes.copy(),
      scenarioInputs: from.scenarioInputs.copy(),
      columnWidths: [...from.columnWidths],
      hiddenRows: new Set(from.hiddenRows),
      rowHeights: new Map(from.rowHeights),
      fittedRows: new Map(from.fittedRows),
      filteredRows: new Set(from.filteredRows),
      filterColumn: from.filterColumn,
      frozenRows: from.frozenRows,
      frozenColumns: from.frozenColumns,
      zoom: from.zoom,
      // Copied with fresh ids, because an id is unique across the
      // workbook and two charts sharing one would be one chart as far
      // as selecting and dragging are concerned. The range inside
      // each is unqualified, so it points at the copy's own cells —
      // which is what duplicating a sheet with a chart on it means.
      charts: from.charts.map((chart, offset) => ({ ...chart, id: this.nextChartId() + offset }))
    });
    this.forgetHistory();
    this.activeSheet = at;
    return at;
  }

  setSheetColour(index: number, colour: string | null): boolean {
    const page = this.pages[index];
    if (page === undefined) {
      return false;
    }
    page.sheet.colour = colour;
    return true;
  }

  /**
   * Shapes the workbook to a list of names, for a load.
   *
   * Called before anything is written into the pages, which is the
   * order that matters twice: every write below goes through the
   * active page and there has to be one, and a formula reading
   * `Data!A1` can only find `Data` if `Data` exists by the time it is
   * parsed.
   *
   * Renaming before there are cells is also what keeps this from
   * rewriting anything: `Workbook.renameSheet` walks the formulas,
   * and at this point there are none.
   */
  restoreSheets(names: readonly string[]): void {
    const wanted = names.length === 0 ? ['Sheet1'] : names;
    while (this.pages.length > wanted.length) {
      this.book.removeSheet(this.pages.length - 1);
      this.pages.pop();
    }
    for (let index = 0; index < wanted.length; index++) {
      if (index < this.pages.length) {
        this.book.renameSheet(index, wanted[index]);
        this.pages[index] = newPage(this.book.sheet(index));
      } else {
        const at = this.book.addSheet(wanted[index]);
        this.pages.push(newPage(this.book.sheet(at)));
      }
    }
    this.repage();
    this.activeSheet = 0;
    this.forgetHistory();
  }

  /** Re-binds each page to the sheet now at its index. */
  private repage(): void {
    for (let index = 0; index < this.pages.length; index++) {
      this.pages[index] = { ...this.pages[index], sheet: this.book.sheet(index) };
    }
  }

  /**
   * Empties both stacks.
   *
   * Public for the one caller outside this file, an import: it adds a
   * sheet, which forgets the history anyway, and then writes the
   * sheet's contents, which would otherwise leave one entry that
   * undoes a file into an empty tab.
   */
  forgetHistory(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  /** Every non-empty cell of the active sheet, for a repository or a search. */
  *entries(): Generator<{ row: number; column: number; input: string }> {
    yield* this.sheet.entries();
  }

  /** What the formula bar shows for the active cell. */
  get activeInput(): string {
    return this.inputAt(this.selection.row, this.selection.column);
  }
}

/**
 * Column widths, moved by a column insert or delete.
 *
 * The array stays the same length — the sheet is as wide as it was —
 * so an insert pushes widths off the end and a delete pulls the last
 * one along at it. An inserted column takes the width of the one it
 * pushed aside, which is the only answer that does not make a
 * carefully-widened column suddenly narrow while its neighbour is
 * wide. It is the same arithmetic the cells get, on an array instead
 * of a map.
 */
/** Whether a cell is inside a range, corners in any order. */
function coversCell(range: RangeRef, row: number, column: number): boolean {
  const firstRow = Math.min(range.start.row, range.end.row);
  const lastRow = Math.max(range.start.row, range.end.row);
  const firstColumn = Math.min(range.start.column, range.end.column);
  const lastColumn = Math.max(range.start.column, range.end.column);
  return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
}

function shiftWidths(widths: readonly number[], shift: Shift): number[] {
  if (shift.axis !== 'column') {
    return [...widths];
  }
  const next = [...widths];
  const removed = -shift.by;
  if (shift.by > 0) {
    for (let column = widths.length - 1; column >= shift.at; column--) {
      const from = column - shift.by;
      next[column] = from >= shift.at ? widths[from] : widths[shift.at];
    }
  } else {
    for (let column = shift.at; column < widths.length; column++) {
      next[column] = widths[column + removed] ?? widths[widths.length - 1];
    }
  }
  return next;
}

/**
 * A set of row indices, moved by a shift.
 *
 * Hiding is by index, so an insert above a hidden row must not reveal
 * it and hide its neighbour instead.
 */
function replaceRows(rows: Set<number>, held: readonly number[]): void {
  rows.clear();
  for (const row of held) {
    rows.add(row);
  }
}

/** Row heights, by index, moved by a shift; a deleted row's height goes with it. */
function shiftHeights(heights: readonly (readonly [number, number])[], shift: Shift): [number, number][] {
  const moved: [number, number][] = [];
  for (const [row, height] of heights) {
    const at = shiftIndex(row, shift);
    if (at !== -1) {
      moved.push([at, height]);
    }
  }
  return moved;
}

function shiftRows(rows: Set<number>, held: readonly number[], shift: Shift): void {
  rows.clear();
  for (const row of held) {
    const moved = shiftIndex(row, shift);
    if (moved !== -1) {
      rows.add(moved);
    }
  }
}

/**
 * What a step nobody named is called, from the edits in it.
 *
 * Typing is the common case and the one worth being exact about: one
 * cell's text, with whatever format a typed date brought along, is
 * "typing in B4". Everything else is named by the kind of thing it
 * changed, which is what the commands that do not name their own steps
 * have in common.
 */
function describeStep(step: Step): string {
  const texts = step.filter(edit => edit.kind === 'text');
  if (texts.length === 1) {
    const typed = texts[0];
    const alongside = step.every(
      edit => edit.kind === 'text' || (edit.kind === 'format' && edit.row === typed.row && edit.column === typed.column)
    );
    if (alongside) {
      return `typing in ${columnName(typed.column)}${typed.row + 1}`;
    }
  }
  const kinds = new Set(step.map(edit => edit.kind));
  if (kinds.size === 1) {
    const only = step[0];
    switch (only.kind) {
      case 'text':
        return `changes to ${step.length} cells`;
      case 'format':
      case 'region':
        return 'formatting';
      case 'structure': {
        const { axis, by } = only.shift;
        const count = Math.abs(by);
        const noun = count === 1 ? axis : `${count} ${axis}s`;
        return `${by > 0 ? 'insert' : 'delete'} ${noun}`;
      }
      case 'names':
        return 'naming';
      case 'rules':
        return 'rules';
      case 'charts':
        return 'chart';
      case 'note':
        return 'note';
    }
  }
  if ([...kinds].every(kind => kind === 'format' || kind === 'region')) {
    return 'formatting';
  }
  return 'changes';
}

/**
 * One axis of the grid, as the pixels a chart is placed in: where a
 * position falls (which line, and how far into it) and back.
 */
interface Axis {
  from(pixels: number): { index: number; share: number };
  to(at: { index: number; share: number }): number;
}

/** Walks an axis of lines, each of the size `size` gives it. */
function axisOf(size: (index: number) => number): Axis {
  return {
    from(pixels) {
      let start = 0;
      for (let index = 0; index < 1_000_000; index++) {
        const width = size(index);
        if (pixels < start + width) {
          return { index, share: width > 0 ? (pixels - start) / width : 0 };
        }
        start += width;
      }
      return { index: 0, share: 0 };
    },
    to({ index, share }) {
      let start = 0;
      for (let at = 0; at < index; at++) {
        start += size(at);
      }
      return Math.round(start + share * size(index));
    }
  };
}

function columnAxis(widths: readonly number[]): Axis {
  return axisOf(index => widths[index] ?? COLUMN_WIDTH);
}

function rowAxis(rows: RowsHeld): Axis {
  const heights = new Map(rows.fitted);
  for (const [row, height] of rows.heights) {
    heights.set(row, height);
  }
  const hidden = new Set([...rows.hidden, ...rows.filtered]);
  return axisOf(index => (hidden.has(index) ? 0 : (heights.get(index) ?? ROW_HEIGHT)));
}

function sameWidths(a: readonly number[], b: readonly number[]): boolean {
  if (a === b) {
    return true;
  }
  const length = Math.max(a.length, b.length);
  for (let index = 0; index < length; index++) {
    if ((a[index] ?? COLUMN_WIDTH) !== (b[index] ?? COLUMN_WIDTH)) {
      return false;
    }
  }
  return true;
}
