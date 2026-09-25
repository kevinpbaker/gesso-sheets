import {
  columnKey,
  columnOf,
  keyOn,
  MAX_SHEETS,
  quoteSheetName,
  rangeKeys,
  rowOf,
  sheetOf,
  type CellRef,
  type RangeRef
} from './A1';
import { bareWordsOf, callNamesOf, referencesOf, type Ast } from './Ast';
import { parseTypedDate } from './Dates';
import { DependencyGraph } from './DependencyGraph';
import { Names } from './Names';
import { evaluate, type WorkbookContext } from './Evaluator';
import { nowSerial, VOLATILE, type FunctionContext } from './Functions';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { Sheet } from './Sheet';
import { shiftFormula, shiftIndex, type Shift } from './Shift';
import { tokenize } from './Tokenizer';
import { CIRC, formatValue, VALUE, type CellValue } from './Values';

interface Cell {
  /** Exactly what was typed, which is what an editor puts back. */
  readonly input: string;
  /** The parsed formula, or null for a literal. */
  readonly formula: Ast | null;
  value: CellValue;
}

/** A sheet as the workbook knows it: a name, a colour, and how far down it goes. */
interface SheetEntry {
  name: string;
  colour: string | null;
  /**
   * How far down anything has ever been written on this sheet.
   *
   * A high-water mark, and deliberately not recomputed when cells are
   * cleared: shrinking it would mean walking the store on every
   * delete, and the only cost of it being too large is that a
   * whole-column reference reads a few cells that are empty. Too
   * small would drop data, which is why it never goes down.
   */
  usedRows: number;
}

export interface RecalcResult {
  /** Cells evaluated by this call. */
  readonly evaluated: number;
  /** False when a budget ran out with cells still to do. */
  readonly done: boolean;
}

/**
 * Many sheets, one store, one dependency graph.
 *
 * The third axis is a **multiplier on the key**, not a map of maps.
 * A cell's identity was `row * MAX_COLUMNS + column`; it is now that
 * plus `sheet * SHEET_STRIDE`, which keeps every edge of the graph a
 * single unboxed integer. A `Map<sheet, Map<key, Cell>>` would cost
 * an extra lookup on every edge, and a recalculation does a lookup
 * per edge.
 *
 * **One graph for the workbook, and this is the whole reason the
 * class exists.** A formula on Sheet 1 that reads Sheet 2 has to be
 * recalculated in an order that accounts for both, and there is no
 * such order to be had from two graphs that each know half the
 * edges. One graph, one dirty set, one plan, one budget — and a
 * recalculation that crosses sheets is then the same recalculation it
 * always was, over keys that happen to carry a sheet in them.
 *
 * What is *not* shared is the store's shape: each sheet has its own
 * used-rows mark, its own cells, and its own rows and columns to
 * insert into. A shift on Sheet 2 moves references to Sheet 2 and
 * nothing else; see `Shift.sheet`.
 *
 * Publishing is not this file's problem and deliberately so. The
 * claim Phase 13 has to defend — that a formula depending on 50,000
 * cells on a sheet nobody is looking at publishes nothing — is about
 * the window the application worker sends, and this engine goes on
 * knowing nothing about viewports.
 */
export class Workbook {
  /** Cumulative counts, in the manner of Gesso's `LayoutEngine.stats`. */
  readonly stats = { evaluated: 0, planned: 0, plans: 0 };

  /**
   * The names the workbook knows, and the reason they live here.
   *
   * A name is a reference by another spelling, so everything that
   * reads references has to be able to read names: evaluation, the
   * dependency graph, and the shift that moves both. Holding them
   * anywhere else would mean handing them to all three.
   *
   * Workbook-wide rather than per sheet, because that is what a name
   * is for: `=SUM(Sales)` written on any sheet means the same cells.
   * The range carries the sheet it was taken from.
   */
  readonly names = new Names();

  private readonly sheets: SheetEntry[] = [];
  private readonly facades: Sheet[] = [];

  private readonly cells = new Map<number, Cell>();
  private readonly graph = new DependencyGraph();
  /** Cells whose value is out of date. */
  private readonly dirty = new Set<number>();
  /** The order the current slice is working through, and where it is. */
  private plan: { order: number[]; at: number } | null = null;

  /**
   * The clock and the dice the volatile functions read.
   *
   * Fields rather than direct calls to `Date` and `Math.random` so
   * that a spec can say what day it is, which is the only way to
   * assert anything about `TODAY()` at all.
   */
  clock: () => number = nowSerial;
  dice: () => number = Math.random;

  /**
   * Formulas that answer differently without anything changing:
   * `RAND`, `RANDBETWEEN`, `NOW`, `TODAY`.
   */
  private readonly volatile = new Set<number>();

  /**
   * Formulas whose precedents cannot be read off their own text.
   *
   * `=INDIRECT("A" & B1)` reads a cell that the formula never names,
   * so the graph built from the tree has an edge to B1 and none to
   * the cell it actually fetched.
   */
  private readonly dynamic = new Set<number>();

  /** The keys the cell now being evaluated read, when it is dynamic. */
  private recording: Set<number> | null = null;

  /** Dynamic cells already sent round a second time since the last edit. */
  private readonly redone = new Set<number>();

  /** Fixed for a whole recalculation, so two `NOW()`s agree. */
  private moment: FunctionContext | null = null;

  constructor(names: readonly string[] = ['Sheet1']) {
    for (const name of names) {
      this.addSheet(name);
    }
  }

  // ---------------------------------------------------------------------
  // The sheets
  // ---------------------------------------------------------------------

  get sheetCount(): number {
    return this.sheets.length;
  }

  /** Every sheet's name, in the order the tabs are in. */
  sheetNames(): string[] {
    return this.sheets.map(entry => entry.name);
  }

  nameOf(sheet: number): string {
    return this.sheets[sheet]?.name ?? '';
  }

  colourOf(sheet: number): string | null {
    return this.sheets[sheet]?.colour ?? null;
  }

  setColour(sheet: number, colour: string | null): void {
    const entry = this.sheets[sheet];
    if (entry !== undefined) {
      entry.colour = colour;
    }
  }

  /**
   * The index of a sheet by name, case-insensitively.
   *
   * The lookup a reference goes through, so it is by name: the text
   * holds `Sheet2!` and the key space holds an index, and this is the
   * only place the two meet.
   */
  sheetFor(name: string): number | null {
    const wanted = name.trim().toUpperCase();
    const at = this.sheets.findIndex(entry => entry.name.toUpperCase() === wanted);
    return at === -1 ? null : at;
  }

  /** The facade for a sheet, which is what the rest of the application holds. */
  sheet(index: number): Sheet {
    return this.facades[index];
  }

  /**
   * Adds a sheet, and wakes the formulas that were waiting for it.
   *
   * A formula reading `Sheet4!A1` before Sheet 4 exists holds `#REF!`
   * — it has nowhere to point. Creating the sheet gives it somewhere,
   * and nothing in the graph would have noticed, because the edge
   * that would have noticed is the one that could not be built. So
   * the wiring is redone, which is the same blunt answer
   * `namesChanged` gives for the same reason.
   */
  addSheet(name: string): number {
    if (this.sheets.length >= MAX_SHEETS) {
      throw new Error(`a workbook holds at most ${MAX_SHEETS} sheets`);
    }
    const index = this.sheets.length;
    this.sheets.push({ name: this.freeName(name), colour: null, usedRows: 0 });
    this.facades.push(new Sheet(this, index));
    // A sheet always arrives at the end and is placed with
    // `moveSheet`, because adding one in the middle is a renumber of
    // every key after it and `moveSheet` is where that is written.
    this.rewireAll();
    return index;
  }

  /**
   * A name nothing else has, by adding a number until it is one.
   *
   * Two sheets called `Sheet2` would make `Sheet2!A1` mean whichever
   * the lookup reached first, which is a formula whose meaning
   * depends on tab order.
   *
   * `except` is the sheet being renamed, which does not count as
   * something else. Without it, renaming `Working` to `Working` —
   * which is what fixing a capital or pressing Enter on an unchanged
   * box does — found the sheet itself in the way and called it
   * `Working 2`.
   */
  private freeName(wanted: string, except = -1): string {
    const taken = (name: string): boolean => {
      const at = this.sheetFor(name);
      return at !== null && at !== except;
    };
    const trimmed = wanted.trim() === '' ? 'Sheet' : wanted.trim();
    if (!taken(trimmed)) {
      return trimmed;
    }
    for (let suffix = 2; ; suffix++) {
      const candidate = `${trimmed} ${suffix}`;
      if (!taken(candidate)) {
        return candidate;
      }
    }
  }

  /**
   * Renames a sheet, and rewrites the formulas that named it.
   *
   * The name is in the text of every formula that reads across, so a
   * rename is a rewrite — Phase 10's machinery aimed down the third
   * axis. Doing it the other way, by holding an id in the tree, would
   * make this free and would make `Cell.input` disagree with what the
   * editor puts back, because nothing regenerates what somebody
   * typed.
   *
   * Returns how many formulas were rewritten.
   */
  renameSheet(sheet: number, to: string): number {
    const entry = this.sheets[sheet];
    if (entry === undefined) {
      return 0;
    }
    const from = entry.name;
    const settled = this.freeName(to, sheet);
    if (settled.toUpperCase() === from.toUpperCase()) {
      entry.name = settled;
      return 0;
    }
    entry.name = settled;

    let rewrites = 0;
    for (const [key, cell] of [...this.cells]) {
      if (cell.formula === null) {
        continue;
      }
      const input = renameInFormula(cell.input, from, settled);
      if (input === cell.input) {
        continue;
      }
      rewrites++;
      this.writeFormula(key, input);
    }
    this.rewireAll();
    return rewrites;
  }

  /**
   * Removes a sheet, and leaves `#REF!` behind wherever it was read.
   *
   * The cells go, the sheets after it shift down, and every formula
   * in the workbook is re-read: one that named this sheet now names
   * nothing, which `keyFor` answers with `#REF!` — the same answer a
   * deleted row gives, by the same route.
   *
   * The last sheet cannot be removed. A workbook of no sheets is a
   * state nothing else in the application is written to survive, and
   * the tab strip has nothing to show.
   */
  removeSheet(sheet: number): boolean {
    if (this.sheets.length <= 1 || this.sheets[sheet] === undefined) {
      return false;
    }
    this.sheets.splice(sheet, 1);
    this.recompact(index => (index === sheet ? -1 : index > sheet ? index - 1 : index));
    return true;
  }

  /** Moves a sheet along the strip, carrying its cells with it. */
  moveSheet(from: number, to: number): boolean {
    if (this.sheets[from] === undefined || to < 0 || to >= this.sheets.length || from === to) {
      return false;
    }
    const [entry] = this.sheets.splice(from, 1);
    this.sheets.splice(to, 0, entry);
    this.recompact(index => {
      if (index === from) {
        return to;
      }
      if (from < to) {
        return index > from && index <= to ? index - 1 : index;
      }
      return index >= to && index < from ? index + 1 : index;
    });
    return true;
  }

  /**
   * Copies a sheet and everything on it, under a free name.
   *
   * The formulas are copied as text, so an unqualified reference goes
   * on meaning "my own sheet" and lands on the copy — which is what
   * duplicating a sheet is for. A reference that named a sheet keeps
   * naming it.
   */
  duplicateSheet(sheet: number, name?: string): number {
    const entry = this.sheets[sheet];
    if (entry === undefined) {
      return -1;
    }
    const copies: { row: number; column: number; input: string }[] = [];
    for (const [key, cell] of this.cells) {
      if (sheetOf(key) === sheet) {
        copies.push({ row: rowOf(key), column: columnOf(key), input: cell.input });
      }
    }
    const index = this.addSheet(name ?? `${entry.name} copy`);
    this.sheets[index].colour = entry.colour;
    for (const copy of copies) {
      this.setCell(index, copy.row, copy.column, copy.input);
    }
    return index;
  }

  /**
   * Rebuilds the store after the sheets have been renumbered.
   *
   * Every key holds its sheet, so moving or removing a sheet moves
   * every key on it. The store is rebuilt rather than patched, for
   * the reason `shift` rebuilds rather than patches: a cell moved to
   * a key another cell has not vacated yet is how an in-place version
   * corrupts itself.
   */
  private recompact(whereNow: (sheet: number) => number): void {
    const carried = new Map<number, Cell>();
    for (const [key, cell] of this.cells) {
      const moved = whereNow(sheetOf(key));
      if (moved === -1) {
        continue;
      }
      carried.set(keyOn(moved, rowOf(key), columnOf(key)), cell);
    }
    this.facades.length = 0;
    for (let index = 0; index < this.sheets.length; index++) {
      this.facades.push(new Sheet(this, index));
    }
    this.cells.clear();
    for (const [key, cell] of carried) {
      this.cells.set(key, cell);
    }
    this.rewireAll();
  }

  /**
   * Re-reads every formula in the workbook.
   *
   * The blunt answer, and the right one for the things that need it:
   * a name defined, a sheet added, renamed, moved or removed. Each of
   * those changes what formulas *read* rather than only what they
   * answer, so the graph's edges are wrong until they are rebuilt —
   * and the formulas affected are not findable without parsing all of
   * them anyway. All four are things a person does by hand, one at a
   * time.
   */
  private rewireAll(): void {
    this.graph.clear();
    this.volatile.clear();
    this.dynamic.clear();
    this.redone.clear();
    for (const [key, cell] of this.cells) {
      if (cell.formula !== null) {
        this.wire(key, cell.formula);
        this.classify(key, cell.formula);
        this.dirty.add(key);
      }
    }
    this.plan = null;
  }

  // ---------------------------------------------------------------------
  // Editing
  // ---------------------------------------------------------------------

  setCell(sheet: number, row: number, column: number, input: string, asText = false): void {
    const key = keyOn(sheet, row, column);
    if (input === '') {
      this.clearCell(sheet, row, column);
      return;
    }
    const entry = this.sheets[sheet];
    if (entry !== undefined) {
      entry.usedRows = Math.max(entry.usedRows, row + 1);
    }

    if (input.startsWith('=') && !asText) {
      this.writeFormula(key, input);
    } else {
      this.graph.clearPrecedents(key);
      this.dirty.delete(key);
      this.forget(key);
      this.cells.set(key, { input, formula: null, value: asText ? input : literalValue(input) });
    }
    this.markDependentsDirty(key);
  }

  clearCell(sheet: number, row: number, column: number): void {
    const key = keyOn(sheet, row, column);
    this.graph.clearPrecedents(key);
    this.cells.delete(key);
    this.dirty.delete(key);
    this.forget(key);
    this.markDependentsDirty(key);
  }

  /**
   * Inserts or deletes whole rows or columns of one sheet.
   *
   * Three things happen. The cells on the moved side of the line are
   * **carried** to their new keys — a rebuild of the store rather
   * than an in-place shuffle, because moving a cell to a key that
   * another cell has not vacated yet is how an in-place version
   * corrupts itself.
   *
   * **Every formula in the workbook is offered the shift**, not only
   * the ones on this sheet. A reference is a position, not an offset:
   * `Sheet2!A900` names the cell at A900 of Sheet 2, and after a row
   * is inserted there that cell is A901, wherever the formula reading
   * it happens to live.
   *
   * And the shift is **qualified by the sheet**, so a formula on
   * Sheet 1 reading its own `A900` is left alone when Sheet 2 grows a
   * row. That is what `Shift.sheet` is for, and it is the one thing
   * the single-sheet version could not have got wrong.
   *
   * Returns how many formulas were rewritten, for the budget spec.
   */
  shift(sheet: number, shift: Shift): number {
    const entry = this.sheets[sheet];
    if (entry === undefined) {
      return 0;
    }
    const named: Shift = { ...shift, sheet: entry.name };
    // A name is a reference by another spelling, so it moves like
    // one. Before the cells, because the rewiring below reads it.
    this.names.shift(named);
    const carried = new Map<number, Cell>();
    let rewrites = 0;

    for (const [key, cell] of this.cells) {
      const on = sheetOf(key);
      const row = rowOf(key);
      const column = columnOf(key);
      let at = key;
      if (on === sheet) {
        const index = shift.axis === 'row' ? row : column;
        const moved = shiftIndex(index, shift);
        if (moved === -1) {
          // The cell was in a deleted row. It goes, and everything
          // that referenced it will say `#REF!` after the rewrite.
          continue;
        }
        at =
          moved === index
            ? key
            : shift.axis === 'row'
              ? keyOn(on, moved, column)
              : keyOn(on, row, moved);
      }
      const input = shiftFormula(cell.input, named, this.nameOf(on));
      if (input !== cell.input) {
        rewrites++;
      }
      carried.set(at, input === cell.input ? cell : { input, formula: null, value: null });
    }

    // Rebuilt rather than patched. The graph's edges are keys, and
    // after a shift every key on the moved side is wrong; re-reading
    // each formula is the only version that cannot be half-right.
    this.cells.clear();
    this.graph.clear();
    this.dirty.clear();
    this.volatile.clear();
    this.dynamic.clear();
    this.redone.clear();
    this.plan = null;
    for (const [key, cell] of carried) {
      if (cell.formula === null && cell.value === null) {
        // Rewritten above, so it has to be parsed again.
        this.writeFormula(key, cell.input);
      } else {
        this.cells.set(key, cell);
        if (cell.formula !== null) {
          this.wire(key, cell.formula);
          this.classify(key, cell.formula);
          this.dirty.add(key);
        }
      }
    }
    for (const [key, cell] of this.cells) {
      if (cell.formula !== null) {
        this.dirty.add(key);
      }
      // An insert pushes cells past the old mark, and a whole-column
      // reference that stopped at it would stop reading them.
      const grown = this.sheets[sheetOf(key)];
      if (grown !== undefined) {
        grown.usedRows = Math.max(grown.usedRows, rowOf(key) + 1);
      }
    }
    return rewrites;
  }

  private writeFormula(key: number, input: string): void {
    let formula: Ast | null = null;
    try {
      formula = parseFormula(input.slice(1));
    } catch (error) {
      if (!(error instanceof FormulaSyntaxError)) {
        throw error;
      }
    }
    if (formula === null) {
      // A formula that does not parse is kept as typed — losing what
      // someone wrote because they have not finished writing it would
      // be worse than showing an error next to it — and holds
      // `#VALUE!` until it does parse.
      this.graph.clearPrecedents(key);
      this.dirty.delete(key);
      this.forget(key);
      this.cells.set(key, { input, formula: null, value: VALUE });
      return;
    }
    this.cells.set(key, { input, formula, value: null });
    this.wire(key, formula);
    this.classify(key, formula);
    this.dirty.add(key);
  }

  /** Puts a formula's edges into the graph, both kinds. */
  private wire(key: number, formula: Ast): void {
    const on = sheetOf(key);
    const { keys, columns } = this.precedentsOf(formula, on);
    /**
     * A named range is read, so it is a precedent.
     *
     * Without this a formula saying `=SUM(Sales)` has no edge to the
     * cells it sums, and editing one of them leaves the total stale —
     * the silent kind of wrong, because the formula looks right and
     * the number is old.
     */
    const words = new Set<string>();
    bareWordsOf(formula, words);
    for (const word of words) {
      const range = this.names.rangeOf(word);
      if (range === null) {
        continue;
      }
      const sheet = this.sheetOfRef(range.start, on);
      if (sheet === null) {
        continue;
      }
      if (range.wholeColumn === true) {
        for (let column = range.start.column; column <= range.end.column; column++) {
          columns.push(columnKey(sheet, column));
        }
        continue;
      }
      for (const cell of rangeKeys(range, sheet)) {
        keys.push(cell);
      }
    }
    this.graph.setPrecedents(key, keys);
    this.graph.setWatchedColumns(key, columns);
  }

  /**
   * What a formula reads: cells by key, and columns by column key.
   *
   * Two lists because a whole-column reference cannot be a key list —
   * `A:A` is 1,048,576 cells — and must not be turned into one. The
   * graph watches those columns instead; see `DependencyGraph`.
   *
   * A reference to a sheet the workbook does not have contributes no
   * edge at all. There is nothing for it to depend on, and it reads
   * `#REF!` for exactly as long as that is true; creating the sheet
   * rewires everything, which is where the edge comes from.
   */
  private precedentsOf(formula: Ast, on: number): { keys: number[]; columns: number[] } {
    const keys: number[] = [];
    const columns: number[] = [];
    referencesOf(formula, {
      ref: (ref: CellRef) => {
        const sheet = this.sheetOfRef(ref, on);
        if (sheet !== null) {
          keys.push(keyOn(sheet, ref.row, ref.column));
        }
      },
      range: (range: RangeRef) => {
        const sheet = this.sheetOfRef(range.start, on);
        if (sheet === null) {
          return;
        }
        if (range.wholeColumn === true) {
          for (let column = range.start.column; column <= range.end.column; column++) {
            columns.push(columnKey(sheet, column));
          }
          return;
        }
        for (const key of rangeKeys(range, sheet)) {
          keys.push(key);
        }
      }
    });
    return { keys, columns };
  }

  /** Which sheet a reference points at: the one it names, or its own. */
  private sheetOfRef(ref: CellRef, on: number): number | null {
    return ref.sheet === undefined ? on : this.sheetFor(ref.sheet);
  }

  private classify(key: number, formula: Ast): void {
    const names = new Set<string>();
    callNamesOf(formula, names);
    let isVolatile = false;
    let isDynamic = false;
    for (const name of names) {
      isVolatile ||= VOLATILE.has(name);
      isDynamic ||= name === 'INDIRECT' || name === 'OFFSET';
    }
    setMembership(this.volatile, key, isVolatile);
    setMembership(this.dynamic, key, isDynamic);
  }

  private forget(key: number): void {
    this.volatile.delete(key);
    this.dynamic.delete(key);
    this.redone.delete(key);
  }

  /**
   * Marks everything downstream of a changed cell.
   *
   * The closure is taken now rather than at `recalculate` because the
   * graph is correct now: a later edit may rewire it, and a cell that
   * needed recomputing because of *this* edit still needs it.
   */
  private markDependentsDirty(key: number): void {
    for (const dependent of this.graph.closureOf([key])) {
      this.dirty.add(dependent);
    }
    // An edit is a recalculation event, and a volatile formula is one
    // that has to be redone on every one of them however far away the
    // edit was.
    if (this.volatile.size > 0) {
      for (const cell of this.volatile) {
        this.dirty.add(cell);
      }
      for (const dependent of this.graph.closureOf(this.volatile)) {
        this.dirty.add(dependent);
      }
    }
    // A new edit is a new chance for the dynamic formulas to settle.
    this.redone.clear();
    // The clock moves on between edits, so two recalculations may
    // legitimately disagree about the time; within one they may not.
    this.moment = null;
    // Any plan in flight was ordered over a different set.
    this.plan = null;
  }

  // ---------------------------------------------------------------------
  // Recalculation
  // ---------------------------------------------------------------------

  /** Cells still waiting to be evaluated, anywhere in the workbook. */
  get pending(): number {
    return this.dirty.size;
  }

  /**
   * Evaluates at most `budget` cells, in an order where every cell
   * comes after the cells it reads.
   *
   * One queue for the workbook, because a cross-sheet formula has no
   * evaluation order that two queues could agree on.
   */
  recalculate(budget: number = Number.POSITIVE_INFINITY): RecalcResult {
    let evaluated = 0;
    /** A cap on how many times one call will rebuild its plan. */
    let plans = 0;

    while (this.dirty.size > 0 && evaluated < budget && plans < 64) {
      if (this.plan === null) {
        this.plan = this.buildPlan();
        plans++;
      }
      const plan = this.plan;
      while (plan.at < plan.order.length && evaluated < budget) {
        const key = plan.order[plan.at++];
        // A cell can leave the dirty set between plans, when an edit
        // turned a formula into a literal.
        if (!this.dirty.delete(key)) {
          continue;
        }
        this.evaluateCell(key);
        evaluated++;
      }
      if (plan.at < plan.order.length) {
        break;
      }
      this.plan = null;
    }

    this.stats.evaluated += evaluated;
    const done = this.plan === null && this.dirty.size === 0;
    if (done) {
      this.redone.clear();
    }
    return { evaluated, done };
  }

  private buildPlan(): { order: number[]; at: number } {
    const { order, circular } = this.graph.topological(this.dirty);
    for (const key of circular) {
      const cell = this.cells.get(key);
      if (cell !== undefined) {
        cell.value = CIRC;
      }
      this.dirty.delete(key);
    }
    this.stats.plans++;
    this.stats.planned += order.length;
    return { order, at: 0 };
  }

  private evaluateCell(key: number): void {
    const cell = this.cells.get(key);
    if (cell === undefined || cell.formula === null) {
      return;
    }
    const context = this.contextOn(sheetOf(key));
    if (!this.dynamic.has(key)) {
      cell.value = evaluate(cell.formula, context);
      return;
    }

    /**
     * A formula whose edges have to be discovered by running it.
     *
     * The reads are recorded and the graph is rebuilt from them, so
     * that editing the cell an `INDIRECT` landed on wakes the formula
     * that read it.
     */
    const reads = new Set<number>();
    this.recording = reads;
    try {
      cell.value = evaluate(cell.formula, context);
    } finally {
      this.recording = null;
    }
    if (!sameKeys(this.graph.precedentsOf(key), reads)) {
      this.graph.setPrecedents(key, reads);
    }
    if (!this.redone.has(key) && anyDirty(reads, this.dirty)) {
      this.redone.add(key);
      this.dirty.add(key);
    }
  }

  // ---------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------

  /**
   * The evaluation context for a formula on a given sheet.
   *
   * Built per cell rather than held, because `onSheet` is the only
   * thing that differs and a recalculation crosses sheets freely. The
   * object is three properties and a closure; the alternative is a
   * mutable field that would have to be right at every re-entry.
   */
  private contextOn(sheet: number) {
    return {
      valueAt: (key: number) => this.valueAt(key),
      functions: this.functions,
      rangeForName: (name: string) => this.names.rangeOf(name),
      book: this.asContext,
      onSheet: sheet
    };
  }

  private readonly asContext: WorkbookContext = {
    sheetFor: (name: string) => this.sheetFor(name),
    usedRowsOf: (sheet: number) => this.sheets[sheet]?.usedRows ?? 0
  };

  /**
   * Re-reads every formula, because a name changed underneath them.
   *
   * Defining or removing a name changes what formulas *read*, not
   * just what they answer.
   */
  namesChanged(): void {
    this.rewireAll();
  }

  /** The value behind a key, wherever in the workbook it is. */
  valueAt(key: number): CellValue {
    // Recorded only while a dynamic formula is running, which is the
    // one case where what was read is not what the tree said.
    this.recording?.add(key);
    return this.cells.get(key)?.value ?? null;
  }

  /**
   * The clock and dice half of `EvaluationContext`.
   *
   * The time is *read once* and handed back as a constant: two
   * `NOW()`s in one recalculation that each called the clock would
   * disagree by however long the pass took.
   */
  get functions(): FunctionContext {
    if (this.moment === null) {
      const at = this.clock();
      this.moment = { now: () => at, random: () => this.dice() };
    }
    return this.moment;
  }

  usedRowsOf(sheet: number): number {
    return this.sheets[sheet]?.usedRows ?? 0;
  }

  value(sheet: number, row: number, column: number): CellValue {
    return this.valueAt(keyOn(sheet, row, column));
  }

  /** What was typed, which is what a cell editor opens with. */
  input(sheet: number, row: number, column: number): string {
    return this.cells.get(keyOn(sheet, row, column))?.input ?? '';
  }

  display(sheet: number, row: number, column: number): string {
    return formatValue(this.value(sheet, row, column));
  }

  /** Cells that hold something on a sheet. */
  sizeOf(sheet: number): number {
    let count = 0;
    for (const key of this.cells.keys()) {
      if (sheetOf(key) === sheet) {
        count++;
      }
    }
    return count;
  }

  /** Cells that hold something, anywhere. */
  get size(): number {
    return this.cells.size;
  }

  /** Every non-empty cell of a sheet, for a repository to write out. */
  *entriesOf(sheet: number): Generator<{ row: number; column: number; input: string }> {
    for (const [key, cell] of this.cells) {
      if (sheetOf(key) === sheet) {
        yield { row: rowOf(key), column: columnOf(key), input: cell.input };
      }
    }
  }

  /** What a cell reads, for `engine.explain` in Phase 7. */
  precedentsOfCell(sheet: number, row: number, column: number): number[] {
    return [...this.graph.precedentsOf(keyOn(sheet, row, column))];
  }

  /** What reads a cell. */
  dependentsOfCell(sheet: number, row: number, column: number): number[] {
    return [...this.graph.dependentsOf(keyOn(sheet, row, column))];
  }
}

/**
 * A formula with one sheet name swapped for another.
 *
 * Done on the text rather than on the tree, and the reason is that a
 * tree printed back is a *different* text: `=A1+B1` comes back as
 * `=(A1+B1)`, because the printer is fully parenthesised on purpose.
 * Rewriting every formula in the workbook through it would reformat
 * formulas that have nothing to do with the rename.
 *
 * So the tokenizer says where the sheet names are and only those
 * spans are replaced. The scan is the authority on what is a sheet
 * name, which is what stops a rename from `Data` to `Figures` from
 * touching the text inside `="Data"`.
 */
function renameInFormula(input: string, from: string, to: string): string {
  if (!input.startsWith('=')) {
    return input;
  }
  const source = input.slice(1);
  let tokens;
  try {
    tokens = tokenize(source);
  } catch {
    // Text that does not tokenize is left exactly as it was typed.
    return input;
  }
  const wanted = from.toUpperCase();
  let out = '';
  let at = 0;
  for (const token of tokens) {
    if (token.kind !== 'sheet' || token.value.toUpperCase() !== wanted) {
      continue;
    }
    out += source.slice(at, token.start) + `${quoteSheetName(to)}!`;
    at = token.end;
  }
  return at === 0 ? input : `=${out}${source.slice(at)}`;
}


/**
 * What typing something that is not a formula means.
 *
 * Numbers, TRUE/FALSE and dates are recognised; everything else is
 * text. A date becomes its serial number, because in a spreadsheet
 * that is what a date *is* — see `Dates.ts`.
 */
function literalValue(input: string): CellValue {
  const trimmed = input.trim();
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') {
    return true;
  }
  if (upper === 'FALSE') {
    return false;
  }
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }
  return parseTypedDate(trimmed)?.serial ?? input;
}

/** Adds or removes a key, which reads better than a branch at each call. */
function setMembership(into: Set<number>, key: number, member: boolean): void {
  if (member) {
    into.add(key);
  } else {
    into.delete(key);
  }
}

function sameKeys(left: ReadonlySet<number>, right: ReadonlySet<number>): boolean {
  if (left.size !== right.size) {
    return false;
  }
  for (const key of left) {
    if (!right.has(key)) {
      return false;
    }
  }
  return true;
}

function anyDirty(keys: ReadonlySet<number>, dirty: ReadonlySet<number>): boolean {
  for (const key of keys) {
    if (dirty.has(key)) {
      return true;
    }
  }
  return false;
}
