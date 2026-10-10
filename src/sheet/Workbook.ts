import {
  columnKey,
  columnOf,
  inBounds,
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
import { evaluateArray, type WorkbookContext } from './Evaluator';
import type { ScriptFunctions } from './ScriptFunctions';
import { isArray, type ArrayValue } from './FunctionKit';
import { nowSerial, UNCERTAIN, VOLATILE, type FunctionContext } from './Functions';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { Sheet } from './Sheet';
import { Words } from './Words';
import { shiftFormula, shiftIndex, type Shift } from './Shift';
import { tokenize } from './Tokenizer';
import { CIRC, formatValue, SPILL, VALUE, type CellValue } from './Values';

interface Cell {
  /** Exactly what was typed, which is what an editor puts back. */
  readonly input: string;
  /** The parsed formula, or null for a literal. */
  readonly formula: Ast | null;
  value: CellValue;
}

/**
 * A formula whose answer is an array, and the cells it covers.
 *
 * The area is where the array *would* go, kept whether or not it got
 * there: a spill that is blocked has to be tried again when whatever
 * is in its way is cleared, and that is found by asking which areas a
 * written cell falls in. `values` is null while it is blocked.
 */
interface Spill {
  readonly keys: readonly number[];
  readonly columns: number;
  readonly values: readonly CellValue[] | null;
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

/** An input typed into a fork: one cell of a scenario. */
export interface Override {
  readonly sheet: number;
  readonly row: number;
  readonly column: number;
  readonly input: string;
  /** Written as text even if it reads as a number or a formula, as a Text-formatted cell is. */
  readonly asText?: boolean;
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
  private graph = new DependencyGraph();
  /** Cells whose value is out of date. */
  private readonly dirty = new Set<number>();
  /**
   * The order the current slice is working through, and where it is,
   * and the cells that could not be ordered because they are in a
   * circle or downstream of one.
   */
  private plan: { order: number[]; at: number; circular: number[] } | null = null;

  /**
   * Whether circular formulas are worked out by going round them, and
   * how far: Excel's iterative calculation, off unless a workbook asks
   * for it. Off, a circle is `#CIRC!`, which is almost always what a
   * circle is — a mistake. On, each cell in the circle is evaluated in
   * turn, from the values the last round left, until no value moves by
   * more than `delta` or `count` rounds have gone by: how an interest
   * payment that depends on a balance that depends on it is written.
   */
  iteration: { readonly count: number; readonly delta: number } | null = null;

  /**
   * The functions the workbook's script defines, or null for none; see
   * `ScriptFunctions`. Set by the app layer, which owns the interpreter,
   * and followed by `scriptsChanged` so every formula reads them again.
   */
  scripts: ScriptFunctions | null = null;

  /**
   * How far a sheet goes, which is as far as an array may spill: one
   * that would run off the edge says `#SPILL!`, as in Excel, rather
   * than filling rows nobody can scroll to. Excel's own limits when the
   * application has not said.
   */
  extent: { readonly rows: number; readonly columns: number } | null = null;

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
   * Formulas that state a guess — `NORMAL`, `UNIFORM`, `TRIANGULAR` —
   * which a simulation draws afresh in every trial; see `resample`.
   */
  private readonly uncertain = new Set<number>();

  /**
   * Where a simulation's draws come from, or null outside one, when an
   * uncertain formula is its likeliest value. Set on a fork, never on
   * the workbook somebody is editing.
   */
  sampler: (() => number) | null = null;

  /**
   * How many edits the workbook has taken: a count that any change to
   * what a cell holds, or to the shape of a sheet, moves on. A result
   * worked out from the workbook — a simulation — is only true of it
   * while this is what it was.
   */
  private edited = 0;

  get edits(): number {
    return this.edited;
  }

  /**
   * Formulas whose precedents cannot be read off their own text.
   *
   * `=INDIRECT("A" & B1)` reads a cell that the formula never names,
   * so the graph built from the tree has an edge to B1 and none to
   * the cell it actually fetched.
   */
  private readonly dynamic = new Set<number>();

  /**
   * Formulas that call `SUBTOTAL`, which answer differently when a row
   * is hidden or filtered — a change to no cell at all, so no edge in
   * the graph wakes them; `visibilityChanged` does.
   */
  private readonly subtotals = new Set<number>();

  /**
   * Whether a row of a sheet is hidden by hand, filtered, or neither.
   *
   * The document's answer, not the workbook's: which rows show is a
   * fact about the screen, and the workbook holds cells. The document
   * installs it; without one every row shows, which is what a
   * workbook in a spec means.
   */
  rowState: ((sheet: number, row: number) => 'hidden' | 'filtered' | null) | null = null;

  /** The keys the cell now being evaluated read, when it is dynamic. */
  private recording: Set<number> | null = null;

  /** Dynamic cells already sent round a second time since the last edit. */
  private readonly redone = new Set<number>();

  /** Fixed for a whole recalculation, so two `NOW()`s agree. */
  private moment: FunctionContext | null = null;

  /**
   * True for a workbook made by `fork`, whose volatile formulas keep
   * the values they were forked with: an edit here does not send every
   * `RAND()` round again, only one the edit itself reaches. A scenario
   * is a question about the inputs it changes, and a fork whose dice
   * had been rolled again would differ from its base in every cell
   * downstream of one, whatever the scenario changed.
   */
  private pinned = false;

  /**
   * Formulas that spill, by the key of the cell they are written in.
   *
   * The cells an array spills into hold nothing of their own: reading
   * one reads the array, through `spilled`, and each carries one edge
   * in the graph — to the formula — so everything that reads a spilled
   * cell is ordered after the formula that fills it, by the same
   * topological sort as every other cell. Nothing is stored for them,
   * so nothing is saved: the formula is the cell, and the rest is its
   * answer.
   */
  private readonly spills = new Map<number, Spill>();
  /** The cells an array has filled, and which formula filled them. */
  private readonly spilled = new Map<number, number>();
  /**
   * The words each column holds, for AutoComplete, by `columnKey`.
   *
   * Built the first time a column is asked and kept up to date by
   * every write after that, so a keystroke costs a binary search and
   * never a walk. Anything that moves cells wholesale — a shift, a
   * sheet moved or removed — drops the lot, and the next question
   * builds again.
   */
  private readonly words = new Map<number, Words>();

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
    this.words.clear();
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
    this.edited++;
    this.graph.clear();
    // The edges from spilled cells went with the graph; every formula is
    // dirtied below and spills again.
    this.spills.clear();
    this.spilled.clear();
    this.volatile.clear();
    this.uncertain.clear();
    this.dynamic.clear();
    this.subtotals.clear();
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
    this.unword(sheet, column, key);
    this.dropSpill(key);
    this.spilled.delete(key);
    this.wakeSpillsOver(key);
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
    this.enword(sheet, column, key);
    this.markDependentsDirty(key);
  }

  /**
   * What typing `prefix` into a column would complete to: the one text
   * the column holds that starts that way, or null.
   */
  completeIn(sheet: number, column: number, prefix: string): string | null {
    return this.wordsOf(sheet, column).complete(prefix);
  }

  private wordsOf(sheet: number, column: number): Words {
    const at = columnKey(sheet, column);
    let words = this.words.get(at);
    if (words === undefined) {
      words = new Words();
      for (const [key, cell] of this.cells) {
        if (sheetOf(key) === sheet && columnOf(key) === column && isWord(cell)) {
          words.add(cell.input);
        }
      }
      this.words.set(at, words);
    }
    return words;
  }

  /** A cell about to change, out of its column's words if they are kept. */
  private unword(sheet: number, column: number, key: number): void {
    const words = this.words.get(columnKey(sheet, column));
    const cell = this.cells.get(key);
    if (words !== undefined && cell !== undefined && isWord(cell)) {
      words.remove(cell.input);
    }
  }

  private enword(sheet: number, column: number, key: number): void {
    const words = this.words.get(columnKey(sheet, column));
    const cell = this.cells.get(key);
    if (words !== undefined && cell !== undefined && isWord(cell)) {
      words.add(cell.input);
    }
  }

  clearCell(sheet: number, row: number, column: number): void {
    const key = keyOn(sheet, row, column);
    if (!this.cells.has(key)) {
      // Nothing is stored in a cell an array spilled into, so clearing
      // it is clearing nothing — the array is its formula's.
      return;
    }
    this.dropSpill(key);
    this.wakeSpillsOver(key);
    this.graph.clearPrecedents(key);
    this.unword(sheet, column, key);
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
    this.edited++;
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
    this.words.clear();
    this.graph.clear();
    this.dirty.clear();
    // Every formula is re-read and dirtied below, and each spills again
    // from where it has moved to.
    this.spills.clear();
    this.spilled.clear();
    this.volatile.clear();
    this.uncertain.clear();
    this.dynamic.clear();
    this.subtotals.clear();
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
     * A name that holds a formula is read *through*: `=Double(A1)`
     * depends on A1, and on every cell `Double`'s own formula reads, and
     * on what the names inside that read. Otherwise a `LAMBDA` whose body
     * reads `TaxRate` leaves every caller stale when the rate changes.
     * On the caller's sheet, because that is where the name's unqualified
     * references are read.
     */
    const bodies = this.namedBodiesOf(formula);
    for (const body of bodies) {
      const through = this.precedentsOf(body, on);
      keys.push(...through.keys);
      columns.push(...through.columns);
    }
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
    for (const body of bodies) {
      bareWordsOf(body, words);
    }
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
   * The formulas of every name a formula reaches, through names inside
   * names, each once — so a `LAMBDA` that calls itself by name is read
   * once and not forever.
   */
  private namedBodiesOf(formula: Ast): Ast[] {
    if (this.names.size === 0) {
      return [];
    }
    const bodies: Ast[] = [];
    const seen = new Set<string>();
    const visit = (node: Ast): void => {
      const called = new Set<string>();
      callNamesOf(node, called);
      for (const name of called) {
        if (seen.has(name)) {
          continue;
        }
        seen.add(name);
        const body = this.names.formulaOf(name);
        if (body !== null) {
          bodies.push(body);
          visit(body);
        }
      }
    };
    visit(formula);
    return bodies;
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
    for (const body of this.namedBodiesOf(formula)) {
      callNamesOf(body, names);
    }
    let isVolatile = false;
    let isDynamic = false;
    let isUncertain = false;
    for (const name of names) {
      isVolatile ||= VOLATILE.has(name);
      isUncertain ||= UNCERTAIN.has(name);
      isDynamic ||= name === 'INDIRECT' || name === 'OFFSET';
    }
    setMembership(this.volatile, key, isVolatile);
    setMembership(this.uncertain, key, isUncertain);
    setMembership(this.dynamic, key, isDynamic);
    setMembership(this.subtotals, key, names.has('SUBTOTAL'));
  }

  /**
   * Rows were hidden, shown or filtered, so every `SUBTOTAL` — and
   * whatever reads one — has to be asked again. Called by the document
   * after it changes which rows show.
   */
  visibilityChanged(): void {
    if (this.subtotals.size === 0) {
      return;
    }
    for (const key of this.subtotals) {
      this.dirty.add(key);
    }
    for (const dependent of this.graph.closureOf(this.subtotals)) {
      this.dirty.add(dependent);
    }
    this.plan = null;
  }

  private forget(key: number): void {
    this.volatile.delete(key);
    this.uncertain.delete(key);
    this.dynamic.delete(key);
    this.subtotals.delete(key);
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
    this.edited++;
    for (const dependent of this.graph.closureOf([key])) {
      this.dirty.add(dependent);
    }
    // An edit is a recalculation event, and a volatile formula is one
    // that has to be redone on every one of them however far away the
    // edit was. Not in a fork, which keeps its base's roll of the dice.
    if (this.volatile.size > 0 && !this.pinned) {
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
  // Forks
  // ---------------------------------------------------------------------

  /**
   * A second workbook that starts as this one and then has `overrides`
   * typed into it: a scenario.
   *
   * Copied, not rebuilt. A rebuild types every input in again, which
   * parses every formula and recalculates every cell, nearly all of
   * them to the answer they already have. A fork copies the cell
   * records with their values and shares their parsed formulas, which
   * nothing ever writes, then writes the overrides through `setCell`
   * like any edit, so what is dirty is what the overrides reach and the
   * recalculation that follows is that and nothing else.
   *
   * Everything an edit or a recalculation writes is the fork's own: the
   * cells, the graph, the dirty set, the sheets' entries. Spills are
   * shared as they stand, because a spill is replaced and never changed.
   * The fork reads the same script functions, rows and extent as this
   * workbook, and its volatile formulas are pinned (see `pinned`).
   *
   * Taken from a settled workbook, so that what the fork starts from is
   * an answer; a fork of one still recalculating carries the dirty set
   * across and finishes the work itself.
   */
  fork(overrides: readonly Override[] = []): Workbook {
    const copy = new Workbook([]);
    copy.names.restore(this.names.all());
    for (const entry of this.sheets) {
      copy.sheets.push({ ...entry });
      copy.facades.push(new Sheet(copy, copy.facades.length));
    }
    for (const [key, cell] of this.cells) {
      copy.cells.set(key, { input: cell.input, formula: cell.formula, value: cell.value });
    }
    copy.graph = this.graph.clone();
    for (const key of this.dirty) {
      copy.dirty.add(key);
    }
    for (const key of this.volatile) {
      copy.volatile.add(key);
    }
    for (const key of this.uncertain) {
      copy.uncertain.add(key);
    }
    for (const key of this.dynamic) {
      copy.dynamic.add(key);
    }
    for (const key of this.subtotals) {
      copy.subtotals.add(key);
    }
    for (const [key, spill] of this.spills) {
      copy.spills.set(key, spill);
    }
    for (const [key, at] of this.spilled) {
      copy.spilled.set(key, at);
    }
    copy.iteration = this.iteration;
    copy.scripts = this.scripts;
    copy.extent = this.extent;
    copy.clock = this.clock;
    copy.dice = this.dice;
    copy.rowState = this.rowState;
    copy.pinned = true;
    for (const override of overrides) {
      copy.setCell(override.sheet, override.row, override.column, override.input, override.asText ?? false);
    }
    return copy;
  }

  // ---------------------------------------------------------------------
  // Uncertainty
  // ---------------------------------------------------------------------

  /** The cells holding a guess, by key; see `UNCERTAIN`. */
  get uncertainCells(): ReadonlySet<number> {
    return this.uncertain;
  }

  /**
   * Every cell a guess reaches: the guesses themselves and everything
   * downstream of them, which are the cells a simulation has something
   * to say about. Everything else comes out the same in every trial.
   */
  uncertainReach(): Set<number> {
    const reach = this.graph.closureOf(this.uncertain);
    for (const key of this.uncertain) {
      reach.add(key);
    }
    return reach;
  }

  /**
   * Marks every guess and everything it reaches to be worked out again:
   * one trial, once `recalculate` has run. With a `sampler` set, each
   * guess is drawn afresh; without one each comes back to its likeliest
   * value.
   */
  resample(): void {
    for (const key of this.uncertainReach()) {
      this.dirty.add(key);
    }
    this.moment = null;
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
    // A slice is bounded by time spent in scripts as well as by count:
    // two thousand cells of arithmetic is a millisecond, and two
    // thousand calls to a slow function is not.
    const scripts = this.scripts;
    scripts?.startSlice();
    const within = (): boolean => evaluated < budget && !(scripts?.overBudget() ?? false);

    while (this.dirty.size > 0 && within() && plans < 64) {
      if (this.plan === null) {
        this.plan = this.buildPlan();
        plans++;
      }
      const plan = this.plan;
      while (plan.at < plan.order.length && within()) {
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
      // The circle last, because it may read cells the plan has just
      // brought up to date.
      evaluated += this.goRound(plan.circular);
      this.plan = null;
    }

    this.stats.evaluated += evaluated;
    const done = this.plan === null && this.dirty.size === 0;
    if (done) {
      this.redone.clear();
    }
    return { evaluated, done };
  }

  private buildPlan(): { order: number[]; at: number; circular: number[] } {
    const { order, circular } = this.graph.topological(this.dirty);
    for (const key of circular) {
      if (this.iteration === null) {
        const cell = this.cells.get(key);
        if (cell !== undefined) {
          cell.value = CIRC;
        }
      }
      this.dirty.delete(key);
    }
    this.stats.plans++;
    this.stats.planned += order.length;
    return { order, at: 0, circular: this.iteration === null ? [] : circular };
  }

  /**
   * The cells of a circle, evaluated in turn until they settle: the
   * order Excel goes round in, sheet by sheet and row by row, each cell
   * reading what the others were left at. Returns how many evaluations
   * that took.
   */
  private goRound(circular: readonly number[]): number {
    const iteration = this.iteration;
    if (iteration === null || circular.length === 0) {
      return 0;
    }
    const order = circular.filter(key => this.cells.get(key)?.formula != null).sort((a, b) => a - b);
    // A circle refused before starts from nothing: `#CIRC!` is not a
    // value to go round from, and every round would only copy it.
    for (const key of order) {
      const cell = this.cells.get(key);
      if (cell !== undefined && cell.value === CIRC) {
        cell.value = null;
      }
    }
    let evaluated = 0;
    for (let round = 0; round < iteration.count; round++) {
      let moved = 0;
      for (const key of order) {
        const cell = this.cells.get(key);
        if (cell === undefined) {
          continue;
        }
        const before = cell.value;
        this.evaluateCell(key);
        evaluated++;
        const after = cell.value;
        moved = Math.max(
          moved,
          typeof before === 'number' && typeof after === 'number'
            ? Math.abs(after - before)
            : before === after
              ? 0
              : Number.POSITIVE_INFINITY
        );
      }
      if (moved < iteration.delta) {
        break;
      }
    }
    // Evaluating put some of them back in the dirty set, through the
    // spills or through the dynamic formulas; they are settled.
    for (const key of order) {
      this.dirty.delete(key);
    }
    return evaluated;
  }

  private evaluateCell(key: number): void {
    const cell = this.cells.get(key);
    if (cell === undefined || cell.formula === null) {
      return;
    }
    const context = this.contextOn(sheetOf(key), key);
    if (!this.dynamic.has(key)) {
      this.settle(key, cell, evaluateArray(cell.formula, context));
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
      this.settle(key, cell, evaluateArray(cell.formula, context));
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

  /**
   * A formula's answer, into its cell — and, for an array, into the
   * cells beside and below it.
   *
   * The array goes where it fits or not at all: a cell in the way that
   * holds anything, or one another array already filled, and the
   * formula says `#SPILL!` and fills nothing. Whatever reads a cell
   * whose spilled value changed is dirtied here, since the graph only
   * learns of those cells from this.
   */
  private settle(key: number, cell: Cell, result: CellValue | ArrayValue): void {
    if (!isArray(result)) {
      this.dropSpill(key);
      cell.value = resultOf(result);
      return;
    }
    const sheet = sheetOf(key);
    const row = rowOf(key);
    const column = columnOf(key);
    const previous = this.spills.get(key);
    const keys: number[] = [];
    let blocked =
      !inBounds(row + result.rows - 1, column + result.columns - 1) ||
      (this.extent !== null && (row + result.rows > this.extent.rows || column + result.columns > this.extent.columns));
    if (!blocked) {
      for (let r = 0; r < result.rows; r++) {
        for (let c = 0; c < result.columns; c++) {
          const at = keyOn(sheet, row + r, column + c);
          keys.push(at);
          if (at !== key && (this.cells.has(at) || (this.spilled.get(at) ?? key) !== key)) {
            blocked = true;
          }
        }
      }
    }
    const values = blocked ? null : result.values.map(resultOf);
    if (previous !== undefined && sameSpill(previous, keys, result.columns, values)) {
      cell.value = values === null ? SPILL : (values[0] ?? 0);
      return;
    }
    this.unfill(key, previous);
    this.spills.set(key, { keys, columns: result.columns, values });
    if (values !== null) {
      for (const at of keys) {
        if (at !== key) {
          this.spilled.set(at, key);
          this.graph.setPrecedents(at, [key]);
        }
      }
    }
    cell.value = values === null ? SPILL : (values[0] ?? 0);
    this.dirtyReaders(key, [...(previous?.keys ?? []), ...keys]);
  }

  /** Takes an array's values out of the cells it had filled. */
  private unfill(key: number, spill: Spill | undefined): void {
    if (spill === undefined || spill.values === null) {
      return;
    }
    for (const at of spill.keys) {
      if (this.spilled.get(at) === key) {
        this.spilled.delete(at);
        this.graph.clearPrecedents(at);
      }
    }
  }

  /** A formula that no longer spills, or no longer exists, gives its cells back. */
  private dropSpill(key: number): void {
    const spill = this.spills.get(key);
    if (spill === undefined) {
      return;
    }
    this.unfill(key, spill);
    this.spills.delete(key);
    this.dirtyReaders(key, spill.keys);
    // Another array blocked by this one may fit now.
    for (const at of spill.keys) {
      if (at !== key) {
        this.wakeSpillsOver(at);
      }
    }
  }

  /** Everything that reads the cells of an area, dirtied — never the formula itself. */
  private dirtyReaders(key: number, keys: readonly number[]): void {
    const around = keys.filter(at => at !== key);
    if (around.length === 0) {
      return;
    }
    for (const dependent of this.graph.closureOf(around)) {
      if (dependent !== key && this.cells.has(dependent)) {
        this.dirty.add(dependent);
      }
    }
    this.plan = null;
  }

  /**
   * A cell was written or cleared inside the area of some array, which
   * is now blocked, or free to spill again.
   */
  private wakeSpillsOver(key: number): void {
    if (this.spills.size === 0) {
      return;
    }
    for (const [anchor, spill] of this.spills) {
      if (anchor !== key && spill.keys.includes(key) && this.cells.get(anchor)?.formula != null) {
        this.dirty.add(anchor);
        this.plan = null;
      }
    }
  }

  /** The value an array put in a cell it spilled into. */
  private spilledValue(anchor: number, key: number): CellValue {
    const spill = this.spills.get(anchor);
    if (spill === undefined || spill.values === null) {
      return null;
    }
    const origin = spill.keys[0];
    const index = (rowOf(key) - rowOf(origin)) * spill.columns + (columnOf(key) - columnOf(origin));
    return spill.values[index] ?? null;
  }

  /**
   * The cells a formula's array fills, for the screen to show as its
   * own; null for a cell that is not an array formula's, or one that is
   * blocked.
   */
  spillOf(sheet: number, row: number, column: number): { rows: number; columns: number } | null {
    const spill = this.spills.get(keyOn(sheet, row, column));
    if (spill === undefined || spill.values === null) {
      return null;
    }
    return { rows: spill.keys.length / spill.columns, columns: spill.columns };
  }

  /** The formula whose array fills a cell, if one does. */
  spilledFrom(sheet: number, row: number, column: number): { row: number; column: number } | null {
    const anchor = this.spilled.get(keyOn(sheet, row, column));
    return anchor === undefined ? null : { row: rowOf(anchor), column: columnOf(anchor) };
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
  private contextOn(sheet: number, key?: number) {
    return {
      valueAt: (key: number) => this.valueAt(key),
      functions: this.functions,
      rangeForName: (name: string) => this.names.rangeOf(name),
      formulaForName: (name: string) => this.names.formulaOf(name),
      scripts: this.scripts ?? undefined,
      book: this.asContext,
      onSheet: sheet,
      at: key === undefined ? undefined : { row: rowOf(key), column: columnOf(key) },
      rowState: this.rowState ?? undefined,
      isSubtotal: (cell: number) => this.subtotals.has(cell),
      spillAt: (cell: number) => {
        const spill = this.spills.get(cell);
        return spill === undefined || spill.values === null
          ? null
          : { rows: spill.keys.length / spill.columns, columns: spill.columns };
      }
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

  /**
   * Re-reads every formula, because the script's functions changed:
   * defined, edited, or gone. A formula calling one that was `#NAME?` a
   * moment ago may have an answer now, and one whose function changed
   * has a different answer.
   */
  scriptsChanged(): void {
    this.rewireAll();
  }

  /** The value behind a key, wherever in the workbook it is. */
  valueAt(key: number): CellValue {
    // Recorded only while a dynamic formula is running, which is the
    // one case where what was read is not what the tree said.
    this.recording?.add(key);
    const cell = this.cells.get(key);
    if (cell !== undefined) {
      return cell.value;
    }
    const anchor = this.spilled.get(key);
    return anchor === undefined ? null : this.spilledValue(anchor, key);
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
      this.moment = {
        now: () => at,
        random: () => this.dice(),
        sample: () => (this.sampler === null ? null : this.sampler())
      };
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

  /**
   * Where Ctrl+Arrow lands from a cell: the edge of the data.
   *
   * From a filled cell with a filled one beside it, the last filled
   * cell of that run; otherwise the next filled cell in the direction,
   * or the edge of the sheet when there is none. A cell an array
   * spilled into is filled, as it is on screen.
   *
   * Inside a run this steps a cell at a time, which costs the run.
   * Across a gap it steps a little way and then, rather than walking a
   * million empty cells, reads the cells the sheet actually holds once
   * and takes the nearest — which costs the sheet's contents and is
   * never more than the walk it replaces on a sheet that is full.
   */
  edgeFrom(
    sheet: number,
    row: number,
    column: number,
    rows: -1 | 0 | 1,
    columns: -1 | 0 | 1,
    extent: { rowCount: number; columnCount: number }
  ): { row: number; column: number } {
    const inside = (r: number, c: number) => r >= 0 && c >= 0 && r < extent.rowCount && c < extent.columnCount;
    const filled = (r: number, c: number) => {
      const key = keyOn(sheet, r, c);
      return this.cells.has(key) || this.spilled.has(key);
    };
    let r = row + rows;
    let c = column + columns;
    if (!inside(r, c)) {
      return { row, column };
    }
    if (filled(row, column) && filled(r, c)) {
      while (inside(r + rows, c + columns) && filled(r + rows, c + columns)) {
        r += rows;
        c += columns;
      }
      return { row: r, column: c };
    }
    for (let step = 0; step < NEAR_GAP && inside(r, c); step++) {
      if (filled(r, c)) {
        return { row: r, column: c };
      }
      r += rows;
      c += columns;
    }
    // Along one line: the other coordinate is fixed, and distance is
    // how far along it a filled cell is.
    const along = rows !== 0 ? row : column;
    const direction = rows !== 0 ? rows : columns;
    let nearest = direction > 0 ? (rows !== 0 ? extent.rowCount : extent.columnCount) : -1;
    const consider = (key: number): void => {
      if (sheetOf(key) !== sheet) {
        return;
      }
      const keyRow = rowOf(key);
      const keyColumn = columnOf(key);
      const [at, across, fixed] = rows !== 0 ? [keyRow, keyColumn, column] : [keyColumn, keyRow, row];
      if (across !== fixed || (at - along) * direction <= 0) {
        return;
      }
      if (direction > 0 ? at < nearest : at > nearest) {
        nearest = at;
      }
    };
    for (const key of this.cells.keys()) {
      consider(key);
    }
    for (const key of this.spilled.keys()) {
      consider(key);
    }
    const edge = rows !== 0 ? extent.rowCount - 1 : extent.columnCount - 1;
    const landed = Math.min(Math.max(nearest, 0), edge);
    return rows !== 0 ? { row: landed, column } : { row, column: landed };
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
 * tree printed back is a *different* text: the spaces somebody typed
 * and the brackets they put in for their own reading are not in the
 * tree, so `= (A1 + B1) * 2` would come back as `=(A1+B1)*2`. Rewriting
 * every formula in the workbook through it would reformat formulas
 * that have nothing to do with the rename.
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
export function literalOf(input: string): CellValue {
  return literalValue(input);
}

/**
 * What a formula's cell holds once it is worked out: its answer, except
 * that an empty answer is zero.
 *
 * `=D7` with D7 empty shows 0 in Excel, and it matters beyond the
 * display — `COUNT` counts it, `AVERAGE` divides by it, `LEN` of it is
 * one. Inside a formula a blank stays a blank, so `ISBLANK(D7)` is still
 * true; it is only a *cell's* value that cannot be nothing once a
 * formula has put something there.
 */
function sameSpill(spill: Spill, keys: readonly number[], columns: number, values: readonly CellValue[] | null): boolean {
  if (spill.columns !== columns || spill.keys.length !== keys.length || (spill.values === null) !== (values === null)) {
    return false;
  }
  if (spill.keys.some((key, at) => key !== keys[at])) {
    return false;
  }
  if (spill.values === null || values === null) {
    return true;
  }
  return spill.values.every((value, at) => sameValue(value, values[at]));
}

function sameValue(a: CellValue, b: CellValue): boolean {
  if (typeof a === 'object' && a !== null && typeof b === 'object' && b !== null) {
    return a.code === b.code;
  }
  return a === b;
}

function resultOf(value: CellValue): CellValue {
  return value === null ? 0 : value;
}

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

/**
 * How far Ctrl+Arrow steps across empty cells before it reads the whole
 * sheet instead. A gap in a table is usually a row or two, and a step
 * is a lookup; the sheet's contents may be a hundred thousand cells.
 */
const NEAR_GAP = 64;

/**
 * Whether a cell is a word AutoComplete offers: text somebody typed,
 * not a number, a date, a formula or its answer.
 */
function isWord(cell: { readonly formula: unknown; readonly value: CellValue; readonly input: string }): boolean {
  return cell.formula === null && typeof cell.value === 'string' && cell.input !== '';
}
