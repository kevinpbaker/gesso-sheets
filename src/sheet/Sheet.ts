import { cellKey, columnOf, rangeKeys, rowOf, type CellRef, type RangeRef } from './A1';
import { callNamesOf, referencesOf, type Ast } from './Ast';
import { parseTypedDate } from './Dates';
import { DependencyGraph } from './DependencyGraph';
import { evaluate } from './Evaluator';
import { nowSerial, VOLATILE, type FunctionContext } from './Functions';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { shiftFormula, shiftIndex, type Shift } from './Shift';
import { CIRC, formatValue, VALUE, type CellValue } from './Values';

interface Cell {
  /** Exactly what was typed, which is what an editor puts back. */
  readonly input: string;
  /** The parsed formula, or null for a literal. */
  readonly formula: Ast | null;
  value: CellValue;
}

export interface RecalcResult {
  /** Cells evaluated by this call. */
  readonly evaluated: number;
  /** False when a budget ran out with cells still to do. */
  readonly done: boolean;
}

/**
 * The sheet: cells, what they read, and what has to be redone.
 *
 * Nothing here imports the framework, touches a worker or knows what a
 * viewport is. That is the phase's whole constraint and it buys two
 * things: these specs run in node in milliseconds, and Phase 2 can put
 * this behind a channel without it noticing.
 *
 * **Recalculation is a queue, not a call.** `setCell` marks work and
 * returns; `recalculate(budget)` does at most that many cells and says
 * whether more remain. Phase 0 measured why: thirty milliseconds of
 * uninterrupted application thread leaves the sheet blank in 89% of
 * frames while the render thread goes on scrolling at 60fps, because
 * the window it is asking for cannot be served until the recalc lets
 * go. A budget is what lets the worker interleave — recalculate a
 * slice, publish the window, recalculate the next — and it has to be
 * in the shape of the engine from the start, because a recalc written
 * as one long call cannot be cut into slices afterwards.
 */
export class Sheet {
  /** Cumulative counts, in the manner of Gesso's `LayoutEngine.stats`. */
  readonly stats = { evaluated: 0, planned: 0, plans: 0 };

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
   *
   * Nothing in the graph would ever wake them — `=TODAY()` reads no
   * cell — so they are woken by hand on every edit. Without this the
   * date on a sheet is the date it was opened, for as long as it
   * stays open.
   */
  private readonly volatile = new Set<number>();

  /**
   * Formulas whose precedents cannot be read off their own text.
   *
   * `=INDIRECT("A" & B1)` reads a cell that the formula never names,
   * so the graph built from the tree has an edge to B1 and none to
   * the cell it actually fetched. These are evaluated with their
   * reads recorded and their edges rebuilt afterwards; see
   * `evaluateCell`.
   */
  private readonly dynamic = new Set<number>();

  /** The keys the cell now being evaluated read, when it is dynamic. */
  private recording: Set<number> | null = null;

  /**
   * Dynamic cells already sent round a second time since the last
   * edit.
   *
   * A dynamic cell that read something still dirty has to be redone
   * once that cell is settled. Once is enough and once is the limit:
   * two formulas pointing at each other through `INDIRECT` could
   * otherwise re-dirty each other for as long as anybody watched.
   */
  private readonly redone = new Set<number>();

  /** Fixed for a whole recalculation, so two `NOW()`s agree. */
  private moment: FunctionContext | null = null;

  /**
   * How far down anything has ever been written.
   *
   * A high-water mark, and deliberately not recomputed when cells are
   * cleared: shrinking it would mean walking the store on every
   * delete, and the only cost of it being too large is that a
   * whole-column reference reads a few cells that are empty. Too
   * small would drop data, which is why it never goes down.
   */
  private highWaterRow = 0;

  // ---------------------------------------------------------------------
  // Editing
  // ---------------------------------------------------------------------

  /**
   * Writes what someone typed.
   *
   * A leading `=` makes it a formula. Anything else is a literal, and
   * a literal that looks like a number is one — typing `3` gives the
   * number three, not the text "3", which is the behaviour everybody
   * expects and the reason a store of strings would be wrong.
   *
   * Nothing is evaluated here. The cell's own value is settled if it
   * is a literal, its formula is parsed if it is one, and everything
   * that reads it is marked for `recalculate`.
   *
   * `asText` is the Text number format reaching back into parsing,
   * and it is the only place a format touches a value. A cell
   * formatted as Text holds `007` as the three characters somebody
   * typed rather than the number seven, which is the whole reason
   * that format exists; a Text format that changed only the display
   * would be a menu item that does nothing anybody wanted. The
   * decision is the document's, because the document is what owns
   * both halves — see `SheetDocument.setCell`.
   */
  setCell(row: number, column: number, input: string, asText = false): void {
    const key = cellKey(row, column);
    if (input === '') {
      this.clearCell(row, column);
      return;
    }
    this.highWaterRow = Math.max(this.highWaterRow, row + 1);

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

  /**
   * Inserts or deletes whole rows or columns.
   *
   * Two things happen, and only one of them is obvious. The cells on
   * the moved side of the line are **carried** to their new keys —
   * that is the obvious half, and it is a rebuild of the store rather
   * than an in-place shuffle, because moving a cell to a key that
   * another cell has not vacated yet is how an in-place version
   * corrupts itself.
   *
   * The other half is that **every formula in the sheet is offered
   * the shift**, wherever it lives. A reference is a position, not an
   * offset: a formula in row 1 reading `=A900` names the cell at
   * A900, and after a row is inserted at 500 that cell is A901. So
   * the walk is over the whole store and the rewrites are the
   * formulas that actually mention the line — which is the number
   * `Structure.budget.spec.ts` counts.
   *
   * Returns how many formulas were rewritten, for that spec and for
   * nobody else.
   */
  shift(shift: Shift): number {
    const carried = new Map<number, Cell>();
    let rewrites = 0;

    for (const [key, cell] of this.cells) {
      const row = rowOf(key);
      const column = columnOf(key);
      const index = shift.axis === 'row' ? row : column;
      const moved = shiftIndex(index, shift);
      if (moved === -1) {
        // The cell was in a deleted row. It goes, and everything that
        // referenced it will say `#REF!` after the rewrite below.
        continue;
      }
      const input = shiftFormula(cell.input, shift);
      if (input !== cell.input) {
        rewrites++;
      }
      const at = moved === index ? key : shift.axis === 'row' ? cellKey(moved, column) : cellKey(row, moved);
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
    // Everything that reads a moved cell has to be redone, and after
    // a shift that is anything with a formula at all: the cheap
    // closure is the right one here.
    for (const [key, cell] of this.cells) {
      if (cell.formula !== null) {
        this.dirty.add(key);
      }
      // An insert pushes cells past the old mark, and a whole-column
      // reference that stopped at it would stop reading them.
      this.highWaterRow = Math.max(this.highWaterRow, rowOf(key) + 1);
    }
    return rewrites;
  }

  clearCell(row: number, column: number): void {
    const key = cellKey(row, column);
    this.graph.clearPrecedents(key);
    this.cells.delete(key);
    this.dirty.delete(key);
    this.forget(key);
    this.markDependentsDirty(key);
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

  /**
   * Files a formula under the two things the graph cannot see.
   *
   * Done once, when the formula is parsed. Walking the tree on every
   * recalculation would give the same answer at a cost per evaluation
   * instead of a cost per edit.
   */
  /** Puts a formula's edges into the graph, both kinds. */
  private wire(key: number, formula: Ast): void {
    const { keys, columns } = precedentsOf(formula);
    this.graph.setPrecedents(key, keys);
    this.graph.setWatchedColumns(key, columns);
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
    // edit was. This is what Excel means by the word, and the reason
    // the set is kept to four functions: the cost is paid per edit by
    // every volatile cell in the sheet.
    //
    // And by everything downstream of them. A `=TODAY()` that is
    // redone while the `=A1+1` beside it is not leaves two cells
    // disagreeing about the date — which is worse than either being
    // stale, because one of them looks right.
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

  /** Cells still waiting to be evaluated. */
  get pending(): number {
    return this.dirty.size;
  }

  /**
   * Evaluates at most `budget` cells, in an order where every cell
   * comes after the cells it reads.
   *
   * Called with no budget it runs to completion, which is what a test
   * wants. Called with one it is resumable: the plan and the position
   * in it survive between calls, and an edit arriving in between drops
   * the plan and rebuilds it from what is still dirty — the cells
   * already evaluated were evaluated correctly and are not redone
   * unless the new edit reaches them.
   */
  recalculate(budget: number = Number.POSITIVE_INFINITY): RecalcResult {
    let evaluated = 0;
    /**
     * A cap on how many times one call will rebuild its plan.
     *
     * New work can appear *during* a pass: a dynamic formula that read
     * a cell which had not been evaluated yet asks to be redone once
     * that cell is settled. The `redone` set already bounds that to
     * one retry per cell per edit, so this should never be reached —
     * it is here because a recalculation that spun forever would hang
     * the application worker, and a wrong answer is recoverable where
     * a hang is not.
     */
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
        // The budget ran out with the plan half done; it resumes on
        // the next call from exactly here.
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

  /**
   * Orders the dirty set, and settles the cells that cannot be
   * ordered.
   *
   * Circular cells are given `#CIRC!` here rather than during
   * evaluation because a cycle has no evaluation order by definition —
   * there is no point at which the answer could be computed — and they
   * leave the dirty set at the same moment, so they are not counted
   * as evaluations. A budget spec that counted them would be
   * measuring the cycle detector, not the recalc.
   */
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
    if (!this.dynamic.has(key)) {
      cell.value = evaluate(cell.formula, this);
      return;
    }

    /**
     * A formula whose edges have to be discovered by running it.
     *
     * The reads are recorded and the graph is rebuilt from them, so
     * that editing the cell an `INDIRECT` landed on wakes the formula
     * that read it. Without this the formula is correct once and stale
     * forever after — the worst bug a spreadsheet can have, because
     * nothing on the screen says so.
     */
    const reads = new Set<number>();
    this.recording = reads;
    try {
      cell.value = evaluate(cell.formula, this);
    } finally {
      this.recording = null;
    }
    if (!sameKeys(this.graph.precedentsOf(key), reads)) {
      this.graph.setPrecedents(key, reads);
    }
    // It has just read cells that are themselves still waiting, so the
    // answer it gave may be one recalculation out of date. Doing it
    // again once they are settled is the fix; doing it again at most
    // once is what stops two of these chasing each other.
    if (!this.redone.has(key) && anyDirty(reads, this.dirty)) {
      this.redone.add(key);
      this.dirty.add(key);
    }
  }

  // ---------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------

  /** The `EvaluationContext` the evaluator reads through. */
  valueAt(key: number): CellValue {
    // Recorded only while a dynamic formula is running, which is the
    // one case where what was read is not what the tree said.
    this.recording?.add(key);
    return this.cells.get(key)?.value ?? null;
  }

  /**
   * The clock and dice half of `EvaluationContext`.
   *
   * The time is *read once* and handed back as a constant, which is
   * the whole reason this is cached rather than passed straight
   * through. Two `NOW()`s in one recalculation that each called the
   * clock would disagree by however long the pass took, and a sheet
   * that contradicts itself about its own timestamps is worse than
   * one that is a moment out of date.
   */
  get functions(): FunctionContext {
    if (this.moment === null) {
      const at = this.clock();
      this.moment = { now: () => at, random: () => this.dice() };
    }
    return this.moment;
  }

  /** How far a whole-column reference reads; see `highWaterRow`. */
  get usedRows(): number {
    return this.highWaterRow;
  }

  value(row: number, column: number): CellValue {
    return this.valueAt(cellKey(row, column));
  }

  /** What was typed, which is what a cell editor opens with. */
  input(row: number, column: number): string {
    return this.cells.get(cellKey(row, column))?.input ?? '';
  }

  /**
   * The value as general text.
   *
   * **Not what the screen shows** — that was true until Phase 9 and
   * is not any more. A cell's display string depends on its number
   * format as well as its value, and a format is not the sheet's; ask
   * `SheetDocument.display` for what a person sees. This is the
   * unformatted answer, which is what a `General` cell shows and what
   * everything comparing values as text wants.
   */
  display(row: number, column: number): string {
    return formatValue(this.value(row, column));
  }

  /** Cells that hold something, in no order. */
  get size(): number {
    return this.cells.size;
  }

  /** Every non-empty cell, for a repository to write out. */
  *entries(): Generator<{ row: number; column: number; input: string }> {
    for (const [key, cell] of this.cells) {
      yield { row: rowOf(key), column: columnOf(key), input: cell.input };
    }
  }

  /** What a cell reads, for `engine.explain` in Phase 7. */
  precedentsOf(row: number, column: number): number[] {
    return [...this.graph.precedentsOf(cellKey(row, column))];
  }

  /** What reads a cell. */
  dependentsOf(row: number, column: number): number[] {
    return [...this.graph.dependentsOf(cellKey(row, column))];
  }
}

/**
 * What a formula reads: cells by key, and columns by number.
 *
 * Two lists because a whole-column reference cannot be a key list —
 * `A:A` is 1,048,576 cells — and must not be turned into one. The
 * graph watches those columns instead; see `DependencyGraph`.
 */
function precedentsOf(formula: Ast): { keys: number[]; columns: number[] } {
  const keys: number[] = [];
  const columns: number[] = [];
  referencesOf(formula, {
    ref(ref: CellRef) {
      keys.push(cellKey(ref.row, ref.column));
    },
    range(range: RangeRef) {
      if (range.wholeColumn === true) {
        for (let column = range.start.column; column <= range.end.column; column++) {
          columns.push(column);
        }
        return;
      }
      for (const key of rangeKeys(range)) {
        keys.push(key);
      }
    }
  });
  return { keys, columns };
}

/**
 * What typing something that is not a formula means.
 *
 * Numbers, TRUE/FALSE and dates are recognised; everything else is
 * text. The number test is deliberately stricter than `Number()`,
 * which reads `''` as 0 and `'0x10'` as 16 — neither of which is what
 * someone who typed them meant.
 *
 * A date becomes its serial number, because in a spreadsheet that is
 * what a date *is* — see `Dates.ts`. The number test runs first, so
 * `2026` stays the number two thousand and twenty-six rather than
 * becoming a year.
 *
 * The format that makes the serial legible is the document's to
 * apply, and `SheetDocument.setCell` applies it in the same undo step.
 * The two have to agree: a serial written here without a format there
 * is a cell showing 46,289 to somebody who typed a date.
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
