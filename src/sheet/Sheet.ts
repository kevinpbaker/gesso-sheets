import { cellKey, columnOf, rangeKeys, rowOf, type CellRef, type RangeRef } from './A1';
import { referencesOf, type Ast } from './Ast';
import { DependencyGraph } from './DependencyGraph';
import { evaluate } from './Evaluator';
import { FormulaSyntaxError, parseFormula } from './Parser';
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

    if (input.startsWith('=') && !asText) {
      this.writeFormula(key, input);
    } else {
      this.graph.clearPrecedents(key);
      this.dirty.delete(key);
      this.cells.set(key, { input, formula: null, value: asText ? input : literalValue(input) });
    }
    this.markDependentsDirty(key);
  }

  clearCell(row: number, column: number): void {
    const key = cellKey(row, column);
    this.graph.clearPrecedents(key);
    this.cells.delete(key);
    this.dirty.delete(key);
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
      this.cells.set(key, { input, formula: null, value: VALUE });
      return;
    }
    this.cells.set(key, { input, formula, value: null });
    this.graph.setPrecedents(key, precedentsOf(formula));
    this.dirty.add(key);
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
    if (this.dirty.size === 0) {
      return { evaluated: 0, done: true };
    }
    if (this.plan === null) {
      this.plan = this.buildPlan();
    }

    let evaluated = 0;
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

    const done = plan.at >= plan.order.length;
    if (done) {
      this.plan = null;
    }
    this.stats.evaluated += evaluated;
    return { evaluated, done: done && this.dirty.size === 0 };
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
    cell.value = evaluate(cell.formula, this);
  }

  // ---------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------

  /** The `EvaluationContext` the evaluator reads through. */
  valueAt(key: number): CellValue {
    return this.cells.get(key)?.value ?? null;
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

/** Every cell a formula reads, ranges expanded. */
function precedentsOf(formula: Ast): number[] {
  const keys: number[] = [];
  referencesOf(formula, {
    ref(ref: CellRef) {
      keys.push(cellKey(ref.row, ref.column));
    },
    range(range: RangeRef) {
      for (const key of rangeKeys(range)) {
        keys.push(key);
      }
    }
  });
  return keys;
}

/**
 * What typing something that is not a formula means.
 *
 * Numbers and TRUE/FALSE are recognised; everything else is text. The
 * number test is deliberately stricter than `Number()`, which reads
 * `''` as 0 and `'0x10'` as 16 — neither of which is what someone who
 * typed them meant.
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
  return input;
}
