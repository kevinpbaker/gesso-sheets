import {
  covers,
  isEmptyPaint,
  matches,
  overlay,
  scaleColour,
  type ConditionalPaint,
  type ConditionalRule,
  type Extent
} from '../sheet/Conditional';
import type { Ast } from '../sheet/Ast';
import { evaluate } from '../sheet/Evaluator';
import { FormulaSyntaxError, parseFormula } from '../sheet/Parser';
import { shiftAstBy } from '../sheet/Rewrite';
import { toBoolean, type CellValue } from '../sheet/Values';
import type { Sheet } from '../sheet/Sheet';

/**
 * A sheet's conditional rules, resolved for the cells in a window.
 *
 * This is where the phase's claim is kept. A rule covering a million
 * cells is asked only about the cells somebody can see, so a rule
 * over the whole sheet costs the viewport — the sentence this
 * project keeps writing.
 *
 * Two things are not free and both are bounded here.
 *
 * **A colour scale needs the extent of its range**, which is not a
 * question about one cell. It is computed by walking the cells that
 * *hold something* and keeping the ones inside the range — bounded
 * by how much has been typed rather than by how large the range is,
 * the same trick `retextRegion` uses — and then cached until the
 * sheet changes underneath it. A scroll does not invalidate it,
 * which is what the exit criterion turns on.
 *
 * **A formula rule is parsed once and shifted per cell.** Through
 * the text that would be a parse per cell per publish; through the
 * tree it is an allocation per node. Both numbers are asserted in
 * `Conditional.budget.spec.ts` rather than assumed, because the
 * roadmap asked for a measurement and not a prediction.
 *
 * A sheet with no rules costs nothing at all: every entry point
 * leaves immediately, so the work this file does is invisible to
 * `pnpm proof`, which measures a sheet that has none.
 */
export class ConditionalPainter {
  /** Counted for the budget spec, and for nothing else. */
  readonly stats = { scans: 0, scanned: 0, evaluations: 0 };

  private rules: readonly ConditionalRule[] = [];
  /** Parsed formula rules, by rule index. Undefined when not one. */
  private trees: (Ast | null | undefined)[] = [];
  /** A scale's extent, by rule index, or undefined when not worked out. */
  private extents: (Extent | null | undefined)[] = [];

  constructor(private readonly sheetOf: () => Sheet) {}

  get isEmpty(): boolean {
    return this.rules.length === 0;
  }

  /**
   * The rules changed, so everything derived from them is stale.
   *
   * Separate from `invalidate` because the trees survive an edit and
   * do not survive a rule being rewritten.
   */
  setRules(rules: readonly ConditionalRule[]): void {
    this.rules = rules;
    this.trees = rules.map(() => undefined);
    this.extents = rules.map(() => undefined);
  }

  /**
   * The sheet changed, so the extents may have.
   *
   * Called on any edit rather than on the edits that could actually
   * move a minimum or a maximum, because working out which those are
   * means reading the range — which is the thing being avoided.
   */
  invalidate(): void {
    if (this.rules.length > 0) {
      this.extents = this.rules.map(() => undefined);
    }
  }

  /**
   * What a cell is painted over its own format, or null.
   *
   * Null rather than an empty object, so the caller can tell "no rule
   * touched this" from "a rule touched it and said nothing" without
   * looking inside.
   */
  paintFor(row: number, column: number, value: CellValue): ConditionalPaint | null {
    if (this.rules.length === 0) {
      return null;
    }
    const found: ConditionalPaint[] = [];
    for (let at = 0; at < this.rules.length; at++) {
      const rule = this.rules[at];
      if (!covers(rule.range, row, column)) {
        continue;
      }
      if (rule.scale !== undefined) {
        const colour = scaleColour(rule.scale, value, this.extentOf(at, rule));
        if (colour !== null) {
          found.push({ fill: colour });
        }
        continue;
      }
      if (rule.test === null || rule.paint === undefined) {
        continue;
      }
      const plain = matches(rule.test, value);
      const hit = plain ?? this.formulaHolds(at, rule, row, column);
      if (hit) {
        found.push(rule.paint);
      }
    }
    if (found.length === 0) {
      return null;
    }
    const folded = overlay(found);
    return isEmptyPaint(folded) ? null : folded;
  }

  /**
   * A formula rule, moved to this cell and asked.
   *
   * The rule is written as it would be for the range's first cell —
   * `=A1<0` over `A1:D9` means "this cell is negative" — and shifted
   * by how far the cell is from that corner, which is the same rule a
   * fill goes by and is what people already know.
   */
  private formulaHolds(at: number, rule: ConditionalRule, row: number, column: number): boolean {
    const test = rule.test;
    if (test === null || test.kind !== 'formula') {
      return false;
    }
    let tree = this.trees[at];
    if (tree === undefined) {
      tree = parseOrNull(test.input);
      this.trees[at] = tree;
    }
    if (tree === null) {
      return false;
    }
    const sheet = this.sheetOf();
    const moved = shiftAstBy(
      tree,
      row - Math.min(rule.range.start.row, rule.range.end.row),
      column - Math.min(rule.range.start.column, rule.range.end.column)
    );
    this.stats.evaluations++;
    const answer = evaluate(moved, {
      valueAt: (key: number) => sheet.valueAt(key),
      usedRows: sheet.usedRows,
      rangeForName: (name: string) => sheet.rangeForName(name)
    });
    const truth = toBoolean(answer);
    return typeof truth === 'boolean' && truth;
  }

  /**
   * The smallest and largest number in a scale's range.
   *
   * Bounded by the cells that hold something rather than by the size
   * of the range, and cached until the sheet changes. A range with no
   * numbers in it has no extent, and the scale then says nothing —
   * which is better than colouring everything the same and calling it
   * a gradient.
   */
  private extentOf(at: number, rule: ConditionalRule): Extent {
    const held = this.extents[at];
    if (held !== undefined) {
      return held ?? { low: 0, high: 0 };
    }
    const sheet = this.sheetOf();
    let low = Number.POSITIVE_INFINITY;
    let high = Number.NEGATIVE_INFINITY;
    let scanned = 0;
    for (const cell of sheet.entries()) {
      scanned++;
      if (!covers(rule.range, cell.row, cell.column)) {
        continue;
      }
      const value = sheet.value(cell.row, cell.column);
      if (typeof value === 'number' && Number.isFinite(value)) {
        low = Math.min(low, value);
        high = Math.max(high, value);
      }
    }
    this.stats.scans++;
    this.stats.scanned += scanned;
    const found: Extent | null = low <= high ? { low, high } : null;
    this.extents[at] = found;
    return found ?? { low: 0, high: 0 };
  }

  /** Whether any rule covers a cell at all, for the caller's fast path. */
  touches(row: number, column: number): boolean {
    for (const rule of this.rules) {
      if (covers(rule.range, row, column)) {
        return true;
      }
    }
    return false;
  }
}

function parseOrNull(input: string): Ast | null {
  const text = input.startsWith('=') ? input.slice(1) : input;
  try {
    return parseFormula(text);
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    // A rule that does not parse paints nothing, for the same reason
    // a formula that does not parse holds `#VALUE!`: what somebody
    // wrote is kept, and it stops being wrong when they finish.
    return null;
  }
}
