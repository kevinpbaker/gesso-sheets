import { describe, expect, it } from 'vitest';

import type { Sheet } from './Sheet';
import { Workbook } from './Workbook';

/**
 * Phase 28's budget: a `LAMBDA` is not a tax on the sheet that uses it.
 *
 * A hundred thousand rows each calling a named `LAMBDA` against the same
 * hundred thousand written out inline. Both read one shared cell, so an
 * edit to it recalculates every row, and that recalculation is what is
 * timed — the parse is paid once, when the formulas are typed, and is
 * not the claim.
 *
 * A ratio rather than a number of milliseconds, because it is the same
 * machine on both sides of it. The count of evaluations is asserted
 * exactly as well: a function call is one evaluation of the cell that
 * makes it, and not one more.
 */

const ROWS = 100_000;
/** How much slower the function may be than the arithmetic it wraps. */
const FACTOR = 3;

function column(formula: (line: number) => string, withName: boolean): Sheet {
  const sheet = new Workbook().sheet(0);
  sheet.setCell(0, 2, '1.2');
  if (withName) {
    sheet.names.defineFormula('Scaled', '=LAMBDA(x, x * $C$1 + 1)');
    sheet.namesChanged();
  }
  for (let row = 0; row < ROWS; row++) {
    sheet.setCell(row, 0, String(row));
    sheet.setCell(row, 1, formula(row + 1));
  }
  sheet.recalculate();
  return sheet;
}

/** The best of a few recalculations of every row, woken by the shared cell. */
function timed(sheet: Sheet): { ms: number; evaluated: number } {
  let best = Infinity;
  let evaluated = 0;
  for (let round = 0; round < 3; round++) {
    sheet.setCell(0, 2, String(1.2 + round));
    const before = sheet.stats.evaluated;
    const start = performance.now();
    sheet.recalculate();
    best = Math.min(best, performance.now() - start);
    evaluated = sheet.stats.evaluated - before;
  }
  return { ms: best, evaluated };
}

describe('a LAMBDA down a column of 100,000', () => {
  it('costs within a small factor of the same arithmetic inline', () => {
    const inline = column(line => `=A${line} * $C$1 + 1`, false);
    const called = column(line => `=Scaled(A${line})`, true);
    expect(called.value(ROWS - 1, 1)).toBe(inline.value(ROWS - 1, 1));

    const plain = timed(inline);
    const through = timed(called);
    console.info(
      `[lambda budget] inline ${plain.ms.toFixed(1)} ms, through a LAMBDA ${through.ms.toFixed(1)} ms` +
        ` (${(through.ms / plain.ms).toFixed(2)}×)`
    );
    // Every row once, and nothing else: the shared cell's own edit is a
    // value, not a formula, and is not counted.
    expect(through.evaluated).toBe(ROWS);
    expect(plain.evaluated).toBe(ROWS);
    expect(through.ms).toBeLessThan(plain.ms * FACTOR);
  }, 60_000);
});
