import { describe, expect, it } from 'vitest';

import { MAX_COLUMNS, MAX_ROWS } from './A1';
import { Workbook } from './Workbook';

/**
 * Ctrl+Arrow across a gap, on a sheet as tall as Excel's — Phase 19.
 *
 * The claim is that finding the next filled cell costs what the sheet
 * holds rather than how far away the cell is. Measured against the
 * walk it replaces, a lookup per empty cell, on the same sheet in the
 * same run: a fixed number of milliseconds would be a claim about this
 * machine, and the ratio is a claim about the arithmetic.
 */
describe('the edge of the data', () => {
  const extent = { rowCount: MAX_ROWS, columnCount: MAX_COLUMNS };
  const LAST = MAX_ROWS - 1;

  /** A thousand-row table in B:K, and one cell at the foot of column A. */
  function sheet() {
    const book = new Workbook();
    const one = book.sheet(0);
    for (let row = 0; row < 1_000; row++) {
      for (let column = 1; column <= 10; column++) {
        one.setCell(row, column, String(row * column));
      }
    }
    one.setCell(LAST, 0, 'foot');
    one.recalculate();
    return one;
  }

  function timed(run: () => void, times = 5): number {
    let best = Infinity;
    for (let i = 0; i < times; i++) {
      const start = performance.now();
      run();
      best = Math.min(best, performance.now() - start);
    }
    return best;
  }

  it('finds a cell a million rows away without walking to it', () => {
    const one = sheet();

    let landed = { row: 0, column: 0 };
    const jump = timed(() => {
      landed = one.edgeFrom(0, 0, 1, 0, extent);
    });
    const walk = timed(() => {
      let row = 1;
      while (row < LAST && one.value(row, 0) === null) {
        row++;
      }
    });
    console.info(`[edge budget] jump ${jump.toFixed(2)} ms, walk ${walk.toFixed(2)} ms`);

    expect(landed).toEqual({ row: LAST, column: 0 });
    expect(jump * 5).toBeLessThan(walk);
  });

  it('lands on the edge of the sheet when there is nothing to find', () => {
    const one = sheet();
    expect(one.edgeFrom(0, 12, 1, 0, extent)).toEqual({ row: LAST, column: 12 });
    expect(one.edgeFrom(5, 12, 0, 1, extent)).toEqual({ row: 5, column: MAX_COLUMNS - 1 });
  });

  it('stays put at the edge it is already on', () => {
    const one = sheet();
    expect(one.edgeFrom(0, 0, -1, 0, extent)).toEqual({ row: 0, column: 0 });
  });

  it('runs to the end of a block, and from there to the next', () => {
    const one = sheet();
    expect(one.edgeFrom(0, 1, 1, 0, extent)).toEqual({ row: 999, column: 1 });
    expect(one.edgeFrom(0, 1, 0, 1, extent)).toEqual({ row: 0, column: 10 });
    expect(one.edgeFrom(0, 10, 0, -1, extent)).toEqual({ row: 0, column: 1 });
    expect(one.edgeFrom(0, 1, 0, -1, extent)).toEqual({ row: 0, column: 0 });
  });
});
