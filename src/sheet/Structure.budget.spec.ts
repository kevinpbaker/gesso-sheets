import { describe, expect, it } from 'vitest';

import { columnName } from './A1';
import { Sheet } from './Sheet';
import { Workbook } from './Workbook';

/**
 * Structural budgets — the exit criterion for Phase 10.
 *
 * "Inserting a row above a column of 50,000 formulas rewrites exactly
 * the formulas that reference it, `toBe` and not `toBeLessThan`."
 *
 * A count rather than a timing, for the reason every budget in this
 * repository is a count: it fails the build with a number when a
 * change makes an insert touch the sheet instead of the references.
 * It is the same shape as `Sheet.budget.spec.ts`, which counts
 * evaluations, and it is measuring the other half of the same claim —
 * that the work an edit causes is proportional to what actually
 * depends on it.
 */
describe('what an insert rewrites', () => {
  /** A column of `count` formulas, each reading the cell above it. */
  function chain(count: number): Sheet {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '1');
    for (let row = 1; row < count; row++) {
      sheet.setCell(row, 0, `=A${row}`);
    }
    sheet.recalculate();
    return sheet;
  }

  /**
   * The claim, in one number.
   *
   * Fifty thousand formulas, a row inserted at the top, and every one
   * of them references a row at or below the line — so every one is
   * rewritten, and the count says *exactly* that rather than "at
   * most".
   */
  it('rewrites every formula that references the line', () => {
    const CELLS = 50_000;
    const sheet = chain(CELLS);
    // 49,999 formulas: A1 is the literal head.
    expect(sheet.shift({ axis: 'row', at: 0, by: 1 })).toBe(CELLS - 1);
  });

  /**
   * The same claim stated the other way round, which is the version
   * that catches the regression the first one would not: a row
   * inserted *below* everything rewrites nothing at all, however many
   * formulas the sheet holds.
   */
  it('rewrites nothing when nothing references the line', () => {
    const sheet = chain(50_000);
    expect(sheet.shift({ axis: 'row', at: 60_000, by: 1 })).toBe(0);
  });

  /** And the boundary: only the formulas below the line, not above. */
  it('rewrites exactly the formulas below the line and no others', () => {
    const sheet = chain(1_000);
    // Inserting at row 500 moves rows 500 and below. A formula in row
    // r reads row r-1, so the formulas that mention a moved row are
    // the ones from row 501 onwards — 499 of them — plus the one in
    // row 500 is *itself* moved but reads row 499, which did not
    // move, so it is not rewritten.
    expect(sheet.shift({ axis: 'row', at: 500, by: 1 })).toBe(499);
  });

  /**
   * A shift on the other axis rewrites nothing in a column of row
   * references. The two axes are independent and a shift that moved
   * both would be a shift that corrupted one.
   */
  it('rewrites nothing on the axis it did not move', () => {
    const sheet = chain(1_000);
    expect(sheet.shift({ axis: 'column', at: 5, by: 1 })).toBe(0);
  });

  /**
   * A range is one rewrite however many cells it covers — the
   * counterpart of Phase 1's "an edit inside a SUM range evaluates
   * the formula once and not once per cell".
   */
  it('counts a range as one formula, not one per cell', () => {
    const sheet = new Workbook().sheet(0);
    for (let row = 0; row < 10_000; row++) {
      sheet.setCell(row, 0, String(row));
    }
    sheet.setCell(0, 1, '=SUM(A1:A10000)');
    sheet.recalculate();

    expect(sheet.shift({ axis: 'row', at: 0, by: 1 })).toBe(1);
  });

  /** The sheet still agrees with itself afterwards. */
  it('leaves a sheet that recalculates to the same answers', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '10');
    sheet.setCell(1, 0, '20');
    sheet.setCell(2, 0, '=SUM(A1:A2)');
    sheet.recalculate();
    expect(sheet.value(2, 0)).toBe(30);

    sheet.shift({ axis: 'row', at: 1, by: 1 });
    sheet.recalculate();

    // A row went in between them; the sum grew to cover it and still
    // adds the same two numbers.
    expect(sheet.input(3, 0)).toBe('=SUM(A1:A3)');
    expect(sheet.value(3, 0)).toBe(30);
    expect(sheet.value(0, 0)).toBe(10);
    expect(sheet.value(2, 0)).toBe(20);
  });

  it('breaks what pointed at a row it deleted, and only that', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '10');
    sheet.setCell(1, 0, '20');
    sheet.setCell(2, 0, '=A1');
    sheet.setCell(3, 0, '=A2');
    sheet.recalculate();

    sheet.shift({ axis: 'row', at: 1, by: -1 });
    sheet.recalculate();

    // A2 is gone. What read A1 still reads A1; what read A2 says so.
    expect(sheet.input(1, 0)).toBe('=A1');
    expect(sheet.value(1, 0)).toBe(10);
    expect(sheet.input(2, 0)).toBe('=#REF!');
    expect(sheet.value(2, 0)).toEqual({ kind: 'error', code: '#REF!' });
  });

  /** Nothing is left pointing at a cell that used to be somewhere else. */
  it('leaves no stale edges behind', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '1');
    sheet.setCell(1, 0, '=A1*2');
    sheet.recalculate();

    sheet.shift({ axis: 'row', at: 0, by: 1 });
    sheet.recalculate();
    expect(sheet.value(2, 0)).toBe(2);

    // The head moved to A2; editing it has to wake the formula that
    // now reads it, and editing the empty A1 must wake nothing.
    sheet.setCell(1, 0, '5');
    sheet.recalculate();
    expect(sheet.value(2, 0)).toBe(10);
  });
});

/** The proof surface's chain, for a sense of what an insert costs. */
describe('how far an insert reaches', () => {
  it('walks the cells that exist, not the cells that could', () => {
    const sheet = new Workbook().sheet(0);
    // Four cells in a sheet a million rows tall.
    sheet.setCell(0, 0, '1');
    sheet.setCell(500_000, 0, '=A1');
    sheet.setCell(999_999, 3, '=A500001');
    sheet.recalculate();

    expect(sheet.size).toBe(3);
    // Two formulas reference rows at or below the line; both move.
    expect(sheet.shift({ axis: 'row', at: 0, by: 1 })).toBe(2);
    expect(sheet.size).toBe(3);
  });
});
