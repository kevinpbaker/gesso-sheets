import { describe, expect, it } from 'vitest';

import { serialOfDate } from './Dates';
import { Sheet } from './Sheet';
import { Workbook } from './Workbook';
import { formatValue } from './Values';

/**
 * Phase 11's other exit criterion, and the three hazards it named.
 *
 * The budget is stated in **evaluations**, not milliseconds, for the
 * reason every budget in this repository is: a count is the same on
 * every machine and a timing is a description of this one. A
 * `VLOOKUP` down a full column must be evaluated once per edit inside
 * that column and not once per row, and the number is asserted
 * exactly rather than with `toBeLessThan` — a ceiling passes for a
 * regression that stays under it.
 */

const ROWS = 5_000;

/** A column of numbers with a lookup table beside it. */
function bigSheet(): Sheet {
  const sheet = new Workbook().sheet(0);
  for (let row = 0; row < ROWS; row++) {
    sheet.setCell(row, 0, String(row));
    sheet.setCell(row, 1, `name-${row}`);
  }
  sheet.recalculate();
  return sheet;
}

describe('a full-column reference', () => {
  it('stores no edge per cell', () => {
    const sheet = bigSheet();
    sheet.setCell(0, 5, '=SUM(A:A)');
    sheet.recalculate();

    // The formula's *recorded* precedents are the cells it names, and
    // it names none: the column is watched instead. A version that
    // expanded the range would have five thousand here, and a full
    // sheet would have a million.
    expect(sheet.precedentsOf(0, 5)).toEqual([]);
    expect(formatValue(sheet.value(0, 5))).toBe(String((ROWS * (ROWS - 1)) / 2));
  });

  /**
   * The thing the watch exists for. Without it the sum is correct
   * once and stale forever, which is the silent failure the roadmap
   * named as the worst a spreadsheet can have.
   */
  it('is woken by a write anywhere in the column', () => {
    const sheet = bigSheet();
    sheet.setCell(0, 5, '=SUM(A:A)');
    sheet.recalculate();
    const before = sheet.value(0, 5) as number;

    sheet.setCell(ROWS + 10, 0, '1000');
    sheet.recalculate();

    expect(sheet.value(0, 5)).toBe(before + 1000);
  });

  it('is not woken by a write to another column', () => {
    const sheet = bigSheet();
    sheet.setCell(0, 5, '=SUM(A:A)');
    sheet.recalculate();

    const before = sheet.stats.evaluated;
    sheet.setCell(3, 3, '99');
    sheet.recalculate();

    // Nothing to redo: the write was outside the watched column.
    expect(sheet.stats.evaluated - before).toBe(0);
  });

  it('reads only as far as the sheet has ever been written', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '5');
    sheet.setCell(1, 0, '7');
    sheet.setCell(0, 5, '=COUNT(A:A)');
    sheet.recalculate();

    // Two cells, not 1,048,576 blanks.
    expect(sheet.value(0, 5)).toBe(2);
  });

  it('is empty in a column of a sheet with nothing in it', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 5, '=COUNT(B:B)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(0);
  });
});

describe('the budget a lookup down a full column has', () => {
  /**
   * The exit criterion, stated exactly.
   *
   * One edit inside the column, one evaluation of the formula that
   * reads it. Five thousand would mean the range had become an edge
   * per row; two would mean something was being done twice.
   */
  it('evaluates once per edit, not once per row', () => {
    const sheet = bigSheet();
    sheet.setCell(0, 5, '=VLOOKUP(4000, A:B, 2, FALSE)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe('name-4000');

    const before = sheet.stats.evaluated;
    sheet.setCell(2_500, 0, '-1');
    sheet.recalculate();

    expect(sheet.stats.evaluated - before).toBe(1);
  });

  /** And a thousand edits are a thousand evaluations, not a million. */
  it('evaluates once per edit when there are many', () => {
    const sheet = bigSheet();
    sheet.setCell(0, 5, '=VLOOKUP(4000, A:B, 2, FALSE)');
    sheet.recalculate();

    const before = sheet.stats.evaluated;
    for (let row = 0; row < 1_000; row++) {
      sheet.setCell(row, 0, String(row + 1));
      sheet.recalculate();
    }

    expect(sheet.stats.evaluated - before).toBe(1_000);
  });
});

/**
 * The volatile set: functions that answer differently when nothing
 * they read has changed.
 */
describe('a volatile formula', () => {
  function sheetAt(day: number): Sheet {
    const sheet = new Workbook().sheet(0);
    sheet.clock = () => day;
    return sheet;
  }

  it('is redone when an unrelated cell is edited', () => {
    const sheet = sheetAt(serialOfDate(2026, 9, 24));
    sheet.setCell(0, 0, '=TODAY()');
    sheet.recalculate();
    expect(sheet.value(0, 0)).toBe(46_289);

    // The next day, and an edit somewhere else entirely.
    sheet.clock = () => serialOfDate(2026, 9, 25);
    sheet.setCell(9, 9, 'anything');
    sheet.recalculate();

    expect(sheet.value(0, 0)).toBe(46_290);
  });

  it('carries everything downstream of it along', () => {
    const sheet = sheetAt(serialOfDate(2026, 9, 24));
    sheet.setCell(0, 0, '=TODAY()');
    sheet.setCell(1, 0, '=A1+1');
    sheet.recalculate();
    expect(sheet.value(1, 0)).toBe(46_290);

    sheet.clock = () => serialOfDate(2026, 9, 25);
    sheet.setCell(9, 9, 'anything');
    sheet.recalculate();

    expect(sheet.value(1, 0)).toBe(46_291);
  });

  it('stops being volatile when the formula is replaced', () => {
    const sheet = sheetAt(serialOfDate(2026, 9, 24));
    sheet.setCell(0, 0, '=TODAY()');
    sheet.recalculate();
    sheet.setCell(0, 0, '=1+1');
    sheet.recalculate();

    const before = sheet.stats.evaluated;
    sheet.setCell(9, 9, 'anything');
    sheet.recalculate();
    expect(sheet.stats.evaluated - before).toBe(0);
  });

  /** A sheet with no volatile cell pays nothing for the mechanism. */
  it('costs nothing when there is none', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '=1+1');
    sheet.recalculate();

    const before = sheet.stats.evaluated;
    sheet.setCell(9, 9, 'anything');
    sheet.recalculate();
    expect(sheet.stats.evaluated - before).toBe(0);
  });

  /** Two `NOW()`s in one pass must not disagree about the time. */
  it('gives the same answer to every cell in one recalculation', () => {
    let ticks = 0;
    const sheet = new Workbook().sheet(0);
    sheet.clock = () => 46_289 + ticks++ / 1000;
    sheet.setCell(0, 0, '=NOW()');
    sheet.setCell(1, 0, '=NOW()');
    sheet.recalculate();

    expect(sheet.value(0, 0)).toBe(sheet.value(1, 0));
  });
});

/**
 * The dangerous pair: references computed while the formula runs.
 *
 * The roadmap named the failure mode — "a cell that is stale and
 * never woken, which is the worst bug a spreadsheet can have because
 * it is silent" — so these are the specs that say it does not happen.
 */
describe('a formula whose references are computed', () => {
  it('reads the cell its text never names', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '3');
    sheet.setCell(2, 0, '30');
    sheet.setCell(0, 5, '=INDIRECT("A" & A1)');
    sheet.recalculate();

    expect(sheet.value(0, 5)).toBe(30);
  });

  /** The whole point: editing the cell it landed on wakes it. */
  it('is woken by an edit to the cell it landed on', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '3');
    sheet.setCell(2, 0, '30');
    sheet.setCell(0, 5, '=INDIRECT("A" & A1)');
    sheet.recalculate();

    sheet.setCell(2, 0, '99');
    sheet.recalculate();

    expect(sheet.value(0, 5)).toBe(99);
  });

  /** And moving it re-points the edge, rather than keeping the old one. */
  it('follows when the reference it computes moves', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '3');
    sheet.setCell(2, 0, '30');
    sheet.setCell(3, 0, '40');
    sheet.setCell(0, 5, '=INDIRECT("A" & A1)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(30);

    // Point it at A4 instead, then edit A4.
    sheet.setCell(0, 0, '4');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(40);

    sheet.setCell(3, 0, '41');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(41);

    // And the cell it used to read no longer wakes it.
    const before = sheet.stats.evaluated;
    sheet.setCell(2, 0, '31');
    sheet.recalculate();
    expect(sheet.stats.evaluated - before).toBe(0);
  });

  it('does the same for OFFSET', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '1');
    sheet.setCell(1, 0, '2');
    sheet.setCell(2, 0, '3');
    sheet.setCell(0, 5, '=SUM(OFFSET(A1, 0, 0, 3, 1))');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(6);

    sheet.setCell(1, 0, '20');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(24);
  });

  /**
   * The ordering case: the computed reference points at a cell that is
   * itself waiting to be recalculated.
   */
  it('is correct when it reads a cell that is itself dirty', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '3');
    sheet.setCell(1, 0, '10');
    // A3 is a formula, and the INDIRECT lands on it.
    sheet.setCell(2, 0, '=A2*2');
    sheet.setCell(0, 5, '=INDIRECT("A" & A1)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(20);

    sheet.setCell(1, 0, '50');
    sheet.recalculate();

    expect(sheet.value(2, 0)).toBe(100);
    expect(sheet.value(0, 5)).toBe(100);
  });

  /** Two of them pointing at each other must settle rather than spin. */
  it('settles when two of them point at each other', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '2');
    sheet.setCell(1, 0, '1');
    sheet.setCell(0, 5, '=INDIRECT("F" & A1)');
    sheet.setCell(1, 5, '=INDIRECT("F" & A2)');
    sheet.recalculate();

    // Whatever they settle on, the recalculation finished.
    expect(sheet.pending).toBe(0);
  });

  it('stops being dynamic when the formula is replaced', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '3');
    sheet.setCell(2, 0, '30');
    sheet.setCell(0, 5, '=INDIRECT("A" & A1)');
    sheet.recalculate();
    sheet.setCell(0, 5, '=1+1');
    sheet.recalculate();

    const before = sheet.stats.evaluated;
    sheet.setCell(2, 0, '99');
    sheet.recalculate();
    expect(sheet.stats.evaluated - before).toBe(0);
  });
});
