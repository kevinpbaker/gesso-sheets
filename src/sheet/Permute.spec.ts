import { describe, expect, it } from 'vitest';

import { permuteFormula, rowsRead } from './Permute';

/** Columns A..C, the span a sort of three columns would pick up. */
const columns = { first: 0, last: 2 };

/** Rows 1 and 3 swap; row 2 stays. */
const swap = new Map([
  [1, 3],
  [2, 2],
  [3, 1]
]);

describe('permuting a formula', () => {
  it('follows a reference whose cell the sort moved', () => {
    // B2 is row 1, column 1 — picked up and put down at row 3.
    expect(permuteFormula('=B2', swap, columns)).toBe('=B4');
  });

  it('leaves a reference to a row the sort did not touch', () => {
    // B1 is row 0, outside the permutation.
    expect(permuteFormula('=B1', swap, columns)).toBe('=B1');
  });

  /**
   * The bug a spec caught: a sort of columns A..C does not move column
   * E, however much the rows beside it were shuffled.
   */
  it('leaves a reference to a column the sort did not touch', () => {
    expect(permuteFormula('=E2', swap, columns)).toBe('=E2');
  });

  /**
   * `$` says "do not drift when I am copied". This is not a copy — the
   * cell genuinely moved, and a reference left behind would point at
   * whatever row took its place. `Shift` treats an insert the same way.
   */
  it('moves an absolute reference, because the cell moved', () => {
    expect(permuteFormula('=$B$2', swap, columns)).toBe('=$B$4');
  });

  it('leaves a range alone', () => {
    expect(permuteFormula('=SUM(B2:B4)', swap, columns)).toBe('=SUM(B2:B4)');
  });

  it('hands back text it did not have to change, character for character', () => {
    // Not laundered through the printer, which would parenthesise it.
    expect(permuteFormula('=B1 * 2', swap, columns)).toBe('=B1 * 2');
  });

  it('leaves what is not a formula', () => {
    expect(permuteFormula('120', swap, columns)).toBe('120');
  });

  it('leaves text it cannot parse', () => {
    expect(permuteFormula('=B2 +', swap, columns)).toBe('=B2 +');
  });

  it('cannot write a reference off the sheet', () => {
    // Every row it can name is a row the sort put somewhere real, so
    // there is no arithmetic here to push one negative.
    for (const input of ['=B2', '=$B$2', '=SUM(B2:B4)', '=B2+B4*B3']) {
      expect(permuteFormula(input, swap, columns)).not.toContain('#REF!');
    }
  });
});

describe('what a formula reads', () => {
  it('reports a single reference as one row', () => {
    expect(rowsRead('=B2', columns)).toEqual([{ first: 1, last: 1 }]);
  });

  it('reports a range as the run it covers', () => {
    expect(rowsRead('=SUM(B2:B4)', columns)).toEqual([{ first: 1, last: 3 }]);
  });

  it('ignores what lies outside the columns', () => {
    expect(rowsRead('=E2', columns)).toEqual([]);
  });

  it('finds references inside a call', () => {
    expect(rowsRead('=ROUND(B2/B4, 1)', columns)).toEqual([
      { first: 1, last: 1 },
      { first: 3, last: 3 }
    ]);
  });

  it('reads nothing out of a literal', () => {
    expect(rowsRead('120', columns)).toEqual([]);
  });
});
