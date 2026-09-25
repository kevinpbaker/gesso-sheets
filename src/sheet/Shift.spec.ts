import { describe, expect, it } from 'vitest';

import { shiftFormula, shiftIndex, type Shift } from './Shift';

const insertRows = (at: number, count = 1): Shift => ({ axis: 'row', at, by: count });
const deleteRows = (at: number, count = 1): Shift => ({ axis: 'row', at, by: -count });
const insertColumns = (at: number, count = 1): Shift => ({ axis: 'column', at, by: count });
const deleteColumns = (at: number, count = 1): Shift => ({ axis: 'column', at, by: -count });

describe('where an index lands', () => {
  it('moves down what an insert pushed down', () => {
    expect(shiftIndex(9, insertRows(5))).toBe(10);
    expect(shiftIndex(5, insertRows(5))).toBe(6);
  });

  it('leaves alone what is above the line', () => {
    expect(shiftIndex(4, insertRows(5))).toBe(4);
    expect(shiftIndex(4, deleteRows(5))).toBe(4);
  });

  it('moves up what a delete pulled up', () => {
    expect(shiftIndex(9, deleteRows(5))).toBe(8);
    expect(shiftIndex(9, deleteRows(5, 3))).toBe(6);
  });

  /** There is no index to point at, and saying so is the only honest answer. */
  it('has nowhere for a deleted index to go', () => {
    expect(shiftIndex(5, deleteRows(5))).toBe(-1);
    expect(shiftIndex(7, deleteRows(5, 3))).toBe(-1);
    expect(shiftIndex(8, deleteRows(5, 3))).toBe(5);
  });
});

describe('inserting rows', () => {
  it('moves a reference below the line', () => {
    expect(shiftFormula('=A10', insertRows(4))).toBe('=A11');
  });

  it('leaves a reference above the line alone', () => {
    expect(shiftFormula('=A2', insertRows(4))).toBe('=A2');
  });

  /**
   * The distinction every implementation gets wrong exactly once. A
   * fill pins `$A$1`; an insert does not. The dollar sign says "do
   * not move when I am copied", not "do not notice the sheet" — and a
   * reference that ignored an insert would quietly point at somebody
   * else's data.
   */
  it('moves an absolute reference too', () => {
    expect(shiftFormula('=$A$10', insertRows(4))).toBe('=$A$11');
    expect(shiftFormula('=A$10', insertRows(4))).toBe('=A$11');
  });

  it('moves every reference in a formula', () => {
    // Fully parenthesised, which is `Print.ts`'s documented choice.
    expect(shiftFormula('=ROUND(D10/$D$20*100,1)', insertRows(4))).toBe('=ROUND(((D11/$D$21)*100),1)');
  });

  /**
   * A formula that did not move still has to be rewritten. Shifting
   * is positional: a formula in row 1 reading `=A900` names the cell
   * at A900, and after an insert that cell is somewhere else.
   */
  it('rewrites a formula that did not itself move', () => {
    expect(shiftFormula('=A900', insertRows(500))).toBe('=A901');
  });

  it('says so when nothing in it referenced the line', () => {
    const formula = '=A1+B2';
    expect(shiftFormula(formula, insertRows(50))).toBe(formula);
  });
});

describe('deleting rows', () => {
  it('pulls a reference below the deletion up', () => {
    expect(shiftFormula('=A10', deleteRows(4))).toBe('=A9');
    expect(shiftFormula('=A10', deleteRows(4, 3))).toBe('=A7');
  });

  /** The cell it named is gone, and there is nothing else to name. */
  it('breaks a reference to a deleted cell', () => {
    expect(shiftFormula('=A5', deleteRows(4))).toBe('=#REF!');
    expect(shiftFormula('=A5+1', deleteRows(4))).toBe('=(#REF!+1)');
  });

  it('breaks an absolute reference to a deleted cell too', () => {
    expect(shiftFormula('=$A$5', deleteRows(4))).toBe('=#REF!');
  });
});

describe('ranges, whose corners do not move together', () => {
  /** The rows it was adding are all still in it, and there is one more. */
  it('grows when a row is inserted inside it', () => {
    expect(shiftFormula('=SUM(A1:A10)', insertRows(4))).toBe('=SUM(A1:A11)');
  });

  it('moves whole when a row is inserted above it', () => {
    expect(shiftFormula('=SUM(A5:A10)', insertRows(0))).toBe('=SUM(A6:A11)');
  });

  it('is untouched by a row inserted below it', () => {
    expect(shiftFormula('=SUM(A1:A10)', insertRows(50))).toBe('=SUM(A1:A10)');
  });

  /** Still a run of cells; there are fewer of them. */
  it('shrinks when a row inside it is deleted', () => {
    expect(shiftFormula('=SUM(A1:A10)', deleteRows(4))).toBe('=SUM(A1:A9)');
    expect(shiftFormula('=SUM(A1:A10)', deleteRows(4, 3))).toBe('=SUM(A1:A7)');
  });

  /**
   * A range whose first rows went away keeps the ones that survived.
   *
   * `A5:A10` is rows 5 to 10; deleting three rows from row 5 takes 5,
   * 6 and 7, leaving 8, 9 and 10 — which are now rows 5, 6 and 7. So
   * the same three cells, under their new names.
   */
  it('clamps rather than breaking when a corner is deleted', () => {
    expect(shiftFormula('=SUM(A5:A10)', deleteRows(4, 3))).toBe('=SUM(A5:A7)');
    // And from the other end: `A1:A6` loses rows 5 and 6, keeping 1–4.
    expect(shiftFormula('=SUM(A1:A6)', deleteRows(4, 3))).toBe('=SUM(A1:A4)');
  });

  /** Down to one cell rather than to nothing, while one survives. */
  it('shrinks to what is left of it', () => {
    // `A5:A7` with rows 6, 7 and 8 deleted keeps only row 5.
    expect(shiftFormula('=SUM(A5:A7)', deleteRows(5, 3))).toBe('=SUM(A5:A5)');
  });

  /** Only when every cell it covered is gone is there nothing to name. */
  it('breaks when the whole range is deleted', () => {
    expect(shiftFormula('=SUM(A5:A7)', deleteRows(4, 3))).toBe('=SUM(#REF!)');
    expect(shiftFormula('=SUM(A5:A7)', deleteRows(4, 9))).toBe('=SUM(#REF!)');
  });

  it('keeps a range written backwards written backwards', () => {
    expect(shiftFormula('=SUM(A10:A1)', insertRows(4))).toBe('=SUM(A11:A1)');
  });
});

describe('columns, which are the same rules on the other axis', () => {
  it('moves a reference to the right of the line', () => {
    expect(shiftFormula('=C1', insertColumns(1))).toBe('=D1');
    expect(shiftFormula('=$C$1', insertColumns(1))).toBe('=$D$1');
  });

  it('pulls a reference left when a column is deleted', () => {
    expect(shiftFormula('=C1', deleteColumns(0))).toBe('=B1');
  });

  it('breaks a reference to a deleted column', () => {
    expect(shiftFormula('=C1', deleteColumns(2))).toBe('=#REF!');
  });

  it('grows a range a column was inserted into', () => {
    expect(shiftFormula('=SUM(A1:D1)', insertColumns(2))).toBe('=SUM(A1:E1)');
  });

  /** A row insert does not move a column reference, and the reverse. */
  it('leaves the other axis alone', () => {
    expect(shiftFormula('=C5', insertColumns(0))).toBe('=D5');
    expect(shiftFormula('=C5', insertRows(0))).toBe('=C6');
  });
});

describe('what shifting does not touch', () => {
  it('leaves a literal alone', () => {
    expect(shiftFormula('42', insertRows(0))).toBe('42');
    expect(shiftFormula('hello', deleteRows(0))).toBe('hello');
  });

  it('leaves text that is not a formula alone even if it looks like one', () => {
    expect(shiftFormula('A1', insertRows(0))).toBe('A1');
  });

  /** It is already showing an error; mangling it would lose what was typed. */
  it('leaves an unparseable formula as it was typed', () => {
    expect(shiftFormula('=SUM(', insertRows(0))).toBe('=SUM(');
  });

  it('does nothing for a shift of nothing', () => {
    expect(shiftFormula('=A10', { axis: 'row', at: 0, by: 0 })).toBe('=A10');
  });
});

/**
 * A shift is positional, and a position on Sheet 2 is not a position
 * on Sheet 1.
 *
 * Two facts decide whether a reference moves, and the reference only
 * holds one of them: an unqualified `A5` means A5 on the formula's
 * own sheet, so which sheet the formula sits on is part of the
 * question. A workbook of one sheet leaves the shift unqualified and
 * everything moves, which is every caller before Phase 13.
 */
describe('an insert on one sheet of several', () => {
  const rows = (at: number, by: number, sheet?: string) => ({ axis: 'row' as const, at, by, sheet });

  it('moves a reference to the sheet that changed, from anywhere', () => {
    expect(shiftFormula('=Sheet2!A5', rows(0, 1, 'Sheet2'), 'Sheet1')).toBe('=Sheet2!A6');
    expect(shiftFormula('=Sheet2!A5', rows(0, 1, 'Sheet2'), 'Sheet2')).toBe('=Sheet2!A6');
  });

  it('leaves a reference to any other sheet alone', () => {
    expect(shiftFormula('=Sheet3!A5', rows(0, 1, 'Sheet2'), 'Sheet1')).toBe('=Sheet3!A5');
    expect(shiftFormula('=Sheet3!A5', rows(0, 1, 'Sheet2'), 'Sheet2')).toBe('=Sheet3!A5');
  });

  /** Unqualified means "my own sheet", so where the formula lives decides. */
  it('reads a bare reference as one to the formula\u2019s own sheet', () => {
    expect(shiftFormula('=A5', rows(0, 1, 'Sheet2'), 'Sheet2')).toBe('=A6');
    expect(shiftFormula('=A5', rows(0, 1, 'Sheet2'), 'Sheet1')).toBe('=A5');
  });

  it('does not care what case the name was written in', () => {
    expect(shiftFormula('=sheet2!A5', rows(0, 1, 'SHEET2'), 'Sheet1')).toBe('=sheet2!A6');
  });

  it('moves the ones that point at it and no others, in one formula', () => {
    expect(shiftFormula('=A5+Sheet2!A5+Sheet3!A5', rows(0, 1, 'Sheet2'), 'Sheet1')).toBe(
      '=((A5+Sheet2!A6)+Sheet3!A5)'
    );
  });

  it('deletes across sheets the same way', () => {
    expect(shiftFormula('=Sheet2!A5', rows(4, -1, 'Sheet2'), 'Sheet1')).toBe('=#REF!');
    expect(shiftFormula('=Sheet3!A5', rows(4, -1, 'Sheet2'), 'Sheet1')).toBe('=Sheet3!A5');
  });

  it('moves everything when the workbook has one sheet', () => {
    expect(shiftFormula('=A5+Sheet2!A5', rows(0, 1))).toBe('=(A6+Sheet2!A6)');
  });
});
