import { describe, expect, it } from 'vitest';

import { Sheet } from './Sheet';
import { Workbook } from './Workbook';
import { NA, SPILL } from './Values';

/**
 * Arrays, and the cells they spill into.
 *
 * A formula whose answer is more than one value fills the cells beside
 * and below its own, as Excel 365 does; a cell in the way stops it with
 * `#SPILL!`, and clearing that cell lets it through. `@` asks for the
 * one value instead, which is how a formula from an older file keeps
 * meaning what it meant.
 */

function sheetOf(grid: Record<string, string>): Sheet {
  const sheet = new Workbook().sheet(0);
  for (const [address, input] of Object.entries(grid)) {
    const column = address.charCodeAt(0) - 65;
    const row = Number(address.slice(1)) - 1;
    sheet.setCell(row, column, input);
  }
  sheet.recalculate();
  return sheet;
}

const at = (sheet: Sheet, address: string) =>
  sheet.value(Number(address.slice(1)) - 1, address.charCodeAt(0) - 65);

const NUMBERS = { A1: '1', A2: '2', A3: '3', B1: '10', B2: '20', B3: '30' };

describe('an array formula', () => {
  it('spills a range across the cells below it', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3' });
    expect([at(sheet, 'D1'), at(sheet, 'D2'), at(sheet, 'D3')]).toEqual([1, 2, 3]);
    expect(sheet.spillOf(0, 3)).toEqual({ rows: 3, columns: 1 });
    expect(sheet.spilledFrom(2, 3)).toEqual({ row: 0, column: 3 });
  });

  it('does arithmetic cell by cell, two-dimensional included', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:B3*2' });
    expect([at(sheet, 'D1'), at(sheet, 'E1'), at(sheet, 'D3'), at(sheet, 'E3')]).toEqual([2, 20, 6, 60]);
  });

  it('pairs two ranges of the same shape', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3+B1:B3' });
    expect([at(sheet, 'D1'), at(sheet, 'D2'), at(sheet, 'D3')]).toEqual([11, 22, 33]);
  });

  it('stretches a column across a row, and says #N/A where shapes do not meet', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3*B1:B2' });
    expect([at(sheet, 'D1'), at(sheet, 'D2'), at(sheet, 'D3')]).toEqual([10, 40, NA]);
  });

  it('takes each branch of an IF cell by cell', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=IF(A1:A3>1, "big", "small")' });
    expect([at(sheet, 'D1'), at(sheet, 'D2'), at(sheet, 'D3')]).toEqual(['small', 'big', 'big']);
  });

  it('shows 0 for an empty cell it copied, as a formula does', () => {
    const sheet = sheetOf({ A1: '1', A3: '3', D1: '=A1:A3' });
    expect(at(sheet, 'D2')).toBe(0);
  });

  it('is one value to whatever reads the cell it is written in', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3*2', F1: '=D1+1', F2: '=SUM(D1:D3)' });
    expect(at(sheet, 'F1')).toBe(3);
    expect(at(sheet, 'F2')).toBe(12);
  });
});

describe('what reads a spilled cell', () => {
  it('follows the array when what it was worked out from changes', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3*2', F1: '=D3+1' });
    expect(at(sheet, 'F1')).toBe(7);
    sheet.setCell(2, 0, '100');
    sheet.recalculate();
    expect(at(sheet, 'D3')).toBe(200);
    expect(at(sheet, 'F1')).toBe(201);
  });

  it('reads nothing once the array shrinks away from it', () => {
    const sheet = sheetOf({ ...NUMBERS, C1: '3', D1: '=OFFSET(A1, 0, 0, C1, 1)', F1: '=D3' });
    expect(at(sheet, 'F1')).toBe(3);
    sheet.setCell(0, 2, '2');
    sheet.recalculate();
    expect(at(sheet, 'D3')).toBeNull();
    expect(at(sheet, 'F1')).toBe(0);
  });
});

describe('a spill with something in its way', () => {
  it('says #SPILL! and fills nothing', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3', D3: 'in the way' });
    expect(at(sheet, 'D1')).toEqual(SPILL);
    expect(at(sheet, 'D2')).toBeNull();
    expect(at(sheet, 'D3')).toBe('in the way');
  });

  it('spills again when what was in the way is cleared', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3', D3: 'in the way' });
    sheet.clearCell(2, 3);
    sheet.recalculate();
    expect([at(sheet, 'D1'), at(sheet, 'D2'), at(sheet, 'D3')]).toEqual([1, 2, 3]);
  });

  it('is blocked by typing into its area, and freed by clearing it', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3' });
    sheet.setCell(1, 3, 'mine');
    sheet.recalculate();
    expect(at(sheet, 'D1')).toEqual(SPILL);
    expect(at(sheet, 'D3')).toBeNull();
    sheet.clearCell(1, 3);
    sheet.recalculate();
    expect(at(sheet, 'D2')).toBe(2);
  });

  it('leaves a spilled cell alone when it is cleared, since it holds nothing', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3' });
    sheet.clearCell(1, 3);
    sheet.recalculate();
    expect(at(sheet, 'D2')).toBe(2);
  });

  it('is blocked by another array, and let through when that one goes', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3', C2: '=B1:C1' });
    // C2's array spills right into D2, which D1's array already filled.
    expect(at(sheet, 'C2')).toEqual(SPILL);
    sheet.clearCell(0, 3);
    sheet.recalculate();
    expect(at(sheet, 'C2')).toBe(10);
    expect(at(sheet, 'D2')).toBe(0);
  });

  it('gives its cells back when its formula is replaced', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3', F1: '=D2' });
    sheet.setCell(0, 3, '5');
    sheet.recalculate();
    expect(at(sheet, 'D2')).toBeNull();
    expect(at(sheet, 'F1')).toBe(0);
  });

  it('keeps nothing of its spill in what is saved', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3' });
    expect([...sheet.entries()].filter(cell => cell.column === 3)).toEqual([{ row: 0, column: 3, input: '=A1:A3' }]);
  });

  it('spills again from where an inserted row moved it', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=A1:A3' });
    sheet.shift({ axis: 'row', at: 0, by: 1 });
    sheet.recalculate();
    expect([at(sheet, 'D2'), at(sheet, 'D3'), at(sheet, 'D4')]).toEqual([1, 2, 3]);
    expect(at(sheet, 'D1')).toBeNull();
  });
});

describe('@, the one value', () => {
  it('takes the cell in its own row from a column', () => {
    const sheet = sheetOf({ ...NUMBERS, D2: '=@A1:A3*2' });
    expect(at(sheet, 'D2')).toBe(4);
    expect(sheet.spillOf(1, 3)).toBeNull();
  });

  it('keeps an argument whole where the function takes a range', () => {
    const sheet = sheetOf({ ...NUMBERS, D2: '=SUM(A1:A3)' });
    expect(at(sheet, 'D2')).toBe(6);
  });
});

describe('arrays handed to functions', () => {
  it('lets SUMPRODUCT and SUM take arithmetic over ranges', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=SUMPRODUCT(A1:A3*B1:B3)', D2: '=SUM(A1:A3*B1:B3)' });
    expect(at(sheet, 'D1')).toBe(140);
    expect(at(sheet, 'D2')).toBe(140);
  });

  /** POI's bug 60848: a double negation over a range of blanks and text. */
  it('turns a range into numbers with --', () => {
    const sheet = sheetOf({ A1: 'x', D1: '=SUMPRODUCT(--(A2:A5))' });
    expect(at(sheet, 'D1')).toBe(0);
  });

  it('counts with a condition written as arithmetic', () => {
    const sheet = sheetOf({ ...NUMBERS, D1: '=SUMPRODUCT((A1:A3>1)*B1:B3)' });
    expect(at(sheet, 'D1')).toBe(50);
  });
});
