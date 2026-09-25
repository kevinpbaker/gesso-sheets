import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import { Names, nameProblem, nameProblemText } from './Names';
import { Sheet } from './Sheet';

const area = (row: number, column: number, lastRow: number, lastColumn: number) => ({
  start: relativeRef(row, column),
  end: relativeRef(lastRow, lastColumn)
});

/**
 * What a name may be.
 *
 * Every rule exists to stop a name being mistaken for something else,
 * and the one that matters most is that a name may not look like a
 * reference: `A1` names a cell and always will.
 */
describe('whether a name may be used', () => {
  it.each(['Sales', '_private', 'Q3.total', 'Tax_2026', 'Costs'])('accepts %o', name => {
    expect(nameProblem(name)).toBeNull();
  });

  /** A bare column letter too: `A:A` is a reference as much as `A1` is. */
  it.each(['A1', '$A$1', 'B12', 'A', 'a', 'ZZ', 'XFD', 'XFD1'])('refuses %o, which is a reference', name => {
    expect(nameProblem(name)).toBe('reference');
  });

  it.each(['2sales', 'my name', 'a-b', 'total!', ''])('refuses %o on its shape', name => {
    expect(nameProblem(name)).not.toBeNull();
  });

  it('refuses one nobody could have meant', () => {
    expect(nameProblem('a'.repeat(300))).toBe('long');
  });

  it('says why, in words somebody can act on', () => {
    expect(nameProblemText('reference')).toContain('already means something else');
    expect(nameProblemText('shape')).toContain('starts with a letter');
  });
});

describe('the table of names', () => {
  it('finds a name however it was capitalised', () => {
    const names = new Names();
    names.define('Sales', area(1, 1, 96, 1));
    expect(names.rangeOf('SALES')).not.toBeNull();
    expect(names.rangeOf('sales')).not.toBeNull();
  });

  it('keeps the case it was given, for showing back', () => {
    const names = new Names();
    names.define('Sales', area(0, 0, 0, 0));
    expect(names.all()[0].name).toBe('Sales');
  });

  it('refuses a bad one and changes nothing', () => {
    const names = new Names();
    expect(names.define('A1', area(0, 0, 0, 0))).toBe('reference');
    expect(names.size).toBe(0);
  });

  it('redefines rather than duplicating', () => {
    const names = new Names();
    names.define('Sales', area(0, 0, 0, 0));
    names.define('SALES', area(5, 5, 9, 9));
    expect(names.size).toBe(1);
    expect(names.rangeOf('Sales')?.start.row).toBe(5);
  });

  it('lists them in the order somebody reads', () => {
    const names = new Names();
    names.define('Zulu', area(0, 0, 0, 0));
    names.define('alpha', area(0, 0, 0, 0));
    expect(names.all().map(entry => entry.name)).toEqual(['alpha', 'Zulu']);
  });
});

/**
 * A name is a reference by another spelling, so a sheet has to treat
 * it like one: evaluate it, depend on it, and move it.
 */
describe('a named range in a formula', () => {
  function sheetWith(): Sheet {
    const sheet = new Sheet();
    sheet.setCell(1, 1, '10');
    sheet.setCell(2, 1, '20');
    sheet.setCell(3, 1, '30');
    sheet.names.define('Sales', area(1, 1, 3, 1));
    sheet.namesChanged();
    return sheet;
  }

  it('sums the range it names', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=SUM(Sales)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(60);
  });

  /** The whole point: it is a range, not its first cell. */
  it('is a range argument rather than one value', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=COUNT(Sales)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(3);
  });

  it('is its first cell where a single value was wanted', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=Sales+1');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(11);
  });

  /** Without an edge, the total is correct once and stale after. */
  it('is woken by an edit inside the range it names', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=SUM(Sales)');
    sheet.recalculate();

    sheet.setCell(2, 1, '200');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(240);
  });

  it('is #NAME? when the sheet has no such name', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=SUM(Costs)');
    sheet.recalculate();
    expect(sheet.display(0, 5)).toBe('#NAME?');
  });

  /**
   * A function's name is taken, and saying so beats accepting a name
   * that could never be read.
   */
  it('refuses to name a range after a function', () => {
    const sheet = sheetWith();
    expect(sheet.names.define('TODAY', area(0, 0, 0, 0))).toBe('function');
    expect(sheet.names.define('MEDIAN', area(0, 0, 0, 0))).toBe('function');
    // And the library still works, which is what the rule protects.
    sheet.setCell(0, 5, '=SUM(1,2)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(3);
  });

  it('answers the new range when the name is redefined', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=SUM(Sales)');
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(60);

    sheet.names.define('Sales', area(1, 1, 2, 1));
    sheet.namesChanged();
    sheet.recalculate();
    expect(sheet.value(0, 5)).toBe(30);
  });

  it('moves with an insert, as a written reference would', () => {
    const sheet = sheetWith();
    sheet.setCell(0, 5, '=SUM(Sales)');
    sheet.recalculate();

    sheet.shift({ axis: 'row', at: 0, by: 1 });
    sheet.recalculate();
    expect(sheet.names.rangeOf('Sales')?.start.row).toBe(2);
    expect(sheet.value(1, 5)).toBe(60);
  });

  /**
   * A name whose every cell was deleted goes, rather than staying and
   * answering `#REF!` from something that appears to exist.
   */
  it('is removed when the rows it named are deleted', () => {
    const sheet = sheetWith();
    sheet.shift({ axis: 'row', at: 1, by: -3 });
    expect(sheet.names.rangeOf('Sales')).toBeNull();
  });
});
