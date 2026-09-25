import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import { Workbook } from './Workbook';

/**
 * Many sheets, one graph.
 *
 * The claims worth writing down are the ones a second graph would
 * get wrong: that a formula reading across sheets is recalculated in
 * an order that accounts for both, that a row inserted on one sheet
 * moves the references to *that* sheet and no others, and that a
 * sheet arriving or leaving reaches the formulas that named it.
 */

function book(names: string[] = ['Sheet1', 'Sheet2']): Workbook {
  return new Workbook(names);
}

describe('a formula that reads another sheet', () => {
  it('reads it', () => {
    const w = book();
    w.setCell(1, 0, 0, '10');
    w.setCell(0, 0, 0, '=Sheet2!A1');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(10);
  });

  it('sums a range on it', () => {
    const w = book();
    for (let row = 0; row < 4; row++) {
      w.setCell(1, row, 0, String(row + 1));
    }
    w.setCell(0, 0, 0, '=SUM(Sheet2!A1:A4)');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(10);
  });

  it('notices when the cell it reads changes', () => {
    const w = book();
    w.setCell(1, 0, 0, '10');
    w.setCell(0, 0, 0, '=Sheet2!A1*2');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(20);

    w.setCell(1, 0, 0, '11');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(22);
  });

  /**
   * The claim a second graph could not make. The chain runs
   * Sheet2!A1 → Sheet1!A1 → Sheet2!B1, so there is no order over
   * either sheet alone that evaluates all three correctly.
   */
  it('is recalculated in an order that crosses sheets', () => {
    const w = book();
    w.setCell(1, 0, 0, '2');
    w.setCell(0, 0, 0, '=Sheet2!A1*3');
    w.setCell(1, 0, 1, '=Sheet1!A1+1');
    w.recalculate();
    expect(w.value(1, 0, 1)).toBe(7);

    w.setCell(1, 0, 0, '5');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(15);
    expect(w.value(1, 0, 1)).toBe(16);
  });

  it('reads a whole column of another sheet', () => {
    const w = book();
    w.setCell(1, 0, 0, '1');
    w.setCell(1, 7, 0, '2');
    w.setCell(0, 0, 0, '=SUM(Sheet2!A:A)');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(3);
  });

  /**
   * Column A of Sheet 2 is not column A of Sheet 1.
   *
   * Counted rather than read, because the *value* is right either
   * way: a formula woken by the wrong sheet re-adds its own column
   * and gets the same answer. What a graph keying both columns as
   * `0` would cost is an evaluation per write anywhere in any A, and
   * that is the thing to assert.
   */
  it('does not wake a whole-column formula from the wrong sheet', () => {
    const w = book();
    w.setCell(0, 0, 1, '=SUM(A:A)');
    w.recalculate();

    const evaluated = w.stats.evaluated;
    w.setCell(1, 4, 0, '99');
    w.recalculate();
    expect(w.stats.evaluated).toBe(evaluated);

    w.setCell(0, 4, 0, '1');
    w.recalculate();
    expect(w.stats.evaluated).toBe(evaluated + 1);
    expect(w.value(0, 0, 1)).toBe(1);
  });

  it('says #REF! for a sheet the workbook does not have', () => {
    const w = book(['Sheet1']);
    w.setCell(0, 0, 0, '=Sheet9!A1');
    w.recalculate();
    expect(w.display(0, 0, 0)).toBe('#REF!');
  });

  it('does not care what case the sheet was written in', () => {
    const w = book();
    w.setCell(1, 0, 0, '7');
    w.setCell(0, 0, 0, '=sHeEt2!A1');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(7);
  });

  it('reads a sheet whose name needs quoting', () => {
    const w = book(['Sheet1', 'Q3 Budget']);
    w.setCell(1, 0, 0, '42');
    w.setCell(0, 0, 0, "='Q3 Budget'!A1");
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(42);
  });
});

describe('adding a sheet', () => {
  it('gives it a name nothing else has', () => {
    const w = book(['Sheet1', 'Sheet1', 'Sheet1']);
    expect(w.sheetNames()).toEqual(['Sheet1', 'Sheet1 2', 'Sheet1 3']);
  });

  /**
   * The formula was `#REF!` because there was nowhere for its edge to
   * go — so nothing in the graph could have woken it. The wiring is
   * redone on every structural change for exactly this.
   */
  it('wakes the formulas that were waiting for it', () => {
    const w = book(['Sheet1']);
    w.setCell(0, 0, 0, '=Later!A1');
    w.recalculate();
    expect(w.display(0, 0, 0)).toBe('#REF!');

    const later = w.addSheet('Later');
    w.setCell(later, 0, 0, '8');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(8);
  });
});

describe('renaming a sheet', () => {
  it('rewrites the formulas that named it', () => {
    const w = book();
    w.setCell(1, 0, 0, '3');
    w.setCell(0, 0, 0, '=Sheet2!A1+1');
    w.recalculate();

    expect(w.renameSheet(1, 'Data')).toBe(1);
    expect(w.input(0, 0, 0)).toBe('=Data!A1+1');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(4);
  });

  /**
   * Rewritten on the text and not through the printer, so a formula
   * that mentions the sheet keeps the shape somebody typed. Printed
   * back, `=Sheet2!A1+1` would return as `=(Data!A1+1)`.
   */
  it('leaves the rest of the formula exactly as it was typed', () => {
    const w = book();
    w.setCell(0, 0, 0, '=Sheet2!A1+1*2-3');
    w.renameSheet(1, 'Data');
    expect(w.input(0, 0, 0)).toBe('=Data!A1+1*2-3');
  });

  it('quotes the new name when it needs quoting', () => {
    const w = book();
    w.setCell(0, 0, 0, '=Sheet2!A1');
    w.renameSheet(1, 'Q3 Budget');
    expect(w.input(0, 0, 0)).toBe("='Q3 Budget'!A1");
  });

  /** A rename is not a search and replace over the text. */
  it('does not touch a string that happens to say the name', () => {
    const w = book();
    w.setCell(0, 0, 0, '="Sheet2 is where it is"');
    w.renameSheet(1, 'Data');
    expect(w.input(0, 0, 0)).toBe('="Sheet2 is where it is"');
  });

  it('leaves formulas on other sheets alone', () => {
    const w = book(['Sheet1', 'Sheet2', 'Sheet3']);
    w.setCell(0, 0, 0, '=Sheet3!A1');
    w.renameSheet(1, 'Data');
    expect(w.input(0, 0, 0)).toBe('=Sheet3!A1');
  });

  it('refuses to collide with a name already taken', () => {
    const w = book();
    w.renameSheet(1, 'Sheet1');
    expect(w.sheetNames()).toEqual(['Sheet1', 'Sheet1 2']);
  });
});

describe('removing a sheet', () => {
  it('leaves #REF! wherever it was read', () => {
    const w = book();
    w.setCell(1, 0, 0, '5');
    w.setCell(0, 0, 0, '=Sheet2!A1');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(5);

    expect(w.removeSheet(1)).toBe(true);
    w.recalculate();
    expect(w.display(0, 0, 0)).toBe('#REF!');
  });

  it('takes its cells with it', () => {
    const w = book();
    w.setCell(1, 0, 0, '5');
    w.removeSheet(1);
    expect(w.size).toBe(0);
  });

  /** Nothing in the application is written to survive a workbook of none. */
  it('will not remove the last one', () => {
    const w = book(['Only']);
    expect(w.removeSheet(0)).toBe(false);
    expect(w.sheetNames()).toEqual(['Only']);
  });

  it('carries the sheets after it down, cells and all', () => {
    const w = book(['Sheet1', 'Sheet2', 'Sheet3']);
    w.setCell(2, 3, 3, 'third');
    w.removeSheet(1);
    expect(w.sheetNames()).toEqual(['Sheet1', 'Sheet3']);
    expect(w.input(1, 3, 3)).toBe('third');
  });
});

describe('moving a sheet', () => {
  it('carries its cells to the new position', () => {
    const w = book(['A', 'B', 'C']);
    w.setCell(0, 0, 0, 'from A');
    w.setCell(2, 0, 0, 'from C');
    expect(w.moveSheet(0, 2)).toBe(true);

    expect(w.sheetNames()).toEqual(['B', 'C', 'A']);
    expect(w.input(2, 0, 0)).toBe('from A');
    expect(w.input(1, 0, 0)).toBe('from C');
  });

  it('keeps a cross-sheet formula pointing at the same sheet', () => {
    const w = book(['A', 'B', 'C']);
    w.setCell(2, 0, 0, '9');
    w.setCell(0, 0, 0, '=C!A1');
    w.recalculate();

    w.moveSheet(2, 0);
    w.recalculate();
    expect(w.value(1, 0, 0)).toBe(9);
  });
});

describe('duplicating a sheet', () => {
  it('copies the cells under a free name', () => {
    const w = book(['Data']);
    w.setCell(0, 0, 0, '2');
    w.setCell(0, 1, 0, '=A1*3');
    w.recalculate();

    const copy = w.duplicateSheet(0);
    w.recalculate();
    expect(w.nameOf(copy)).toBe('Data copy');
    expect(w.value(copy, 1, 0)).toBe(6);
  });

  /** A bare reference means "my own sheet", so it follows the copy. */
  it('leaves an unqualified reference pointing at the copy', () => {
    const w = book(['Data']);
    w.setCell(0, 0, 0, '2');
    w.setCell(0, 1, 0, '=A1*3');
    const copy = w.duplicateSheet(0);
    w.setCell(copy, 0, 0, '5');
    w.recalculate();

    expect(w.value(0, 1, 0)).toBe(6);
    expect(w.value(copy, 1, 0)).toBe(15);
  });
});

describe('a row inserted on one sheet', () => {
  it('moves the references to that sheet, from anywhere', () => {
    const w = book();
    w.setCell(0, 0, 0, '=Sheet2!A5');
    w.shift(1, { axis: 'row', at: 0, by: 1 });
    expect(w.input(0, 0, 0)).toBe('=Sheet2!A6');
  });

  it('leaves the other sheets where they were', () => {
    const w = book(['Sheet1', 'Sheet2', 'Sheet3']);
    w.setCell(0, 0, 0, '=A5+Sheet3!A5');
    w.shift(1, { axis: 'row', at: 0, by: 1 });
    expect(w.input(0, 0, 0)).toBe('=A5+Sheet3!A5');
  });

  it('moves its own sheet’s cells and nobody else’s', () => {
    const w = book();
    w.setCell(0, 0, 0, 'stays');
    w.setCell(1, 0, 0, 'moves');
    w.shift(1, { axis: 'row', at: 0, by: 1 });

    expect(w.input(0, 0, 0)).toBe('stays');
    expect(w.input(1, 0, 0)).toBe('');
    expect(w.input(1, 1, 0)).toBe('moves');
  });

  it('says #REF! across sheets when the row is deleted', () => {
    const w = book();
    w.setCell(0, 0, 0, '=Sheet2!A5');
    w.shift(1, { axis: 'row', at: 4, by: -1 });
    expect(w.input(0, 0, 0)).toBe('=#REF!');
  });
});

describe('a named range in a workbook', () => {
  it('names cells on the sheet it was taken from', () => {
    const w = book();
    w.setCell(1, 0, 0, '4');
    w.setCell(1, 1, 0, '6');
    w.names.define('Sales', {
      start: { ...relativeRef(0, 0), sheet: 'Sheet2' },
      end: { ...relativeRef(1, 0), sheet: 'Sheet2' }
    });
    w.namesChanged();
    w.setCell(0, 0, 0, '=SUM(Sales)');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(10);
  });

  it('wakes the formula that reads it when one of its cells changes', () => {
    const w = book();
    w.setCell(1, 0, 0, '4');
    w.names.define('Sales', {
      start: { ...relativeRef(0, 0), sheet: 'Sheet2' },
      end: { ...relativeRef(1, 0), sheet: 'Sheet2' }
    });
    w.namesChanged();
    w.setCell(0, 0, 0, '=SUM(Sales)');
    w.recalculate();

    w.setCell(1, 1, 0, '6');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(10);
  });
});
