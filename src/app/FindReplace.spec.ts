import { describe, expect, it } from 'vitest';

import type { SheetAutofit, SheetFindView, SheetSelection } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';
import type { SheetStats } from './Statistics';
import { GENERAL, PLAIN } from '../sheet/Format';

/**
 * Fill, find and replace, driven as commands through the service.
 *
 * The units underneath have their own specs; what this file asserts
 * is the behaviour those units are wired into — which cell the
 * selection lands on, what the count says, and how many presses of
 * ctrl-Z it takes to put four hundred cells back.
 */

function harness(fill?: (document: SheetDocument) => void) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { schedule, rowCount: 200, columnCount: 20 });
  const drain = () => {
    let guard = 0;
    while (queue.length > 0) {
      queue.shift()!();
      if (guard++ > 10_000) {
        throw new Error('the pump never finished');
      }
    }
  };
  return { document, service, drain };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

const where = (service: SheetService) => {
  const at = latest<SheetSelection>(service.selection);
  return `${at.row},${at.column}`;
};

const finding = (service: SheetService) => latest<SheetFindView>(service.findView);

describe('filling down and across', () => {
  it('repeats the top row of the selection down it, moving the formulas', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '2');
      d.setCell(1, 0, '3');
      d.setCell(2, 0, '4');
      d.setCell(0, 1, '=A1*10');
    });
    service.setSelection(0, 1, 2, 1);
    service.fillDown();
    drain();

    // Fully parenthesised, which is `Rewrite.ts`'s documented choice
    // from Phase 5: a writer that dropped brackets would have to know
    // the precedence table as exactly as the parser, and the one
    // place they could disagree is where a fill silently changes what
    // a formula means.
    expect(document.sheet.input(1, 1)).toBe('=(A2*10)');
    expect(document.sheet.value(2, 1)).toBe(40);
  });

  /**
   * One cell selected takes from the cell above, which is what every
   * spreadsheet does. Without it Ctrl+D silently does nothing nine
   * times out of ten, because people press it on the cell they want
   * filled rather than on a range.
   */
  it('takes from the cell above when only one cell is selected', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '5');
      d.setCell(0, 1, '=A1+1');
    });
    service.setSelection(1, 1, 1, 1);
    service.fillDown();
    drain();

    expect(document.sheet.input(1, 1)).toBe('=(A2+1)');
  });

  it('fills right the same way', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '=1+1');
    });
    service.setSelection(0, 0, 0, 2);
    service.fillRight();
    drain();

    expect(document.sheet.value(0, 2)).toBe(2);
  });

  it('does nothing at the top edge rather than reading off the sheet', () => {
    const { service, document } = harness(d => d.setCell(0, 0, 'x'));
    service.setSelection(0, 0, 0, 0);
    service.fillDown();
    expect(document.sheet.input(0, 0)).toBe('x');
  });

  /** One action, one press of ctrl-Z. */
  it('is one step on the undo stack however many cells it wrote', () => {
    const { service, document, drain } = harness(d => d.setCell(0, 0, '7'));
    service.setSelection(0, 0, 40, 0);
    service.fillDown();
    drain();
    expect(document.sheet.value(40, 0)).toBe(7);

    service.undo();
    drain();
    expect(document.sheet.input(40, 0)).toBe('');
    expect(document.sheet.value(0, 0)).toBe(7);
  });
});

describe('finding', () => {
  const book = (d: SheetDocument) => {
    d.setCell(0, 0, 'apple');
    d.setCell(2, 3, 'Pineapple');
    d.setCell(5, 0, 'apple tart');
    d.setCell(7, 1, 'pear');
  };

  it('counts the matches and lands on the first one at or after the selection', () => {
    const { service } = harness(book);
    service.setSelection(1, 0, 1, 0);
    service.find('apple', false, false, true);

    expect(finding(service).matches).toBe(3);
    expect(where(service)).toBe('2,3');
    expect(finding(service).active).toBe(2);
  });

  /** Somebody who selected a cell and searched for what is in it stays put. */
  it('offers the cell the selection is already on', () => {
    const { service } = harness(book);
    service.setSelection(0, 0, 0, 0);
    service.find('apple', false, false, true);
    expect(where(service)).toBe('0,0');
    expect(finding(service).active).toBe(1);
  });

  it('steps forward and wraps', () => {
    const { service } = harness(book);
    service.find('apple', false, false, true);
    expect(where(service)).toBe('0,0');
    service.findStep(true);
    expect(where(service)).toBe('2,3');
    service.findStep(true);
    expect(where(service)).toBe('5,0');
    service.findStep(true);
    expect(where(service)).toBe('0,0');
  });

  it('steps backward and wraps the other way', () => {
    const { service } = harness(book);
    service.find('apple', false, false, true);
    service.findStep(false);
    expect(where(service)).toBe('5,0');
  });

  it('stays where it is when nothing matches', () => {
    const { service } = harness(book);
    service.setSelection(3, 3, 3, 3);
    service.find('quince', false, false, true);
    expect(finding(service).matches).toBe(0);
    expect(where(service)).toBe('3,3');
  });

  /**
   * Moving off a match has to stop saying "1 of 3". The count is
   * unchanged and only the position moved, which is why the two are
   * published by different paths.
   */
  it('forgets which match it is on when the selection moves away', () => {
    const { service } = harness(book);
    service.find('apple', false, false, true);
    expect(finding(service).active).toBe(1);
    service.setSelection(9, 9, 9, 9);
    expect(finding(service).active).toBe(0);
    expect(finding(service).matches).toBe(3);
  });

  it('empties the search when it is closed', () => {
    const { service } = harness(book);
    service.find('apple', false, false, true);
    service.clearFind();
    expect(finding(service).query).toBe('');
    expect(finding(service).matches).toBe(0);
  });
});

describe('replacing', () => {
  const book = (d: SheetDocument) => {
    d.setCell(0, 0, 'cat');
    d.setCell(1, 0, 'cattle');
    d.setCell(2, 0, 'dog');
  };

  it('replaces the match the selection is on and moves to the next', () => {
    const { service, document, drain } = harness(book);
    service.find('cat', false, false, true);
    expect(where(service)).toBe('0,0');

    service.replaceOne('dog');
    drain();

    expect(document.sheet.input(0, 0)).toBe('dog');
    expect(document.sheet.input(1, 0)).toBe('cattle');
    expect(where(service)).toBe('1,0');
  });

  it('replaces everything in one step of undo', () => {
    const { service, document, drain } = harness(book);
    service.find('cat', false, false, true);
    service.replaceAll('ox');
    drain();

    expect(document.sheet.input(0, 0)).toBe('ox');
    expect(document.sheet.input(1, 0)).toBe('oxtle');
    expect(document.sheet.input(2, 0)).toBe('dog');

    service.undo();
    drain();
    expect(document.sheet.input(0, 0)).toBe('cat');
    expect(document.sheet.input(1, 0)).toBe('cattle');
  });

  /**
   * An edit can destroy a match or create one, and a stale list steps
   * somebody to a cell that no longer says what they searched for.
   */
  it('re-counts after replacing', () => {
    const { service, drain } = harness(book);
    service.find('cat', false, false, true);
    expect(finding(service).matches).toBe(2);
    service.replaceAll('ox');
    drain();
    expect(finding(service).matches).toBe(0);
  });

  it('rewrites a formula as the text it is', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '1');
      d.setCell(1, 0, '=SUM(A1:A1)');
    });
    service.find('SUM(', false, false, true);
    service.replaceAll('AVERAGE(');
    drain();

    expect(document.sheet.input(1, 0)).toBe('=AVERAGE(A1:A1)');
    expect(document.sheet.value(1, 0)).toBe(1);
  });

  it('does nothing when there is nothing to replace', () => {
    const { service, document } = harness(book);
    service.replaceAll('anything');
    expect(document.sheet.input(0, 0)).toBe('cat');
  });
});

describe('the status bar totals', () => {
  it('follows the selection', () => {
    const { service } = harness(d => {
      d.setCell(0, 0, '1');
      d.setCell(1, 0, '2');
      d.setCell(2, 0, '9');
    });
    service.setSelection(0, 0, 1, 0);
    expect(latest<SheetStats>(service.selectionStats).sum).toBe(3);

    service.setSelection(0, 0, 2, 0);
    expect(latest<SheetStats>(service.selectionStats).sum).toBe(12);
  });

  /** A recalculated cell inside the selection changes the total. */
  it('follows a recalculation', () => {
    const { service, drain } = harness(d => {
      d.setCell(0, 0, '1');
      d.setCell(1, 0, '=A1*10');
    });
    service.setSelection(0, 0, 1, 0);
    expect(latest<SheetStats>(service.selectionStats).sum).toBe(11);

    service.setCell(0, 0, '5');
    drain();
    expect(latest<SheetStats>(service.selectionStats).sum).toBe(55);
  });
});

/**
 * Inserting and deleting rows and columns, through the service.
 *
 * `Structure.budget.spec.ts` counts what a shift rewrites;
 * `Shift.spec.ts` says what each reference does. This is the layer
 * those two are wired into: what the selection covers, what the undo
 * stack holds, and what the geometry does with the column widths.
 */
describe('inserting and deleting rows', () => {
  it('pushes the rows below down, and takes their formulas with them', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '10');
      d.setCell(1, 0, '20');
      d.setCell(2, 0, '=SUM(A1:A2)');
    });
    service.insertRows(1, 1);
    drain();

    expect(document.sheet.input(0, 0)).toBe('10');
    expect(document.sheet.input(1, 0)).toBe('');
    expect(document.sheet.input(2, 0)).toBe('20');
    expect(document.sheet.input(3, 0)).toBe('=SUM(A1:A3)');
    expect(document.sheet.value(3, 0)).toBe(30);
  });

  it('inserts as many rows as the selection covers', () => {
    const { service, document, drain } = harness(d => d.setCell(0, 0, 'a'));
    service.insertRows(0, 3);
    drain();
    expect(document.sheet.input(3, 0)).toBe('a');
  });

  it('deletes rows, breaking only what pointed into them', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '10');
      d.setCell(1, 0, '20');
      d.setCell(2, 0, '=A1');
      d.setCell(3, 0, '=A2');
    });
    service.deleteRows(1, 1);
    drain();

    expect(document.sheet.input(1, 0)).toBe('=A1');
    expect(document.sheet.value(1, 0)).toBe(10);
    expect(document.sheet.input(2, 0)).toBe('=#REF!');
  });

  /** One action, one press of ctrl-Z, however far it reached. */
  it('is one step on the undo stack', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '10');
      d.setCell(1, 0, '=A1*2');
    });
    service.insertRows(0, 1);
    drain();
    expect(document.sheet.input(1, 0)).toBe('10');

    service.undo();
    drain();
    expect(document.sheet.input(0, 0)).toBe('10');
    // Exactly as it was typed: undo puts back the text it recorded
    // rather than shifting the shifted version back, so a formula
    // that went out and came home is not quietly reprinted.
    expect(document.sheet.input(1, 0)).toBe('=A1*2');
    expect(document.sheet.value(1, 0)).toBe(20);
  });

  /**
   * The hard direction. A delete destroys two things an opposite
   * shift cannot bring back: the cells that were in the row, and the
   * formulas it turned into `#REF!`.
   */
  it('puts back what a delete destroyed', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '10');
      d.setCell(1, 0, 'gone');
      d.setCell(2, 0, '=A2&"!"');
    });
    service.deleteRows(1, 1);
    drain();
    expect(document.sheet.input(1, 0)).toBe('=(#REF!&"!")');

    service.undo();
    drain();

    expect(document.sheet.input(1, 0)).toBe('gone');
    expect(document.sheet.input(2, 0)).toBe('=A2&"!"');
    expect(document.sheet.value(2, 0)).toBe('gone!');
  });

  it('moves the formats with the rows', () => {
    const { service, document, drain } = harness(d => d.setCell(0, 0, '1'));
    document.setFormat(0, 0, { number: { kind: 'percent', places: 0 }, paint: PLAIN });
    service.insertRows(0, 1);
    drain();

    expect(document.formatAt(1, 0).number.kind).toBe('percent');
    expect(document.formatAt(0, 0).number.kind).toBe('general');
  });
});

describe('inserting and deleting columns', () => {
  it('moves the cells and their references sideways', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 1, '5');
      d.setCell(0, 2, '=B1*2');
    });
    service.insertColumns(0, 1);
    drain();

    expect(document.sheet.input(0, 2)).toBe('5');
    expect(document.sheet.input(0, 3)).toBe('=(C1*2)');
    expect(document.sheet.value(0, 3)).toBe(10);
  });

  /**
   * Widths belong to the columns they describe. Inserting in front of
   * a wide column and leaving the widths alone makes the wrong column
   * wide.
   */
  it('moves the column widths too', () => {
    const { service, drain } = harness();
    service.setColumnWidth(2, 240);
    service.insertColumns(0, 1);
    drain();

    const geometry = latest<{ columnWidths: readonly number[] }>(service.geometry);
    expect(geometry.columnWidths[3]).toBe(240);
    expect(geometry.columnWidths[2]).not.toBe(240);
  });

  it('brings the widths back when the insert is undone', () => {
    const { service, drain } = harness();
    service.setColumnWidth(2, 240);
    service.insertColumns(0, 1);
    drain();
    service.undo();
    drain();

    const geometry = latest<{ columnWidths: readonly number[] }>(service.geometry);
    expect(geometry.columnWidths[2]).toBe(240);
  });
});

/**
 * Autofit, which is the one thing neither thread can do alone.
 *
 * This side knows every string in a column and nothing about fonts;
 * the render worker knows the font and holds thirty rows. So this
 * side narrows a million cells to a shortlist and the other measures
 * it. These assert the shortlist.
 */
describe('the candidates a column sends to be measured', () => {
  const latestFit = (service: SheetService) => latest<SheetAutofit>(service.autofit);

  it('sends the longest strings in the column, longest first', () => {
    const { service } = harness(d => {
      d.setCell(0, 0, 'short');
      d.setCell(1, 0, 'a much longer string');
      d.setCell(2, 0, 'medium one');
    });
    service.measureColumns(0, 0);

    const column = latestFit(service).columns[0];
    expect(column.column).toBe(0);
    expect(column.samples[0]).toBe('a much longer string');
    expect(column.samples).toContain('short');
  });

  /** What the screen shows, not what was typed. */
  it('sends the displayed string and not the formula', () => {
    const { service, document, drain } = harness(d => {
      d.setCell(0, 0, '1');
      d.setCell(1, 0, '=A1+1');
    });
    drain();
    document.setFormat(1, 0, { number: { kind: 'currency', places: 2, symbol: '$' }, paint: PLAIN });
    service.measureColumns(0, 0);

    expect(latestFit(service).columns[0].samples).toContain('$2.00');
  });

  /** A bold heading is wider than the same string plain. */
  it('says which of them are bold', () => {
    const { service, document } = harness(d => d.setCell(0, 0, 'Heading'));
    document.setFormat(0, 0, { number: GENERAL, paint: { ...PLAIN, bold: true } });
    service.measureColumns(0, 0);

    expect(latestFit(service).columns[0].bold[0]).toBe(true);
  });

  it('sends nothing for a column with nothing in it', () => {
    const { service } = harness();
    service.measureColumns(3, 3);
    expect(latestFit(service).columns[0].samples).toEqual([]);
  });

  it('sends a column at a time for a range', () => {
    const { service } = harness(d => {
      d.setCell(0, 0, 'a');
      d.setCell(0, 1, 'b');
    });
    service.measureColumns(0, 1);
    expect(latestFit(service).columns.map(entry => entry.column)).toEqual([0, 1]);
  });

  /** Asking twice is a change, or the second autofit does nothing. */
  it('moves the serial every time it is asked', () => {
    const { service } = harness(d => d.setCell(0, 0, 'a'));
    service.measureColumns(0, 0);
    const first = latestFit(service).serial;
    service.measureColumns(0, 0);
    expect(latestFit(service).serial).toBeGreaterThan(first);
  });

  /** A shortlist, not the whole column. */
  it('does not send a thousand strings', () => {
    const { service } = harness(d => {
      for (let row = 0; row < 1_000; row++) {
        d.setCell(row, 0, `row ${row}`);
      }
    });
    service.measureColumns(0, 0);
    expect(latestFit(service).columns[0].samples.length).toBeLessThan(20);
  });
});

describe('filtering to what the cursor is on', () => {
  const book = (d: SheetDocument) => {
    d.setCell(0, 0, 'Region');
    d.setCell(0, 1, 'Units');
    d.setCell(1, 0, 'North');
    d.setCell(1, 1, '10');
    d.setCell(2, 0, 'South');
    d.setCell(2, 1, '20');
    d.setCell(3, 0, 'North');
    d.setCell(3, 1, '30');
  };

  const hidden = (service: SheetService) =>
    latest<{ hiddenRows: readonly number[] }>(service.geometry).hiddenRows;

  it('hides the rows that do not match, and keeps the heading', () => {
    const { service } = harness(book);
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();

    expect(hidden(service)).toEqual([2]);
  });

  it('shows every row again', () => {
    const { service } = harness(book);
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    service.clearFilter();

    expect(hidden(service)).toEqual([]);
  });

  /**
   * Two sets, because they are undone by different things: clearing a
   * filter must not reveal a row somebody hid on purpose.
   */
  it('leaves a row that was hidden on purpose hidden', () => {
    const { service } = harness(book);
    service.hideRows(3, 3);
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    expect(hidden(service)).toEqual([2, 3]);

    service.clearFilter();
    expect(hidden(service)).toEqual([3]);
  });

  it('stops at the blank row, like a sort does', () => {
    const { service } = harness(d => {
      book(d);
      // A second table, which has nothing to do with the first.
      d.setCell(5, 0, 'Largest');
      d.setCell(6, 0, 'Smallest');
    });
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();

    expect(hidden(service)).toEqual([2]);
  });

  /**
   * A snapshot and not a rule: an edit that changes a cell does not
   * make rows vanish under somebody's hands.
   */
  it('does not re-run itself when a cell changes', () => {
    const { service, drain } = harness(book);
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    expect(hidden(service)).toEqual([2]);

    service.setCell(3, 0, 'South');
    drain();
    expect(hidden(service)).toEqual([2]);
  });

  it('moves the filtered rows when a row is inserted above them', () => {
    const { service, drain } = harness(book);
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    service.insertRows(0, 1);
    drain();

    expect(hidden(service)).toEqual([3]);
  });
});
