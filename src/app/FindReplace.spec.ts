import { describe, expect, it } from 'vitest';

import type { SheetFindView, SheetSelection } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';
import type { SheetStats } from './Statistics';

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
