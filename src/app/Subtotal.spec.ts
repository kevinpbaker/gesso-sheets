import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';

/**
 * `SUBTOTAL` under a filter and over hidden rows.
 *
 * The rows are the document's and the formula is the workbook's, so
 * this is the path between them: a filter or a hide changes no cell,
 * and the total has to move anyway.
 */

function harness() {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, rowCount: 100, columnCount: 10 });
  const drain = () => {
    while (queue.length > 0) {
      queue.shift()!();
    }
  };
  // Region, amount — a list somebody would filter — and three totals.
  const rows = [
    ['Region', 'Amount'],
    ['North', '10'],
    ['South', '20'],
    ['North', '30'],
    ['East', '40']
  ];
  rows.forEach(([region, amount], row) => {
    service.setCell(row, 0, region);
    service.setCell(row, 1, amount);
  });
  service.setCell(6, 1, '=SUBTOTAL(9, B2:B5)');
  service.setCell(7, 1, '=SUBTOTAL(109, B2:B5)');
  service.setCell(8, 1, '=SUM(B2:B5)');
  drain();
  const value = (row: number) => document.sheet.value(row, 1);
  return { service, document, drain, value };
}

describe('SUBTOTAL, and the rows that show', () => {
  it('totals what a filter shows', () => {
    const { service, drain, value } = harness();
    expect(value(6)).toBe(100);
    // Filter to the rows whose region is the one the cursor is on.
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    drain();
    expect(value(6)).toBe(40);
    expect(value(7)).toBe(40);
    // SUM has no such rule, which is the difference.
    expect(value(8)).toBe(100);
  });

  it('comes back when the filter is cleared', () => {
    const { service, drain, value } = harness();
    service.setSelection(1, 0, 1, 0);
    service.filterToSelection();
    drain();
    service.clearFilter();
    drain();
    expect(value(6)).toBe(100);
  });

  /** 1–11 keep a row hidden by hand; 101–111 leave it out. That is the only difference. */
  it('keeps a hidden row in 9 and leaves it out of 109', () => {
    const { service, drain, value } = harness();
    service.hideRows(2, 2);
    drain();
    expect(value(6)).toBe(100);
    expect(value(7)).toBe(80);
    service.showRows(2, 2);
    drain();
    expect(value(7)).toBe(100);
  });
});
