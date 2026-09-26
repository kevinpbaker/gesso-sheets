import { describe, expect, it } from 'vitest';

import { Workbook } from './Workbook';

/**
 * AutoComplete on a long column — Phase 21.
 *
 * The question is asked on every keystroke, so its claim is that a
 * column of a hundred thousand entries answers it as a column of ten
 * does. Measured as a ratio against the short column, in the same run,
 * rather than in milliseconds that belong to this machine.
 */
describe('offering a word', () => {
  function column(length: number): Workbook {
    const book = new Workbook();
    const sheet = book.sheet(0);
    for (let row = 0; row < length; row++) {
      sheet.setCell(row, 0, `Customer ${String(row).padStart(6, '0')}`);
    }
    return book;
  }

  function perKeystroke(book: Workbook, prefixes: readonly string[]): number {
    // The first question builds the index; the keystrokes after it are the claim.
    book.completeIn(0, 0, 'C');
    const start = performance.now();
    for (let round = 0; round < 200; round++) {
      for (const prefix of prefixes) {
        book.completeIn(0, 0, prefix);
      }
    }
    return (performance.now() - start) / (200 * prefixes.length);
  }

  it('costs a long column what it costs a short one', () => {
    const prefixes = ['C', 'Cu', 'Customer 0', 'Customer 00000', 'Customer 000007'];
    const short = column(10);
    const long = column(100_000);

    // Ten customers start this way, so nothing is offered; one zebra does.
    expect(long.completeIn(0, 0, 'Customer 09999')).toBeNull();
    long.setCell(0, 100_000, 0, 'Zebra crossing');
    expect(long.completeIn(0, 0, 'Ze')).toBe('Zebra crossing');

    const small = perKeystroke(short, prefixes);
    const large = perKeystroke(long, prefixes);
    console.info(`[words budget] 10 entries ${(small * 1000).toFixed(2)} µs, 100,000 entries ${(large * 1000).toFixed(2)} µs a keystroke`);

    // A binary search over a hundred thousand is seventeen steps against
    // four; a walk would be ten thousand times the short column's.
    expect(large).toBeLessThan(small * 20 + 0.01);
  });

  it('keeps up with a column typed row by row', () => {
    const book = column(100_000);
    book.completeIn(0, 0, 'C');
    const start = performance.now();
    for (let row = 100_000; row < 100_200; row++) {
      book.setCell(0, row, 0, `Supplier ${row}`);
      book.completeIn(0, 0, 'Su');
    }
    const each = (performance.now() - start) / 200;
    console.info(`[words budget] a row typed and a word offered, 100,000 rows down: ${each.toFixed(3)} ms`);

    expect(book.completeIn(0, 0, 'Supplier 10019')).toBeNull();
    expect(book.completeIn(0, 0, 'Supplier 1001')).toBeNull();
    book.setCell(0, 100_200, 0, 'Walrus');
    expect(book.completeIn(0, 0, 'Wa')).toBe('Walrus');
    expect(each).toBeLessThan(2);
  });
});
