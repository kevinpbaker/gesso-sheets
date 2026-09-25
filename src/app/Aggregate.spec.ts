import { describe, expect, it } from 'vitest';

import { Sheet } from '../sheet/Sheet';
import { aggregateOf } from './Aggregate';
import { describeStats, NO_STATS } from './Statistics';

function sheetOf(values: readonly (string | null)[][]): Sheet {
  const sheet = new Sheet();
  values.forEach((line, row) =>
    line.forEach((input, column) => {
      if (input !== null) {
        sheet.setCell(row, column, input);
      }
    })
  );
  sheet.recalculate();
  return sheet;
}

const rect = (firstRow: number, lastRow: number, firstColumn: number, lastColumn: number) => ({
  firstRow,
  lastRow,
  firstColumn,
  lastColumn
});

describe('what the status bar says about a selection', () => {
  it('counts, sums and averages the numbers', () => {
    const sheet = sheetOf([['1'], ['2'], ['6']]);
    expect(aggregateOf(sheet, rect(0, 2, 0, 0), 100)).toEqual({
      count: 3,
      numeric: 3,
      sum: 9,
      average: 3,
      min: 1,
      max: 6
    });
  });

  it('counts text but does not add it', () => {
    const sheet = sheetOf([['1'], ['apples'], ['3']]);
    const stats = aggregateOf(sheet, rect(0, 2, 0, 0), 100);
    expect(stats.count).toBe(3);
    expect(stats.numeric).toBe(2);
    expect(stats.sum).toBe(4);
    expect(stats.average).toBe(2);
  });

  it('takes the value of a formula, not its text', () => {
    const sheet = sheetOf([['4'], ['=A1*2']]);
    expect(aggregateOf(sheet, rect(0, 1, 0, 0), 100).sum).toBe(12);
  });

  /**
   * An error is a value a cell holds and not a number it holds.
   * Counting `#DIV/0!` as a zero would drag an average down quietly,
   * which is the one thing a spreadsheet must never do.
   */
  it('does not treat an error as a zero', () => {
    const sheet = sheetOf([['10'], ['=1/0'], ['20']]);
    const stats = aggregateOf(sheet, rect(0, 2, 0, 0), 100);
    expect(stats.count).toBe(3);
    expect(stats.numeric).toBe(2);
    expect(stats.average).toBe(15);
  });

  it('says nothing about an empty selection', () => {
    expect(aggregateOf(new Sheet(), rect(0, 5, 0, 5), 100)).toEqual(NO_STATS);
  });

  it('ignores cells outside the rectangle', () => {
    const sheet = sheetOf([
      ['1', '100'],
      ['2', '200']
    ]);
    expect(aggregateOf(sheet, rect(0, 1, 0, 0), 100).sum).toBe(3);
  });

  /**
   * The chain the proof surface builds lives past the end of the
   * sheet. Summing a measuring instrument into somebody's status bar
   * would report it as though it were their data.
   */
  it('stops at the end of the sheet', () => {
    const sheet = sheetOf([['5']]);
    sheet.setCell(100, 0, '1000');
    sheet.recalculate();
    expect(aggregateOf(sheet, rect(0, 10_000, 0, 0), 100).sum).toBe(5);
  });

  /**
   * The performance contract, asserted as a count rather than a
   * timing. Ctrl+A over a sheet holding three numbers must cost three
   * lookups and not a million, because it happens on the thread that
   * is supposed to be answering viewport commands.
   */
  it('costs the store rather than the selection when the selection is bigger', () => {
    const sheet = sheetOf([['1', '2'], ['3']]);
    let lookups = 0;
    const counted = new Proxy(sheet, {
      get(target, property, receiver) {
        if (property === 'value') {
          return (row: number, column: number) => {
            lookups++;
            return target.value(row, column);
          };
        }
        return Reflect.get(target, property, receiver) as unknown;
      }
    }) as Sheet;

    const stats = aggregateOf(counted, rect(0, 9_999, 0, 99), 10_000);
    expect(stats.sum).toBe(6);
    expect(lookups).toBe(3);
  });

  /** And the other way round: a small selection walks itself. */
  it('costs the selection when the selection is smaller', () => {
    const sheet = new Sheet();
    for (let row = 0; row < 500; row++) {
      sheet.setCell(row, 0, String(row));
    }
    sheet.recalculate();
    let lookups = 0;
    const counted = new Proxy(sheet, {
      get(target, property, receiver) {
        if (property === 'value') {
          return (row: number, column: number) => {
            lookups++;
            return target.value(row, column);
          };
        }
        return Reflect.get(target, property, receiver) as unknown;
      }
    }) as Sheet;

    expect(aggregateOf(counted, rect(0, 3, 0, 0), 10_000).sum).toBe(0 + 1 + 2 + 3);
    expect(lookups).toBe(4);
  });
});

describe('printing the statistics', () => {
  it('shows sum, average and count when there are numbers', () => {
    expect(describeStats({ count: 3, numeric: 3, sum: 9, average: 3, min: 1, max: 6 })).toBe(
      'Sum 9  ·  Average 3  ·  Count 3'
    );
  });

  /**
   * `Sum: 0` over a column of names is a true statement about the
   * empty set that reads as a false one about the names.
   */
  it('shows only a count when nothing in the selection is a number', () => {
    expect(describeStats({ ...NO_STATS, count: 4 })).toBe('Count 4');
  });

  it('says nothing at all about an empty selection', () => {
    expect(describeStats(NO_STATS)).toBe('');
  });

  it('keeps a long average short enough for a status bar', () => {
    expect(describeStats({ count: 3, numeric: 3, sum: 10, average: 10 / 3, min: 1, max: 6 })).toContain(
      'Average 3.3333'
    );
  });
});
