import { isError, type CellValue } from '../sheet/Values';
import type { Sheet } from '../sheet/Sheet';
import type { Rect } from './SheetRanges';
import { NO_STATS, type SheetStats } from './Statistics';

/**
 * The selection's statistics, at the cost of whichever is smaller —
 * the rectangle, or the cells the sheet actually holds.
 *
 * This matters more than it looks. Ctrl+A selects ten thousand rows
 * by a hundred columns, and a status bar that walked its selection
 * would do a million lookups on every press of an arrow key, on the
 * thread that is supposed to be answering viewport commands. So a
 * selection bigger than the store is served by walking the store and
 * asking which cells fall inside it: select-all over a sheet holding
 * fifty numbers costs fifty.
 *
 * `rowCount` bounds it the way `snapshotOf` is bounded, and for the
 * same reason: the store will hold a cell anywhere in a million rows,
 * but the *sheet* is `rowCount` tall, and what lives past the end is
 * the proof surface's chain of two hundred thousand formulas. Summing
 * those into somebody's status bar would be reporting an instrument
 * as though it were their data.
 */
export function aggregateOf(sheet: Sheet, rect: Rect, rowCount: number): SheetStats {
  const lastRow = Math.min(rect.lastRow, rowCount - 1);
  if (lastRow < rect.firstRow || rect.lastColumn < rect.firstColumn) {
    return NO_STATS;
  }
  const area = (lastRow - rect.firstRow + 1) * (rect.lastColumn - rect.firstColumn + 1);

  const tally = new Tally();
  if (area <= sheet.size) {
    for (let row = rect.firstRow; row <= lastRow; row++) {
      for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
        tally.add(sheet.value(row, column));
      }
    }
  } else {
    for (const cell of sheet.entries()) {
      if (
        cell.row >= rect.firstRow &&
        cell.row <= lastRow &&
        cell.column >= rect.firstColumn &&
        cell.column <= rect.lastColumn
      ) {
        tally.add(sheet.value(cell.row, cell.column));
      }
    }
  }
  return tally.result();
}

/**
 * The running totals.
 *
 * A class rather than four `let`s because the two walks above have to
 * produce identical answers, and the surest way to make two loops
 * agree is to give them one body.
 */
class Tally {
  private count = 0;
  private numeric = 0;
  private sum = 0;
  private min = Number.POSITIVE_INFINITY;
  private max = Number.NEGATIVE_INFINITY;

  add(value: CellValue): void {
    if (value === null) {
      return;
    }
    this.count++;
    // An error is a value a cell holds and not a number it holds.
    // Counting `#DIV/0!` as a zero would quietly drag an average
    // down, which is the sort of wrong answer a spreadsheet must
    // never give quietly.
    if (typeof value !== 'number' || isError(value)) {
      return;
    }
    this.numeric++;
    this.sum += value;
    this.min = Math.min(this.min, value);
    this.max = Math.max(this.max, value);
  }

  result(): SheetStats {
    if (this.numeric === 0) {
      return { ...NO_STATS, count: this.count };
    }
    return {
      count: this.count,
      numeric: this.numeric,
      sum: this.sum,
      average: this.sum / this.numeric,
      min: this.min,
      max: this.max
    };
  }
}
