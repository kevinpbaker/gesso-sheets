import { describe, expect, it } from 'vitest';

import { GENERAL, PLAIN } from '../sheet/Format';
import { SheetDocument } from './SheetDocument';
import { currentRegion, looksLikeHeader } from './SheetRanges';
import { sortRect } from './SheetSort';

/**
 * Written through the model rather than the document, so the fixture
 * is not itself on the undo stack — the specs below ask what the
 * *sort* put there.
 */
function sheetOf(rows: readonly (readonly string[])[]): SheetDocument {
  const document = new SheetDocument();
  rows.forEach((line, row) => line.forEach((input, column) => document.sheet.setCell(row, column, input)));
  document.sheet.recalculate();
  return document;
}

const column = (document: SheetDocument, index: number, rows: number): string[] =>
  Array.from({ length: rows }, (_, row) => document.sheet.display(row, index));

const rect = (firstRow: number, lastRow: number, firstColumn: number, lastColumn: number) => ({
  firstRow,
  lastRow,
  firstColumn,
  lastColumn
});

describe('sorting a range', () => {
  it('puts the rows in order of the column it was given', () => {
    const document = sheetOf([['c'], ['a'], ['b']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(column(document, 0, 3)).toEqual(['a', 'b', 'c']);
  });

  it('sorts the other way when asked', () => {
    const document = sheetOf([['c'], ['a'], ['b']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: false, hasHeader: false });
    expect(column(document, 0, 3)).toEqual(['c', 'b', 'a']);
  });

  /**
   * The single most destructive thing a spreadsheet can do quietly is
   * sort one column and leave the row beside it behind.
   */
  it('takes the whole row with it', () => {
    const document = sheetOf([
      ['c', '3'],
      ['a', '1'],
      ['b', '2']
    ]);
    sortRect(document, rect(0, 2, 0, 1), { column: 0, ascending: true, hasHeader: false });

    expect(column(document, 0, 3)).toEqual(['a', 'b', 'c']);
    expect(column(document, 1, 3)).toEqual(['1', '2', '3']);
  });

  it('leaves a heading where it is', () => {
    const document = sheetOf([['Name'], ['c'], ['a']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: true });
    expect(column(document, 0, 3)).toEqual(['Name', 'a', 'c']);
  });

  /** The same order `<` uses, so a sort and a comparison never disagree. */
  it('orders numbers before text', () => {
    const document = sheetOf([['apple'], ['10'], ['2']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(column(document, 0, 3)).toEqual(['2', '10', 'apple']);
  });

  /** A blank is an absence, not a small value. */
  it('sorts blanks last, whichever way it runs', () => {
    const document = sheetOf([['b'], [''], ['a']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(column(document, 0, 3)).toEqual(['a', 'b', '']);

    const other = sheetOf([['b'], [''], ['a']]);
    sortRect(other, rect(0, 2, 0, 0), { column: 0, ascending: false, hasHeader: false });
    expect(column(other, 0, 3)).toEqual(['b', 'a', '']);
  });

  /**
   * Ties keep the order they had, which is how somebody sorts by two
   * columns without a dialog for it: sort by the second, then by the
   * first.
   */
  it('is stable', () => {
    const document = sheetOf([
      ['a', 'first'],
      ['a', 'second'],
      ['a', 'third']
    ]);
    sortRect(document, rect(0, 2, 0, 1), { column: 0, ascending: true, hasHeader: false });
    expect(column(document, 1, 3)).toEqual(['first', 'second', 'third']);
  });

  /**
   * A formula that said `=B2*C2` in row 2 has to say `=B7*C7` when it
   * lands in row 7, or it reads somebody else's numbers.
   */
  it('moves a formula so it still reads its own row', () => {
    const document = sheetOf([
      ['b', '2', '=B1*10'],
      ['a', '5', '=B2*10']
    ]);
    sortRect(document, rect(0, 1, 0, 2), { column: 0, ascending: true, hasHeader: false });

    expect(column(document, 0, 2)).toEqual(['a', 'b']);
    document.sheet.recalculate();
    expect(document.sheet.value(0, 2)).toBe(50);
    expect(document.sheet.value(1, 2)).toBe(20);
  });

  /** A `$` pins a reference out of the block, exactly as in a fill. */
  it('leaves an absolute reference alone', () => {
    const document = sheetOf([
      ['b', '=$E$1'],
      ['a', '=$E$1']
    ]);
    sortRect(document, rect(0, 1, 0, 1), { column: 0, ascending: true, hasHeader: false });
    expect(document.sheet.input(0, 1)).toBe('=$E$1');
  });

  it('is one step on the undo stack', () => {
    const document = sheetOf([['c'], ['a'], ['b']]);
    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(column(document, 0, 3)).toEqual(['a', 'b', 'c']);

    document.undo();
    expect(column(document, 0, 3)).toEqual(['c', 'a', 'b']);
  });

  /**
   * Left behind, a sorted table keeps its bold total row where the
   * total *was*, against somebody else's numbers. Found in a browser,
   * by sorting the demo sheet.
   */
  it('takes the formats with the rows', () => {
    const document = sheetOf([['c'], ['a'], ['b']]);
    document.setFormat(0, 0, { number: GENERAL, paint: { ...PLAIN, bold: true } });

    sortRect(document, rect(0, 2, 0, 0), { column: 0, ascending: true, hasHeader: false });

    // `c` went to the bottom and its bold went with it.
    expect(column(document, 0, 3)).toEqual(['a', 'b', 'c']);
    expect(document.formatAt(2, 0).paint.bold).toBe(true);
    expect(document.formatAt(0, 0).paint.bold).toBe(false);
  });

  it('does nothing to a range that is already in order', () => {
    const document = sheetOf([['a'], ['b']]);
    sortRect(document, rect(0, 1, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(document.canUndo).toBe(false);
  });

  it('does nothing to a range of one row', () => {
    const document = sheetOf([['a']]);
    sortRect(document, rect(0, 0, 0, 0), { column: 0, ascending: true, hasHeader: false });
    expect(document.canUndo).toBe(false);
  });
});

/**
 * The block a cell is standing in.
 *
 * This exists because of what happened without it: a single-cell sort
 * widened to the *whole sheet*, swept three unrelated tables into one
 * ordering, and dragged formulas across each other until some pointed
 * off the sheet and said `#REF!`. One press of ctrl-Z put it back,
 * and it should never have been offered.
 */
describe('the block a cell is standing in', () => {
  const sheet = () =>
    sheetOf([
      ['Region', 'Units'],
      ['North', '120'],
      ['South', '157'],
      [],
      ['Largest', '4958'],
      ['Smallest', '1140']
    ]);

  it('grows to the edges of the table and stops at the blank row', () => {
    expect(currentRegion(sheet(), 1, 0, 100, 20)).toEqual({
      firstRow: 0,
      lastRow: 2,
      firstColumn: 0,
      lastColumn: 1
    });
  });

  it('finds the second table from inside it', () => {
    expect(currentRegion(sheet(), 4, 0, 100, 20)).toEqual({
      firstRow: 4,
      lastRow: 5,
      firstColumn: 0,
      lastColumn: 1
    });
  });

  it('is one cell when the cell is on its own', () => {
    expect(currentRegion(sheetOf([['x']]), 5, 5, 100, 20)).toEqual({
      firstRow: 5,
      lastRow: 5,
      firstColumn: 5,
      lastColumn: 5
    });
  });

  it('does not run off the end of the sheet', () => {
    const document = sheetOf([['a'], ['b']]);
    expect(currentRegion(document, 0, 0, 2, 1).lastRow).toBe(1);
  });
});

describe('guessing at a heading', () => {
  it('sees one when text sits over numbers', () => {
    const document = sheetOf([['Units'], ['120'], ['157']]);
    expect(looksLikeHeader(document, { firstRow: 0, lastRow: 2, firstColumn: 0, lastColumn: 0 })).toBe(true);
  });

  /** A column of names under the word `Name` is not a heading over data. */
  it('does not see one when it is text all the way down', () => {
    const document = sheetOf([['Name'], ['Ada'], ['Grace']]);
    expect(looksLikeHeader(document, { firstRow: 0, lastRow: 2, firstColumn: 0, lastColumn: 0 })).toBe(true);
  });

  it('sees none when the first row is numbers', () => {
    const document = sheetOf([['1'], ['2']]);
    expect(looksLikeHeader(document, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 0 })).toBe(false);
  });

  it('sees none in a block of one row', () => {
    const document = sheetOf([['Units']]);
    expect(looksLikeHeader(document, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 })).toBe(false);
  });
});
