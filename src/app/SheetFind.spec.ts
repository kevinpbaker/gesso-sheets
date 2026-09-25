import { describe, expect, it } from 'vitest';

import { cellKey } from '../sheet/A1';
import { Sheet } from '../sheet/Sheet';
import { at, DEFAULT_FIND, findMatches, replaceIn, stepBack, stepTo, type FindOptions } from './SheetFind';

function sheetWith(cells: Record<string, string>): Sheet {
  const sheet = new Sheet();
  for (const [address, input] of Object.entries(cells)) {
    const column = address.charCodeAt(0) - 65;
    const row = Number(address.slice(1)) - 1;
    sheet.setCell(row, column, input);
  }
  sheet.recalculate();
  return sheet;
}

const options = (over: Partial<FindOptions> = {}): FindOptions => ({ ...DEFAULT_FIND, ...over });

describe('finding', () => {
  const sheet = sheetWith({
    A1: 'Apples',
    B1: 'apple pie',
    A2: 'Pears',
    C3: '=1+2',
    A4: 'APPLE'
  });

  /** Reading order, not the order the cells were typed. */
  it('returns matches in the order a person reads them', () => {
    const found = findMatches(sheet, 'apple', options(), 100, 26);
    expect(found.map(at)).toEqual([
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 3, column: 0 }
    ]);
  });

  it('respects case when asked to', () => {
    expect(findMatches(sheet, 'APPLE', options({ matchCase: true }), 100, 26).map(at)).toEqual([
      { row: 3, column: 0 }
    ]);
  });

  it('matches the whole cell when asked to', () => {
    expect(findMatches(sheet, 'apples', options({ wholeCell: true }), 100, 26).map(at)).toEqual([
      { row: 0, column: 0 }
    ]);
  });

  /**
   * The two searches find different cells, which is the whole reason
   * there are two: `=1+2` is `1+2` as a formula and `3` as a value.
   */
  it('searches formulas or values, and they are not the same search', () => {
    expect(findMatches(sheet, '1+2', options({ inFormulas: true }), 100, 26).map(at)).toEqual([
      { row: 2, column: 2 }
    ]);
    expect(findMatches(sheet, '1+2', options({ inFormulas: false }), 100, 26)).toEqual([]);
    expect(findMatches(sheet, '3', options({ inFormulas: false }), 100, 26).map(at)).toEqual([
      { row: 2, column: 2 }
    ]);
  });

  it('finds nothing for an empty query', () => {
    expect(findMatches(sheet, '', options(), 100, 26)).toEqual([]);
  });

  /**
   * The proof surface's chain lives one row past the end of the
   * sheet. Finding somebody's query in a measuring instrument and
   * scrolling them to a cell they cannot reach is worse than not
   * finding it.
   */
  it('does not search past the end of the sheet', () => {
    const withChain = sheetWith({ A1: 'x' });
    withChain.setCell(500, 0, 'x');
    expect(findMatches(withChain, 'x', options(), 100, 26).map(at)).toEqual([{ row: 0, column: 0 }]);
  });
});

describe('stepping through matches', () => {
  const found = [cellKey(0, 0), cellKey(0, 1), cellKey(3, 0)];

  it('offers the cell you are standing on, then moves off it', () => {
    expect(stepTo(found, cellKey(0, 0), false)).toBe(cellKey(0, 0));
    expect(stepTo(found, cellKey(0, 0), true)).toBe(cellKey(0, 1));
  });

  it('wraps to the top from past the last match', () => {
    expect(stepTo(found, cellKey(9, 9), true)).toBe(cellKey(0, 0));
  });

  it('walks backwards and wraps to the bottom', () => {
    expect(stepBack(found, cellKey(0, 1))).toBe(cellKey(0, 0));
    expect(stepBack(found, cellKey(0, 0))).toBe(cellKey(3, 0));
  });

  it('says so when there is nothing to step to', () => {
    expect(stepTo([], 0, true)).toBe(-1);
    expect(stepBack([], 0)).toBe(-1);
  });
});

describe('replacing', () => {
  it('replaces every occurrence in a cell, as a text editor does', () => {
    expect(replaceIn('banana', 'a', 'b', options({ matchCase: true }))).toBe('bbnbnb');
  });

  it('replaces regardless of case, keeping what was not matched', () => {
    expect(replaceIn('Apple apple APPLE', 'apple', 'pear', options())).toBe('pear pear pear');
  });

  it('replaces the whole cell when that is what matched', () => {
    expect(replaceIn('Apples', 'apples', 'Pears', options({ wholeCell: true }))).toBe('Pears');
    expect(replaceIn('Apples and pears', 'apples', 'Pears', options({ wholeCell: true }))).toBe('Apples and pears');
  });

  /**
   * The query is somebody's typing, not a pattern. A find for `SUM(`
   * has to find `SUM(` rather than throwing, and `.` has to mean a
   * full stop.
   */
  it('treats the query as text and not as a pattern', () => {
    expect(replaceIn('=SUM(A1:A9)', 'SUM(', 'AVERAGE(', options())).toBe('=AVERAGE(A1:A9)');
    expect(replaceIn('a.b.c', '.', '-', options())).toBe('a-b-c');
    expect(replaceIn('a*b', '*', '+', options())).toBe('a+b');
  });

  it('leaves a cell alone when the query is empty', () => {
    expect(replaceIn('anything', '', 'x', options())).toBe('anything');
  });
});
