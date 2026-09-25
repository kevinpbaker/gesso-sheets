import { describe, expect, it } from 'vitest';

import {
  cellKey,
  columnIndex,
  columnName,
  columnOf,
  formatRef,
  inBounds,
  MAX_COLUMNS,
  MAX_ROWS,
  parseAddress,
  parseRef,
  rangeKeys,
  rangeSize,
  relativeRef,
  rowOf
} from './A1';

describe('cell keys', () => {
  it('round-trips a row and column through one integer', () => {
    for (const [row, column] of [
      [0, 0],
      [1, 0],
      [0, 1],
      [999, 16_383],
      [MAX_ROWS - 1, MAX_COLUMNS - 1]
    ]) {
      const key = cellKey(row, column);
      expect(rowOf(key)).toBe(row);
      expect(columnOf(key)).toBe(column);
    }
  });

  it('stays inside the safe integer range at the far corner', () => {
    // The packing is only exact while this holds, and the bounds were
    // chosen so that it does.
    expect(cellKey(MAX_ROWS - 1, MAX_COLUMNS - 1)).toBeLessThan(Number.MAX_SAFE_INTEGER);
  });

  it('knows what is on the sheet', () => {
    expect(inBounds(0, 0)).toBe(true);
    expect(inBounds(MAX_ROWS - 1, MAX_COLUMNS - 1)).toBe(true);
    expect(inBounds(MAX_ROWS, 0)).toBe(false);
    expect(inBounds(0, MAX_COLUMNS)).toBe(false);
    expect(inBounds(-1, 0)).toBe(false);
  });
});

describe('column names', () => {
  it('counts in bijective base 26, where AA follows Z', () => {
    expect(columnName(0)).toBe('A');
    expect(columnName(25)).toBe('Z');
    expect(columnName(26)).toBe('AA');
    expect(columnName(27)).toBe('AB');
    expect(columnName(51)).toBe('AZ');
    expect(columnName(52)).toBe('BA');
    expect(columnName(701)).toBe('ZZ');
    expect(columnName(702)).toBe('AAA');
    expect(columnName(MAX_COLUMNS - 1)).toBe('XFD');
  });

  it('reads them back', () => {
    for (const index of [0, 25, 26, 27, 701, 702, MAX_COLUMNS - 1]) {
      expect(columnIndex(columnName(index))).toBe(index);
    }
  });

  it('is not fooled by something that is not a column', () => {
    expect(columnIndex('')).toBeNull();
    expect(columnIndex('A1')).toBeNull();
    expect(columnIndex('-')).toBeNull();
  });
});

describe('A1 references', () => {
  it('reads a plain reference', () => {
    expect(parseRef('A1')).toEqual(relativeRef(0, 0));
    expect(parseRef('B3')).toEqual(relativeRef(2, 1));
    expect(parseRef('XFD1048576')).toEqual(relativeRef(MAX_ROWS - 1, MAX_COLUMNS - 1));
  });

  it('is case-insensitive, as a sheet is', () => {
    expect(parseRef('a1')).toEqual(parseRef('A1'));
  });

  /**
   * The `$` flags are carried rather than resolved, because Phase 5's
   * fill handle is the whole reason they exist: extending `=A1+$B$1`
   * downwards adjusts the first and leaves the second, and it can only
   * do that if the parse remembered which was which.
   */
  it('remembers which halves were absolute', () => {
    expect(parseRef('$A$1')).toEqual({ row: 0, column: 0, rowAbsolute: true, columnAbsolute: true });
    expect(parseRef('$A1')).toEqual({ row: 0, column: 0, rowAbsolute: false, columnAbsolute: true });
    expect(parseRef('A$1')).toEqual({ row: 0, column: 0, rowAbsolute: true, columnAbsolute: false });
  });

  it('writes them back exactly as they were', () => {
    for (const text of ['A1', '$A$1', '$A1', 'A$1', 'XFD1048576', 'BC42']) {
      expect(formatRef(parseRef(text)!)).toBe(text);
    }
  });

  it('refuses what is not a reference', () => {
    for (const text of ['', 'A', '1', 'A0', 'SUM', 'A1:B2', '$', 'ABCD1', '1A']) {
      expect(parseRef(text)).toBeNull();
    }
  });

  /**
   * Off the sheet is a different answer from malformed, and is
   * deliberately not decided here: `ZZZ9999999` is a well-formed
   * reference to a cell that does not exist, and turning it into
   * `#REF!` is evaluation's job.
   */
  it('parses a well-formed reference that points off the sheet', () => {
    const ref = parseRef('XFE1');
    expect(ref).not.toBeNull();
    expect(inBounds(ref!.row, ref!.column)).toBe(false);
  });
});

describe('ranges', () => {
  it('walks its cells in row-major order', () => {
    const keys = [...rangeKeys({ start: relativeRef(0, 0), end: relativeRef(1, 1) })];
    expect(keys).toEqual([cellKey(0, 0), cellKey(0, 1), cellKey(1, 0), cellKey(1, 1)]);
  });

  it('normalises its corners, because a selection can be dragged upwards', () => {
    const downward = [...rangeKeys({ start: relativeRef(0, 0), end: relativeRef(1, 1) })];
    const upward = [...rangeKeys({ start: relativeRef(1, 1), end: relativeRef(0, 0) })];
    expect(upward).toEqual(downward);
  });

  it('counts without walking', () => {
    expect(rangeSize({ start: relativeRef(0, 0), end: relativeRef(99, 9) })).toBe(1000);
    expect(rangeSize({ start: relativeRef(5, 5), end: relativeRef(5, 5) })).toBe(1);
  });
});

describe('an address typed into the name box', () => {
  it('reads a single cell as the range it is', () => {
    expect(parseAddress('B7')).toEqual({
      start: { row: 6, column: 1, rowAbsolute: false, columnAbsolute: false },
      end: { row: 6, column: 1, rowAbsolute: false, columnAbsolute: false }
    });
  });

  it('reads it in either case', () => {
    expect(parseAddress('b7')).toEqual(parseAddress('B7'));
  });

  it('reads a range', () => {
    const range = parseAddress('A1:C9');
    expect(range?.start.column).toBe(0);
    expect(range?.end.row).toBe(8);
    expect(range?.end.column).toBe(2);
  });

  it('keeps the dollars, so `$A$1` is still absolute', () => {
    expect(parseAddress('$A$1')?.start).toEqual({
      row: 0,
      column: 0,
      rowAbsolute: true,
      columnAbsolute: true
    });
  });

  it('ignores space around it', () => {
    expect(parseAddress('  B7  ')).toEqual(parseAddress('B7'));
  });

  /**
   * The whole string has to be the address. `B7+1` is a formula
   * somebody typed in the wrong box, and jumping to B7 would be
   * reading half of what they wrote and acting on it.
   */
  it('refuses anything that is not only an address', () => {
    expect(parseAddress('B7+1')).toBeNull();
    expect(parseAddress('SUM')).toBeNull();
    expect(parseAddress('')).toBeNull();
    expect(parseAddress('A1:B2:C3')).toBeNull();
    expect(parseAddress('7B')).toBeNull();
    expect(parseAddress('A0')).toBeNull();
  });
});
