import { describe, expect, it } from 'vitest';

import { fromTsv, looksTabular, toTsv } from './Tsv';

describe('writing TSV', () => {
  it('joins with tabs and newlines', () => {
    expect(toTsv([['a', 'b'], ['c', 'd']])).toBe('a\tb\nc\td');
  });

  it('writes an empty cell as nothing', () => {
    expect(toTsv([['a', '', 'c']])).toBe('a\t\tc');
  });

  /**
   * A cell holding a tab written raw would arrive back as two cells.
   * The quoting is CSV's, which is what every spreadsheet reads.
   */
  it('quotes a cell that would otherwise break the shape', () => {
    expect(toTsv([['a\tb']])).toBe('"a\tb"');
    expect(toTsv([['two\nlines']])).toBe('"two\nlines"');
    expect(toTsv([['say "no"']])).toBe('"say ""no"""');
  });
});

describe('reading TSV', () => {
  it('splits on tabs and newlines', () => {
    expect(fromTsv('a\tb\nc\td')).toEqual([['a', 'b'], ['c', 'd']]);
  });

  it('reads the quoting back', () => {
    expect(fromTsv('"a\tb"')).toEqual([['a\tb']]);
    expect(fromTsv('"two\nlines"')).toEqual([['two\nlines']]);
    expect(fromTsv('"say ""no"""')).toEqual([['say "no"']]);
  });

  it('round-trips everything that needs quoting', () => {
    const block = [
      ['plain', 'with\ttab'],
      ['with "quotes"', 'two\nlines'],
      ['', 'trailing ']
    ];
    expect(fromTsv(toTsv(block))).toEqual(block);
  });

  it('reads the line endings Windows and Excel produce', () => {
    expect(fromTsv('a\tb\r\nc\td')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(fromTsv('a\rb')).toEqual([['a'], ['b']]);
  });

  /**
   * Excel puts a newline on the end of what it copies. Read as a row
   * it would be a row of blanks, and pasting would wipe a row of the
   * sheet nobody meant to touch.
   */
  it('drops the trailing newline Excel adds', () => {
    expect(fromTsv('a\tb\nc\td\n')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(fromTsv('a\tb\r\n')).toEqual([['a', 'b']]);
  });

  /** A block has to be a rectangle for anything downstream to place it. */
  it('pads a ragged block to the widest row', () => {
    expect(fromTsv('a\tb\tc\nd')).toEqual([['a', 'b', 'c'], ['d', '', '']]);
  });

  it('reads one cell as a block of one', () => {
    expect(fromTsv('hello')).toEqual([['hello']]);
  });

  it('reads nothing as nothing', () => {
    expect(fromTsv('')).toEqual([]);
  });
});

describe('looksTabular', () => {
  it('is what decides between a block and one cell', () => {
    expect(looksTabular('a\tb')).toBe(true);
    expect(looksTabular('a\nb')).toBe(true);
    expect(looksTabular('hello')).toBe(false);
    expect(looksTabular('')).toBe(false);
  });
});
