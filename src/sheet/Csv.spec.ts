import { describe, expect, it } from 'vitest';

import { looksLikeFormula, parseCsv, toCsv } from './Csv';

describe('reading a CSV', () => {
  it('reads plain rows', () => {
    expect(parseCsv('a,b,c\n1,2,3\n').rows).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3']
    ]);
  });

  it('reads quoted cells holding the separator, quotes and line endings', () => {
    const { rows } = parseCsv('name,note\r\n"Smith, J","said ""hi""\r\nthen left"\r\n');
    expect(rows).toEqual([
      ['name', 'note'],
      ['Smith, J', 'said "hi"\nthen left']
    ]);
  });

  it('drops a byte-order mark', () => {
    expect(parseCsv('﻿Region,Units\nNorth,5').rows[0]).toEqual(['Region', 'Units']);
  });

  it('takes the separator a sep= line names, and drops the line', () => {
    const parsed = parseCsv('sep=;\na;b\n1,5;2,5\n');
    expect(parsed.delimiter).toBe(';');
    expect(parsed.rows).toEqual([
      ['a', 'b'],
      ['1,5', '2,5']
    ]);
  });

  /** A locale that writes 1,5 for one and a half has to separate with something else. */
  it('sniffs a semicolon file from its first record', () => {
    const parsed = parseCsv('Region;Price\nNorth;1,50\n');
    expect(parsed.delimiter).toBe(';');
    expect(parsed.rows[1]).toEqual(['North', '1,50']);
  });

  it('counts separators outside quotes only', () => {
    expect(parseCsv('"Smith, J";"Jones, K"\n').delimiter).toBe(';');
  });

  it('sniffs a tab', () => {
    expect(parseCsv('a\tb\n').delimiter).toBe('\t');
  });

  it('prefers the comma on a tie, and on a single column', () => {
    expect(parseCsv('a,b;c\n').delimiter).toBe(',');
    expect(parseCsv('just one\n').delimiter).toBe(',');
  });

  it('reads old Mac line endings', () => {
    expect(parseCsv('a,b\r1,2\r').rows).toEqual([
      ['a', 'b'],
      ['1', '2']
    ]);
  });

  it('keeps a blank line in the middle as a blank row', () => {
    expect(parseCsv('a\n\nb\n').rows).toEqual([['a'], [''], ['b']]);
  });

  it('pads a ragged file to a rectangle', () => {
    expect(parseCsv('a,b,c\n1\n').rows).toEqual([
      ['a', 'b', 'c'],
      ['1', '', '']
    ]);
  });

  it('keeps an empty last cell', () => {
    expect(parseCsv('a,\n').rows).toEqual([['a', '']]);
  });

  it('reads a formula as the characters it is', () => {
    expect(parseCsv('=1+2,=HYPERLINK("x")\n').rows).toEqual([['=1+2', '=HYPERLINK("x")']]);
  });

  it('reads nothing as nothing', () => {
    expect(parseCsv('').rows).toEqual([]);
  });
});

describe('writing a CSV', () => {
  it('writes records with the RFC line ending, after the last as well', () => {
    expect(
      toCsv([
        ['a', 'b'],
        ['1', '2']
      ])
    ).toBe('a,b\r\n1,2\r\n');
  });

  it('quotes what would otherwise be misread', () => {
    expect(toCsv([['Smith, J', 'said "hi"', 'two\nlines', ' padded', 'plain']])).toBe(
      '"Smith, J","said ""hi""","two\nlines"," padded",plain\r\n'
    );
  });

  it('quotes for the separator it is writing with', () => {
    expect(toCsv([['1,5', '2;5']], ';')).toBe('1,5;"2;5"\r\n');
  });

  it('writes nothing for no rows', () => {
    expect(toCsv([])).toBe('');
  });

  it('reads back what it wrote', () => {
    const rows = [
      ['Region', 'Note'],
      ['North', 'a "quoted" word, and\na second line'],
      ['', ' spaced ']
    ];
    expect(parseCsv(toCsv(rows)).rows).toEqual(rows);
  });
});

describe('what a spreadsheet would read as a formula', () => {
  it('is text starting with any character that begins one somewhere', () => {
    for (const text of ['=1+2', '+SUM(A1)', '-2+3', '@A1', '\tx', '\rx']) {
      expect(looksLikeFormula(text)).toBe(true);
    }
  });

  it('is not text that merely contains one', () => {
    expect(looksLikeFormula('a=b')).toBe(false);
    expect(looksLikeFormula('')).toBe(false);
  });
});
