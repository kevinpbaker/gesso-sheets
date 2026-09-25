import { describe, expect, it } from 'vitest';

import { functionNames } from './Functions';
import { completionsFor, signatureOf, SIGNATURES } from './Signatures';

/**
 * The hints the formula editor shows, and the guard that keeps them
 * honest.
 *
 * A signature table separate from the implementations is two places
 * to change when a function's arguments change. The last spec here is
 * what makes that survivable: a function in the library with no
 * signature fails the build, as does a signature for a function that
 * does not exist.
 */
describe('the signature table', () => {
  it('has one for every function the sheet knows', () => {
    const missing = functionNames().filter(name => signatureOf(name) === null);
    expect(missing).toEqual([]);
  });

  it('has none for a function that does not exist', () => {
    const known = new Set(functionNames());
    expect(Object.keys(SIGNATURES).filter(name => !known.has(name))).toEqual([]);
  });

  it('writes optional arguments in brackets and repeats as an ellipsis', () => {
    expect(signatureOf('ROUND')?.args).toEqual(['number', '[places]']);
    expect(signatureOf('SUM')?.args).toEqual(['number', '…']);
    expect(signatureOf('SUM')?.repeats).toBe(true);
  });

  it('gives a no-argument function an empty list rather than nothing', () => {
    expect(signatureOf('TODAY')?.args).toEqual([]);
  });

  it('answers whatever case the name was typed in', () => {
    expect(signatureOf('sum')).toBe(signatureOf('SUM'));
  });

  it('says so about the volatile ones, which is the surprising part', () => {
    for (const name of ['RAND', 'RANDBETWEEN', 'NOW', 'TODAY']) {
      expect(signatureOf(name)?.summary).toContain('every edit');
    }
  });

  /** Every summary is one line somebody can read at speed. */
  it('keeps every summary short', () => {
    const long = Object.entries(SIGNATURES).filter(([, signature]) => signature.summary.length > 80);
    expect(long.map(([name]) => name)).toEqual([]);
  });
});

describe('completing a function name', () => {
  it('offers the names that start with what was typed', () => {
    expect(completionsFor('SUM')).toEqual(['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT']);
  });

  it('ignores the case it was typed in', () => {
    expect(completionsFor('su')).toEqual(completionsFor('SU'));
  });

  /**
   * Prefix, not substring: a list that offered `COUNTIF` for `IF`
   * would bury the function actually being typed under its cousins.
   */
  it('matches the start rather than anywhere', () => {
    expect(completionsFor('IF')).toEqual(['IF', 'IFERROR', 'IFNA', 'IFS']);
  });

  it('offers nothing for nothing', () => {
    expect(completionsFor('')).toEqual([]);
  });

  it('offers nothing for a name it does not know', () => {
    expect(completionsFor('ZZZ')).toEqual([]);
  });

  it('offers the same list in the same order every time', () => {
    expect(completionsFor('CO')).toEqual(completionsFor('CO'));
    expect(completionsFor('CO')).toEqual([...completionsFor('CO')].sort());
  });
});
