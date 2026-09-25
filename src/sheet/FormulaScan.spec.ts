import { describe, expect, it } from 'vitest';

import { formatRef } from './A1';
import { callAt, matchingBracket, referenceAt, scanFormula } from './FormulaScan';

/**
 * Reading a formula that is still being typed.
 *
 * Every case here is deliberately *unfinished* somewhere, because
 * that is the state the formula editor works in: the parser's
 * question is "what does this mean" and half the time the answer is
 * "nothing yet".
 */

const refs = (text: string) =>
  scanFormula(text).references.map(r => ({
    text: text.slice(r.start, r.end),
    at: [r.start, r.end],
    cells: r.isRange ? `${formatRef(r.from)}:${formatRef(r.to)}` : formatRef(r.from)
  }));

describe('finding the references in a formula', () => {
  it('finds a single cell', () => {
    expect(refs('=A1')).toEqual([{ text: 'A1', at: [1, 3], cells: 'A1' }]);
  });

  it('finds several', () => {
    expect(refs('=B2*C2').map(r => r.text)).toEqual(['B2', 'C2']);
  });

  /** A range is one reference, not two with a gap in the middle. */
  it('keeps a range whole', () => {
    expect(refs('=SUM(A1:B9)')).toEqual([{ text: 'A1:B9', at: [5, 10], cells: 'A1:B9' }]);
  });

  it('keeps the dollars somebody wrote', () => {
    expect(refs('=$A$1')).toEqual([{ text: '$A$1', at: [1, 5], cells: '$A$1' }]);
  });

  it('finds a whole-column reference', () => {
    expect(refs('=SUM(A:C)')[0]).toMatchObject({ text: 'A:C', at: [5, 8] });
  });

  /** `LOG10` is a cell address and a function name; the `(` decides. */
  it('does not mistake a function name for a cell', () => {
    expect(refs('=LOG10(2)')).toEqual([]);
    expect(refs('=LOG10')).toEqual([{ text: 'LOG10', at: [1, 6], cells: 'LOG10' }]);
  });

  it('finds nothing in text that is not a formula', () => {
    expect(refs('A1')).toEqual([]);
    expect(refs('120')).toEqual([]);
  });

  /**
   * The case that matters most: this runs on every keystroke, and a
   * formula halfway through being typed must not throw.
   */
  it('reads an unfinished formula without complaining', () => {
    expect(refs('=SUM(A1:')).toEqual([{ text: 'A1', at: [5, 7], cells: 'A1' }]);
    expect(refs('=B2+')).toEqual([{ text: 'B2', at: [1, 3], cells: 'B2' }]);
    expect(refs('=SUM(')).toEqual([]);
  });

  it('gives up quietly on text it cannot even tokenize', () => {
    // An unclosed string is a throw in the tokenizer.
    expect(refs('="abc')).toEqual([]);
  });
});

describe('the reference under the caret', () => {
  const at = (text: string, caret: number) => {
    const found = referenceAt(scanFormula(text), caret);
    return found === null ? null : text.slice(found.start, found.end);
  };

  it('is found from inside it', () => {
    expect(at('=B2*C2', 2)).toBe('B2');
  });

  /** Both edges count, so the caret just after `B2` is still on it. */
  it('is found from either edge', () => {
    expect(at('=B2*C2', 1)).toBe('B2');
    expect(at('=B2*C2', 3)).toBe('B2');
  });

  it('is nothing when the caret is elsewhere', () => {
    expect(at('=B2 + 1', 6)).toBeNull();
  });
});

describe('the call the caret is inside', () => {
  const call = (text: string, caret: number) => {
    const found = callAt(scanFormula(text), caret);
    return found === null ? null : `${found.name}#${found.argument}`;
  };

  it('is the function whose brackets the caret is in', () => {
    expect(call('=SUM(A1)', 5)).toBe('SUM#0');
  });

  it('counts the commas to find the argument', () => {
    expect(call('=ROUND(A1, 2)', 11)).toBe('ROUND#1');
    expect(call('=IF(A1, 2, 3)', 11)).toBe('IF#2');
  });

  /** The normal case while typing: the bracket is not closed yet. */
  it('works with the closing bracket missing', () => {
    expect(call('=SUM(A1, ', 9)).toBe('SUM#1');
  });

  it('is the innermost call when they are nested', () => {
    expect(call('=SUM(ROUND(A1, 2), 3)', 15)).toBe('ROUND#1');
  });

  it('goes back to the outer one past the inner bracket', () => {
    expect(call('=SUM(ROUND(A1, 2), 3)', 19)).toBe('SUM#1');
  });

  /**
   * A bare `(` is grouping, not a call — and a comma inside it must
   * not be counted as an argument of the function outside it.
   */
  it('sees through a bracket that is only grouping', () => {
    expect(call('=SUM((1+2), 3)', 12)).toBe('SUM#1');
  });

  it('is nothing outside any call', () => {
    expect(call('=A1+1', 3)).toBeNull();
  });

  it('is nothing once the call is closed', () => {
    expect(call('=SUM(A1) + ', 11)).toBeNull();
  });
});

describe('matching brackets', () => {
  const pair = (text: string, caret: number) => {
    const found = matchingBracket(scanFormula(text), caret);
    return found === null ? null : [found.here.start, found.there.start];
  };

  /** Just after a closing bracket, which is where you have just typed. */
  it('matches backwards from the bracket before the caret', () => {
    expect(pair('=SUM(A1)', 8)).toEqual([7, 4]);
  });

  it('matches forwards from the bracket after the caret', () => {
    expect(pair('=SUM(A1)', 4)).toEqual([4, 7]);
  });

  it('pairs through nesting', () => {
    expect(pair('=SUM(ROUND(A1, 2))', 18)).toEqual([17, 4]);
    expect(pair('=SUM(ROUND(A1, 2))', 17)).toEqual([16, 10]);
  });

  it('is nothing when the caret is not on a bracket', () => {
    expect(pair('=SUM(A1)', 6)).toBeNull();
  });

  /** An unclosed bracket has no partner, and saying so is the answer. */
  it('is nothing when the bracket was never closed', () => {
    expect(pair('=SUM(A1', 4)).toBeNull();
  });
});
