import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import {
  covers,
  isEmptyPaint,
  matches,
  overlay,
  scaleColour,
  SCALE_STEPS,
  type ColourScale
} from './Conditional';
import { DIV0 } from './Values';

/**
 * The rules, on their own.
 *
 * Everything here is a question about one value and one rule, which
 * is what lets a conditional format be resolved at the window: the
 * expensive half is reading the cells, and reading the cells is the
 * caller's job.
 */

const area = (row: number, column: number, lastRow: number, lastColumn: number) => ({
  start: relativeRef(row, column),
  end: relativeRef(lastRow, lastColumn)
});

describe('what a test asks of a value', () => {
  it('compares numbers', () => {
    expect(matches({ kind: 'greaterThan', value: 5 }, 6)).toBe(true);
    expect(matches({ kind: 'greaterThan', value: 5 }, 5)).toBe(false);
    expect(matches({ kind: 'lessThan', value: 5 }, 4)).toBe(true);
    expect(matches({ kind: 'between', low: 1, high: 3 }, 3)).toBe(true);
    expect(matches({ kind: 'between', low: 1, high: 3 }, 4)).toBe(false);
  });

  /** A number rule is about numbers, and text is not one. */
  it('does not read text as a number', () => {
    expect(matches({ kind: 'greaterThan', value: 5 }, '9')).toBe(false);
  });

  it('compares text without minding the case', () => {
    expect(matches({ kind: 'equalTo', value: 'north' }, 'North')).toBe(true);
    expect(matches({ kind: 'textContains', text: 'or' }, 'North')).toBe(true);
    expect(matches({ kind: 'textContains', text: 'or' }, 'South')).toBe(false);
  });

  it('knows an empty cell from a full one', () => {
    expect(matches({ kind: 'isEmpty' }, null)).toBe(true);
    expect(matches({ kind: 'isEmpty' }, '')).toBe(true);
    expect(matches({ kind: 'isEmpty' }, 0)).toBe(false);
    expect(matches({ kind: 'notEmpty' }, 0)).toBe(true);
  });

  /**
   * An error is not greater than five and not less than it either.
   * Only the emptiness tests have an honest answer about a cell that
   * is showing `#DIV/0!`.
   */
  it('refuses to compare an error', () => {
    expect(matches({ kind: 'greaterThan', value: 5 }, DIV0)).toBe(false);
    expect(matches({ kind: 'lessThan', value: 5 }, DIV0)).toBe(false);
    expect(matches({ kind: 'isEmpty' }, DIV0)).toBe(false);
    expect(matches({ kind: 'notEmpty' }, DIV0)).toBe(true);
  });

  /** It needs an evaluator and a sheet, so it is not answered here. */
  it('says a formula is not its question', () => {
    expect(matches({ kind: 'formula', input: '=A1>5' }, 9)).toBeNull();
  });
});

describe('a colour scale', () => {
  const scale: ColourScale = { from: '#ffffff', to: '#000000' };

  it('puts the smallest at one end and the largest at the other', () => {
    expect(scaleColour(scale, 0, { low: 0, high: 10 })).toBe('#ffffff');
    expect(scaleColour(scale, 10, { low: 0, high: 10 })).toBe('#000000');
  });

  it('puts the middle in the middle', () => {
    expect(scaleColour(scale, 5, { low: 0, high: 10 })).toBe('#808080');
  });

  it('takes a middle stop when it is given one', () => {
    const three: ColourScale = { from: '#ff0000', middle: '#ffffff', to: '#0000ff' };
    expect(scaleColour(three, 0, { low: 0, high: 10 })).toBe('#ff0000');
    expect(scaleColour(three, 5, { low: 0, high: 10 })).toBe('#ffffff');
    expect(scaleColour(three, 10, { low: 0, high: 10 })).toBe('#0000ff');
  });

  it('clamps a value from outside the extent', () => {
    expect(scaleColour(scale, -4, { low: 0, high: 10 })).toBe('#ffffff');
    expect(scaleColour(scale, 99, { low: 0, high: 10 })).toBe('#000000');
  });

  /** Every cell is both the smallest and the largest; neither end is honest. */
  it('takes the middle when there is nothing to spread between', () => {
    expect(scaleColour(scale, 7, { low: 7, high: 7 })).toBe('#808080');
  });

  it('says nothing about a value that is not a number', () => {
    expect(scaleColour(scale, 'North', { low: 0, high: 10 })).toBeNull();
    expect(scaleColour(scale, null, { low: 0, high: 10 })).toBeNull();
  });

  /**
   * The claim the steps exist for: however many cells a scale covers,
   * it can only ask for so many colours — so the palette that has to
   * carry them is bounded, and a scroll reuses what is already there.
   */
  it('asks for no more colours than it has steps', () => {
    const seen = new Set<string>();
    for (let value = 0; value <= 1_000; value++) {
      seen.add(scaleColour(scale, value, { low: 0, high: 1_000 }) ?? '');
    }
    expect(seen.size).toBeLessThanOrEqual(SCALE_STEPS);
    expect(seen.size).toBe(SCALE_STEPS);
  });

  it('gives the same colour for the same step, so it can be interned', () => {
    const one = scaleColour(scale, 300, { low: 0, high: 1_000 });
    const two = scaleColour(scale, 301, { low: 0, high: 1_000 });
    expect(one).toBe(two);
  });

  it('leaves a colour it cannot read alone', () => {
    expect(scaleColour({ from: 'rebeccapurple', to: '#000000' }, 5, { low: 0, high: 10 })).toBe('rebeccapurple');
  });

  it('reads a three-digit colour too', () => {
    expect(scaleColour({ from: '#fff', to: '#000' }, 0, { low: 0, high: 10 })).toBe('#ffffff');
  });
});

describe('several rules over one cell', () => {
  /**
   * Later rules win field by field, which is what makes "red text
   * for negatives" and "grey fill for the weekend" two rules rather
   * than four.
   */
  it('folds them without one clearing another', () => {
    expect(overlay([{ color: '#ff0000' }, { fill: '#eeeeee' }])).toEqual({
      color: '#ff0000',
      fill: '#eeeeee'
    });
  });

  it('lets a later rule replace a field an earlier one set', () => {
    expect(overlay([{ fill: '#eeeeee' }, { fill: '#ff0000' }]).fill).toBe('#ff0000');
  });

  it('knows when nothing was said at all', () => {
    expect(isEmptyPaint(overlay([]))).toBe(true);
    expect(isEmptyPaint(overlay([{ bold: false }]))).toBe(false);
  });
});

describe('the range a rule covers', () => {
  it('holds its corners and everything between', () => {
    const range = area(1, 1, 3, 3);
    expect(covers(range, 1, 1)).toBe(true);
    expect(covers(range, 3, 3)).toBe(true);
    expect(covers(range, 2, 2)).toBe(true);
    expect(covers(range, 0, 1)).toBe(false);
    expect(covers(range, 4, 1)).toBe(false);
  });

  it('does not mind which corner was given first', () => {
    expect(covers(area(3, 3, 1, 1), 2, 2)).toBe(true);
  });
});
