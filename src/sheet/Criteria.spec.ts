import { describe, expect, it } from 'vitest';

import { criterionOf } from './Criteria';
import { DIV0, NA, type CellValue } from './Values';

/**
 * The criterion grammar on its own.
 *
 * `Functions.spec.ts` asserts it through `SUMIF` and `COUNTIF`, which
 * is how anybody meets it. This file asserts it directly, because it
 * is a language shared by six functions and the edges of a language
 * are where the disagreements between them would live.
 */

const matches = (criterion: CellValue, value: CellValue) => criterionOf(criterion).matches(value);

describe('a criterion with no operator', () => {
  it('is equality', () => {
    expect(matches('North', 'North')).toBe(true);
    expect(matches('North', 'South')).toBe(false);
  });

  it('ignores case, as every comparison in the engine does', () => {
    expect(matches('north', 'NORTH')).toBe(true);
  });

  it('compares a number as a number', () => {
    expect(matches(10, 10)).toBe(true);
    expect(matches('10', 10)).toBe(true);
    // And not as the text of one, or "9" would be more than "10".
    expect(matches('10', 10.0)).toBe(true);
  });

  it('does not count a boolean as the number it coerces to', () => {
    expect(matches(1, true)).toBe(false);
  });
});

describe('a criterion with a comparison', () => {
  it('compares numerically when the operand is a number', () => {
    expect(matches('>10', 11)).toBe(true);
    expect(matches('>10', 9)).toBe(false);
    // The case a text comparison gets wrong.
    expect(matches('>9', 10)).toBe(true);
  });

  it('handles all five orderings', () => {
    expect(matches('>=10', 10)).toBe(true);
    expect(matches('<=10', 10)).toBe(true);
    expect(matches('<10', 10)).toBe(false);
    expect(matches('=10', 10)).toBe(true);
    expect(matches('<>10', 10)).toBe(false);
    expect(matches('<>10', 11)).toBe(true);
  });

  it('compares text when the operand is text', () => {
    expect(matches('>M', 'North')).toBe(true);
    expect(matches('>M', 'East')).toBe(false);
  });
});

/**
 * Blanks, which are where the six functions would otherwise disagree.
 *
 * A blank is an absence and not a zero: `SUMIF(A:A, 0)` over an empty
 * column must be zero, not the whole column.
 */
describe('a criterion against a blank cell', () => {
  it('does not match zero', () => {
    expect(matches(0, null)).toBe(false);
  });

  it('does not match the empty string', () => {
    expect(matches('', null)).toBe(true);
    expect(matches('x', null)).toBe(false);
  });

  /** `"<>"` on its own is how people count the cells with anything in. */
  it('is excluded by a bare not-equal', () => {
    expect(matches('<>', null)).toBe(false);
    expect(matches('<>', 0)).toBe(true);
    expect(matches('<>', 'anything')).toBe(true);
  });
});

describe('wildcards', () => {
  it('match any run of characters', () => {
    expect(matches('N*', 'North')).toBe(true);
    expect(matches('N*', 'South')).toBe(false);
    expect(matches('*th', 'North')).toBe(true);
    expect(matches('*or*', 'North')).toBe(true);
  });

  it('match a single character with a question mark', () => {
    expect(matches('?orth', 'North')).toBe(true);
    expect(matches('?orth', 'orth')).toBe(false);
  });

  it('ignore case like everything else', () => {
    expect(matches('n*', 'North')).toBe(true);
  });

  /** A tilde makes one literal, for the sheets that hold file globs. */
  it('can be escaped with a tilde', () => {
    expect(matches('N~*', 'N*')).toBe(true);
    expect(matches('N~*', 'North')).toBe(false);
  });

  it('are anchored to the whole value, not a substring', () => {
    expect(matches('orth', 'North')).toBe(false);
  });

  /** A regular-expression character in a criterion is a literal. */
  it('treat a dot as a dot', () => {
    expect(matches('a.c', 'abc')).toBe(false);
    expect(matches('a.c', 'a.c')).toBe(true);
  });

  it('work with not-equal too', () => {
    expect(matches('<>N*', 'North')).toBe(false);
    expect(matches('<>N*', 'South')).toBe(true);
  });
});

describe('a criterion and an error', () => {
  it('does not match an error against an ordinary criterion', () => {
    expect(matches('>10', DIV0)).toBe(false);
    expect(matches(10, DIV0)).toBe(false);
  });

  it('matches an error against the same error', () => {
    expect(matches(DIV0, DIV0)).toBe(true);
    expect(matches(DIV0, NA)).toBe(false);
  });
});

/**
 * A tilde before anything that is not a wildcard is just a tilde.
 *
 * Excel's rule, and the one that keeps `"~"` usable in the ordinary
 * text people actually have in columns.
 */
describe('a tilde that is not escaping anything', () => {
  it('is a literal character', () => {
    expect(matches('a~b', 'a~b')).toBe(true);
    expect(matches('a~b', 'ab')).toBe(false);
  });
});
