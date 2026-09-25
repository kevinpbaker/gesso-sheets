import { describe, expect, it } from 'vitest';

import { serialOfDate } from './Dates';
import { DIV0 } from './Values';
import { todaySerial, validate } from './Validation';

/**
 * What a cell is allowed to hold.
 *
 * One value and one rule, the same shape the conditional formats
 * have — which is what lets both be resolved at the window.
 */

describe('a list', () => {
  const rule = { kind: 'list' as const, values: ['North', 'South', 'East', 'West'] };

  it('takes one of its values', () => {
    expect(validate(rule, 'North')).toBeNull();
  });

  it('does not mind the case it was typed in', () => {
    expect(validate(rule, 'north')).toBeNull();
  });

  it('says what the values are when it refuses', () => {
    expect(validate(rule, 'Noth')).toContain('North, South, East, West');
  });
});

describe('a number', () => {
  it('holds a range', () => {
    const rule = { kind: 'number' as const, min: 1, max: 5 };
    expect(validate(rule, 3)).toBeNull();
    expect(validate(rule, 0)).toContain('below 1');
    expect(validate(rule, 6)).toContain('above 5');
  });

  it('takes an open end', () => {
    expect(validate({ kind: 'number', min: 0 }, 1_000_000)).toBeNull();
  });

  it('asks for a whole one when it says so', () => {
    expect(validate({ kind: 'number', integer: true }, 2.5)).toContain('whole number');
    expect(validate({ kind: 'number', integer: true }, 2)).toBeNull();
  });

  it('refuses text', () => {
    expect(validate({ kind: 'number' }, 'lots')).toContain('number is wanted');
  });
});

describe('text', () => {
  it('holds a length', () => {
    expect(validate({ kind: 'text', maxLength: 3 }, 'abcd')).toContain('At most 3');
    expect(validate({ kind: 'text', maxLength: 3 }, 'abc')).toBeNull();
  });
});

describe('a date', () => {
  const from = serialOfDate(2026, 1, 1);
  const to = serialOfDate(2026, 12, 31);
  const rule = { kind: 'date' as const, from, to };

  it('takes a serial inside the range', () => {
    expect(validate(rule, serialOfDate(2026, 6, 1))).toBeNull();
  });

  it('reads text the way somebody typing it would be read', () => {
    expect(validate(rule, '2026-06-01')).toBeNull();
    expect(validate(rule, '2025-06-01')).toContain('Earlier');
  });

  it('refuses something that is not a date at all', () => {
    expect(validate(rule, 'soon')).toContain('date is wanted');
  });

  it('knows what day it is, for a rule written against today', () => {
    const now = new Date(Date.UTC(2026, 8, 25));
    expect(validate({ kind: 'date', from: todaySerial(now) }, '2026-09-25')).toBeNull();
  });
});

/**
 * Two things a rule deliberately says nothing about.
 *
 * Emptying a cell is how somebody takes back a mistake, and a
 * validation that refused it would be a cell nobody could clear. A
 * formula that is broken is a different problem, and the cell is
 * already saying so.
 */
describe('what a rule leaves alone', () => {
  const rule = { kind: 'number' as const, min: 10 };

  it('lets a cell be emptied', () => {
    expect(validate(rule, null)).toBeNull();
    expect(validate(rule, '')).toBeNull();
  });

  it('says nothing about a cell that is showing an error', () => {
    expect(validate(rule, DIV0)).toBeNull();
  });
});
