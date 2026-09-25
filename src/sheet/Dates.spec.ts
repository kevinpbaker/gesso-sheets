import { describe, expect, it } from 'vitest';

import { dateOfSerial, parseTypedDate, serialOfDate, serialOfTime, timeOfSerial, weekdayOf } from './Dates';
import { nowSerial } from './Functions';

/**
 * The serial numbers, against Excel.
 *
 * These are the only assertions in this file that are not about this
 * code's own consistency, and they are the ones that matter: a serial
 * that disagrees with Excel's makes every file that leaves this
 * application wrong by a day, silently, forever.
 */
describe('the serial a date has', () => {
  const cases: readonly [string, number][] = [
    ['1900-01-01', 1],
    // The last day before the phantom 1900-02-29.
    ['1900-02-28', 59],
    // And the first one after it: 60 is the day that never happened.
    ['1900-03-01', 61],
    ['1999-12-31', 36_525],
    ['2000-01-01', 36_526],
    // The round number worth having, quoted in `Format.ts` for years.
    ['2023-03-15', 45_000],
    ['2026-09-24', 46_289]
  ];

  it.each(cases)('reads %s as %i', (text, serial) => {
    const [year, month, day] = text.split('-').map(Number);
    expect(serialOfDate(year, month, day)).toBe(serial);
  });

  it.each(cases)('writes %s back from %i', (text, serial) => {
    const { year, month, day } = dateOfSerial(serial);
    expect(`${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`).toBe(text);
  });

  /**
   * The one day this application refuses to invent. Excel's 60 is
   * 1900-02-29; here it reads back as the 28th, the same as 59.
   */
  it('declines to invent the day that never happened', () => {
    expect(dateOfSerial(60)).toEqual({ year: 1900, month: 2, day: 28 });
  });

  /** What makes `=B2-B1` a number of days. */
  it('counts days as the difference between two serials', () => {
    expect(serialOfDate(2026, 9, 24) - serialOfDate(2026, 9, 1)).toBe(23);
    expect(serialOfDate(2025, 1, 1) - serialOfDate(2024, 1, 1)).toBe(366);
  });

  /**
   * Rolling over is the mechanism, not leniency: `EOMONTH` is built
   * out of "day zero of the next month".
   */
  it('rolls a month past twelve into the next year', () => {
    expect(dateOfSerial(serialOfDate(2026, 13, 1))).toEqual({ year: 2027, month: 1, day: 1 });
  });

  it('rolls day zero back to the end of the month before', () => {
    expect(dateOfSerial(serialOfDate(2026, 3, 0))).toEqual({ year: 2026, month: 2, day: 28 });
    expect(dateOfSerial(serialOfDate(2024, 3, 0))).toEqual({ year: 2024, month: 2, day: 29 });
  });

  /**
   * Excel's documented rule, and the only one available: the epoch is
   * 1899-12-30, so there is no serial for the year 26 to give back.
   */
  it('reads a year under 1900 as 1900 plus it', () => {
    expect(dateOfSerial(serialOfDate(26, 1, 1)).year).toBe(1926);
    expect(dateOfSerial(serialOfDate(0, 1, 1)).year).toBe(1900);
  });

  /** Which is also why typing one is refused rather than shifted. */
  it('refuses a typed year it cannot hold', () => {
    expect(parseTypedDate('1800-01-01')).toBeNull();
  });

  it('numbers Sunday one', () => {
    // 2026-09-24 is a Thursday.
    expect(weekdayOf(serialOfDate(2026, 9, 24))).toBe(5);
    expect(weekdayOf(serialOfDate(2026, 9, 27))).toBe(1);
  });
});

describe('the fraction a time is', () => {
  it('makes noon half a day', () => {
    expect(serialOfTime(12, 0, 0)).toBe(0.5);
  });

  it('makes midnight nothing', () => {
    expect(serialOfTime(0, 0, 0)).toBe(0);
  });

  it('reads a clock back out of a fraction', () => {
    expect(timeOfSerial(0.5)).toEqual({ hours: 12, minutes: 0, seconds: 0 });
    expect(timeOfSerial(serialOfTime(13, 45, 30))).toEqual({ hours: 13, minutes: 45, seconds: 30 });
  });

  it('reads the clock of a date that carries one', () => {
    const serial = serialOfDate(2026, 9, 24) + serialOfTime(9, 30, 0);
    expect(dateOfSerial(serial)).toEqual({ year: 2026, month: 9, day: 24 });
    expect(timeOfSerial(serial)).toEqual({ hours: 9, minutes: 30, seconds: 0 });
  });
});

describe('typing a date', () => {
  const serialOf = (text: string) => parseTypedDate(text)?.serial;

  it('reads the unambiguous order', () => {
    expect(parseTypedDate('2026-09-24')).toEqual({ serial: 46_289, date: 'ymd', time: null });
  });

  it('keeps the shape it was typed in', () => {
    expect(parseTypedDate('24 Sep 2026')?.date).toBe('dmy');
    expect(parseTypedDate('Sep 24, 2026')?.date).toBe('mdy');
    expect(parseTypedDate('2026-09-24')?.date).toBe('ymd');
  });

  it('reads a written month either way round', () => {
    expect(serialOf('24 Sep 2026')).toBe(46_289);
    expect(serialOf('Sep 24, 2026')).toBe(46_289);
    expect(serialOf('24 September 2026')).toBe(46_289);
    expect(serialOf('September 24, 2026')).toBe(46_289);
  });

  /** Somebody has to lose, and with no locale there is no evidence. */
  it('reads an ambiguous slashed date month first', () => {
    expect(parseTypedDate('3/4/2026')).toEqual({ serial: serialOfDate(2026, 3, 4), date: 'mdy', time: null });
  });

  /** Unless the other reading is not a date at all. */
  it('reads an unambiguous one day first', () => {
    expect(parseTypedDate('24/9/2026')).toEqual({ serial: 46_289, date: 'dmy', time: null });
  });

  it('follows the convention on two-digit years', () => {
    expect(dateOfSerial(serialOf('1/1/29')!).year).toBe(2029);
    expect(dateOfSerial(serialOf('1/1/30')!).year).toBe(1930);
  });

  /**
   * A year is required. Excel reads `1/2` as this January, which makes
   * the value depend on the day it was typed.
   */
  it('refuses a date with no year', () => {
    expect(parseTypedDate('1/2')).toBeNull();
  });

  it('refuses a day that is not in the month', () => {
    expect(parseTypedDate('2026-02-30')).toBeNull();
    expect(parseTypedDate('2026-04-31')).toBeNull();
    expect(parseTypedDate('2026-13-01')).toBeNull();
  });

  it('accepts the 29th of February in a leap year and not otherwise', () => {
    expect(parseTypedDate('2024-02-29')).not.toBeNull();
    expect(parseTypedDate('2026-02-29')).toBeNull();
  });

  it('reads a clock on its own', () => {
    expect(parseTypedDate('13:45')).toEqual({ serial: serialOfTime(13, 45, 0), date: null, time: 'hm' });
    expect(parseTypedDate('13:45:30')).toEqual({ serial: serialOfTime(13, 45, 30), date: null, time: 'hms' });
  });

  it('reads the twelve-hour clock', () => {
    expect(serialOf('1:45 PM')).toBe(serialOfTime(13, 45, 0));
    expect(serialOf('12:30 AM')).toBe(serialOfTime(0, 30, 0));
    expect(serialOf('12:30 PM')).toBe(serialOfTime(12, 30, 0));
  });

  it('reads a date and a time together', () => {
    expect(parseTypedDate('2026-09-24 09:30')).toEqual({
      serial: 46_289 + serialOfTime(9, 30, 0),
      date: 'ymd',
      time: 'hm'
    });
  });

  it('refuses a clock that is not one', () => {
    expect(parseTypedDate('25:00')).toBeNull();
    expect(parseTypedDate('12:60')).toBeNull();
  });

  /** Everything that is not a date has to stay what it was typed as. */
  it.each(['', 'North', '120', '1-2', 'A1:B2', '3.14', 'Q3 2026', '2026', '-5'])(
    'leaves %o alone',
    text => {
      expect(parseTypedDate(text)).toBeNull();
    }
  );
});

/**
 * The clock the application actually runs on.
 *
 * Every other spec injects a clock, which is what makes `TODAY()`
 * assertable — and leaves the real one, the one that ships, with
 * nothing asserting it at all. These are the two claims worth making
 * about it: that it agrees with the rest of the date arithmetic, and
 * that the day it reports is the local one.
 */
describe('the clock that ships', () => {
  it('agrees with the serial arithmetic about what day it is', () => {
    const at = new Date();
    expect(Math.floor(nowSerial())).toBe(serialOfDate(at.getFullYear(), at.getMonth() + 1, at.getDate()));
  });

  /**
   * Local, not UTC, and this is the one place in the file that is.
   * A UTC `TODAY()` is yesterday all evening for anybody east of
   * Greenwich enough to notice.
   */
  it('reports the day on the local wall, not in UTC', () => {
    const at = new Date();
    const { year, month, day } = dateOfSerial(nowSerial());
    expect({ year, month, day }).toEqual({
      year: at.getFullYear(),
      month: at.getMonth() + 1,
      day: at.getDate()
    });
  });

  it('carries the time of day in the fraction', () => {
    const fraction = nowSerial() - Math.floor(nowSerial());
    expect(fraction).toBeGreaterThanOrEqual(0);
    expect(fraction).toBeLessThan(1);
  });
});
