import { describe, expect, it } from 'vitest';

import {
  DEFAULT_FORMAT,
  formatWith,
  isDefault,
  keyOf,
  PLAIN,
  placesOf,
  withPlaces,
  type CellFormat,
  type NumberFormat
} from './Format';
import { DIV0, REF } from './Values';

const as = (number: NumberFormat): CellFormat => ({ number, paint: PLAIN });

describe('showing a number under a format', () => {
  it('prints a general number as itself', () => {
    expect(formatWith(1234.5, { kind: 'general' })).toBe('1234.5');
  });

  it('fixes the places, and pads short ones', () => {
    expect(formatWith(1234.5, { kind: 'number', places: 2, thousands: false })).toBe('1234.50');
    expect(formatWith(1234.567, { kind: 'number', places: 2, thousands: false })).toBe('1234.57');
    expect(formatWith(7, { kind: 'number', places: 0, thousands: false })).toBe('7');
  });

  it('groups thousands in the integer part only', () => {
    expect(formatWith(1234567.891, { kind: 'number', places: 2, thousands: true })).toBe('1,234,567.89');
    expect(formatWith(999, { kind: 'number', places: 0, thousands: true })).toBe('999');
    expect(formatWith(1000, { kind: 'number', places: 0, thousands: true })).toBe('1,000');
  });

  it('keeps the minus outside the grouping', () => {
    expect(formatWith(-1234.5, { kind: 'number', places: 2, thousands: true })).toBe('-1,234.50');
  });

  /** The accounting convention, which is how a column of figures is read. */
  it('puts negative money in brackets', () => {
    expect(formatWith(1234.5, { kind: 'currency', places: 2, symbol: '$' })).toBe('$1,234.50');
    expect(formatWith(-1234.5, { kind: 'currency', places: 2, symbol: '$' })).toBe('($1,234.50)');
    expect(formatWith(12, { kind: 'currency', places: 2, symbol: '£' })).toBe('£12.00');
  });

  it('multiplies a percentage by a hundred', () => {
    expect(formatWith(0.256, { kind: 'percent', places: 1 })).toBe('25.6%');
    expect(formatWith(1, { kind: 'percent', places: 0 })).toBe('100%');
  });

  /** `1.50E+03`, not `1.50e+3`: the exponent is padded and the E is capital. */
  it('writes scientific the way a spreadsheet writes it', () => {
    expect(formatWith(1500, { kind: 'scientific', places: 2 })).toBe('1.50E+03');
    expect(formatWith(0.0015, { kind: 'scientific', places: 2 })).toBe('1.50E-03');
  });

  it('shows a serial number as a date', () => {
    // Serial 45000 is 2023-03-15 here and in Excel, which is the only
    // property of the epoch worth having.
    expect(formatWith(45_000, { kind: 'date', pattern: 'ymd' })).toBe('2023-03-15');
    expect(formatWith(45_000, { kind: 'date', pattern: 'dmy' })).toBe('15 Mar 2023');
    expect(formatWith(45_000, { kind: 'date', pattern: 'mdy' })).toBe('Mar 15, 2023');
    expect(formatWith(1, { kind: 'date', pattern: 'ymd' })).toBe('1900-01-01');
    expect(formatWith(59, { kind: 'date', pattern: 'ymd' })).toBe('1900-02-28');
    expect(formatWith(61, { kind: 'date', pattern: 'ymd' })).toBe('1900-03-01');
  });

  /**
   * Excel's serial 60 is 1900-02-29, a day that did not happen —
   * Lotus believed 1900 was a leap year and every spreadsheet since
   * has kept the bug so serials move between them. Everything from 61
   * on agrees; the phantom day itself is one this application
   * declines to invent.
   */
  it('does not invent the day that never existed', () => {
    expect(formatWith(60, { kind: 'date', pattern: 'ymd' })).toBe('1900-02-28');
  });

  it('shows the fraction of a day as a time', () => {
    expect(formatWith(45_000.5, { kind: 'time', pattern: 'hm' })).toBe('12:00');
    expect(formatWith(0.25, { kind: 'time', pattern: 'hms' })).toBe('06:00:00');
  });

  describe('what a format may not do', () => {
    /** The format describes a number, and there is no number. */
    it('leaves an error alone', () => {
      expect(formatWith(DIV0, { kind: 'currency', places: 2, symbol: '$' })).toBe('#DIV/0!');
      expect(formatWith(REF, { kind: 'percent', places: 0 })).toBe('#REF!');
    });

    /**
     * Formatting a range must never make what is in it unreadable. A
     * currency column with the word `Total` in it shows `Total`.
     */
    it('leaves text alone', () => {
      expect(formatWith('Total', { kind: 'currency', places: 2, symbol: '$' })).toBe('Total');
    });

    /** A blank cell in a currency column is blank, not `$0.00`. */
    it('leaves an empty cell empty', () => {
      expect(formatWith(null, { kind: 'currency', places: 2, symbol: '$' })).toBe('');
      expect(formatWith(null, { kind: 'percent', places: 2 })).toBe('');
    });

    it('leaves a boolean alone', () => {
      expect(formatWith(true, { kind: 'number', places: 2, thousands: false })).toBe('TRUE');
    });
  });
});

describe('interning a format', () => {
  /**
   * Two formats that are equal must key the same however they were
   * built. `JSON.stringify` depends on the order the fields were
   * written, so a palette keyed by it grows a duplicate the first
   * time a field is set in a different order — and a palette entry is
   * on the wire forever.
   */
  it('keys the same for equal formats built in different orders', () => {
    const one: CellFormat = { number: { kind: 'percent', places: 1 }, paint: { ...PLAIN, bold: true } };
    const two: CellFormat = {
      paint: { ...PLAIN, wrap: false, bold: true },
      number: { places: 1, kind: 'percent' }
    };
    expect(keyOf(one)).toBe(keyOf(two));
  });

  it('keys differently for formats that differ anywhere', () => {
    const base = DEFAULT_FORMAT;
    expect(keyOf({ ...base, paint: { ...PLAIN, bold: true } })).not.toBe(keyOf(base));
    expect(keyOf({ ...base, paint: { ...PLAIN, fill: '#fee' } })).not.toBe(keyOf(base));
    expect(keyOf(as({ kind: 'percent', places: 1 }))).not.toBe(keyOf(as({ kind: 'percent', places: 2 })));
  });

  it('knows the format every unformatted cell has', () => {
    expect(isDefault(DEFAULT_FORMAT)).toBe(true);
    expect(isDefault({ ...DEFAULT_FORMAT, paint: { ...PLAIN, italic: true } })).toBe(false);
  });
});

describe('more and fewer decimal places', () => {
  it('adds and removes places', () => {
    expect(placesOf(withPlaces({ kind: 'number', places: 2, thousands: true }, 1))).toBe(3);
    expect(placesOf(withPlaces({ kind: 'number', places: 2, thousands: true }, -1))).toBe(1);
  });

  it('does not go below none', () => {
    expect(placesOf(withPlaces({ kind: 'number', places: 0, thousands: false }, -1))).toBe(0);
  });

  /**
   * Pressing "more decimals" on an untouched cell is how most people
   * first format anything. A button that did nothing there would be
   * the wrong answer to the commonest use of it.
   */
  it('turns General into a number format on the first press', () => {
    const next = withPlaces({ kind: 'general' }, 1);
    expect(next.kind).toBe('number');
    expect(placesOf(next)).toBe(1);
  });

  it('leaves a format with no places to change alone', () => {
    expect(withPlaces({ kind: 'text' }, 1)).toEqual({ kind: 'text' });
    expect(withPlaces({ kind: 'general' }, -1)).toEqual({ kind: 'general' });
  });
});
