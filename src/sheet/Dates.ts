import type { DatePattern, TimePattern } from './Format';

/**
 * Dates, which in a spreadsheet are numbers that have been told what
 * they are.
 *
 * There is no date *type* here and there must not be one. A date is a
 * count of days from an epoch, and a format that says to draw it as a
 * calendar date rather than as 45,000 — which is why `=B2-B1` gives
 * the number of days between two dates, why a date sorts with the
 * numbers, and why `EOMONTH` can return something you can still add 7
 * to. A separate type would break all three and buy nothing.
 *
 * Phase 9 could already *show* one. This is the other half: text that
 * someone typed becoming a serial and a format, and the functions that
 * do arithmetic on days. It is the single point where the engine and
 * the format axis meet, which is why it waited for both.
 *
 * ## The epoch, and the day that never happened
 *
 * Day zero is 1899-12-30, not 1900-01-01, because Lotus 1-2-3 believed
 * 1900 was a leap year and every spreadsheet since has kept the bug so
 * that serial numbers move between them. Serial 45,000 is 2023-03-15
 * here and in Excel, which is the only property worth having.
 *
 * The bug is one phantom day: Excel's serial 60 is 1900-02-29, which
 * did not happen. So serials 1..59 are a day ahead of what the epoch
 * alone would give, and serial 60 is a day this application declines
 * to invent — it reads back as 1900-02-28, the same as 59. One wrong
 * day in 1900, against reproducing a calendar error on purpose.
 */

export const DAY_MS = 86_400_000;
export const SECONDS_PER_DAY = 86_400;

/** 1899-12-30, as UTC milliseconds. */
const EPOCH_UTC = Date.UTC(1899, 11, 30);
/** Serials below this predate the phantom day and are shifted by one. */
const SHIFT_BEFORE = 60;

/**
 * Built in UTC throughout, everywhere in this file.
 *
 * A local-time construction shifts the displayed day for anybody west
 * of Greenwich for part of the year, so two people would see different
 * dates in the same file — the one thing a date must never do.
 */
function utcDaysOf(serial: number): number {
  return serial < SHIFT_BEFORE ? serial + 1 : serial;
}

function serialOfUtcDays(days: number): number {
  return days <= SHIFT_BEFORE ? days - 1 : days;
}

export interface DateParts {
  readonly year: number;
  /** 1–12, as a person counts months and `MONTH` returns them. */
  readonly month: number;
  readonly day: number;
}

export interface TimeParts {
  readonly hours: number;
  readonly minutes: number;
  readonly seconds: number;
}

/** The calendar date a serial's whole-day part names. */
export function dateOfSerial(serial: number): DateParts {
  const at = new Date(EPOCH_UTC + utcDaysOf(Math.floor(serial)) * DAY_MS);
  return { year: at.getUTCFullYear(), month: at.getUTCMonth() + 1, day: at.getUTCDate() };
}

/**
 * The serial for a calendar date.
 *
 * Out-of-range months and days roll over, as `DATE` is defined to do
 * in every spreadsheet: `DATE(2026, 13, 1)` is January 2027 and
 * `DATE(2026, 3, 0)` is the last day of February. That is not
 * leniency, it is the mechanism `EOMONTH` and `EDATE` are built out
 * of, and code that clamped instead would quietly give February 28th
 * for the day before March 1st in a leap year.
 */
export function serialOfDate(year: number, month: number, day: number): number {
  const at = new Date(0);
  // A year under 1900 means 1900 plus it, which is the rule Excel's
  // `DATE` documents and the only one that can be right here: the
  // epoch is 1899-12-30, so there are no serials for the first
  // century AD to give back. `DATE(26, 1, 1)` is 1926.
  at.setUTCFullYear(year < 1900 ? year + 1900 : year, month - 1, day);
  return serialOfUtcDays(Math.round((at.getTime() - EPOCH_UTC) / DAY_MS));
}

/** The clock time a serial's fractional part names. */
export function timeOfSerial(serial: number): TimeParts {
  const ofDay = serial - Math.floor(serial);
  // Rounded to the second before splitting, so 23:59:59.7 reads as the
  // next midnight rather than as 23:59:60.
  const total = Math.round(ofDay * SECONDS_PER_DAY) % SECONDS_PER_DAY;
  return {
    hours: Math.floor(total / 3600),
    minutes: Math.floor((total % 3600) / 60),
    seconds: total % 60
  };
}

/** A clock time as the fraction of a day it is. */
export function serialOfTime(hours: number, minutes: number, seconds: number): number {
  return (hours * 3600 + minutes * 60 + seconds) / SECONDS_PER_DAY;
}

/** Sunday is 1, as `WEEKDAY`'s default numbering has it. */
export function weekdayOf(serial: number): number {
  const at = new Date(EPOCH_UTC + utcDaysOf(Math.floor(serial)) * DAY_MS);
  return at.getUTCDay() + 1;
}

const MONTH_NAMES = [
  'JANUARY',
  'FEBRUARY',
  'MARCH',
  'APRIL',
  'MAY',
  'JUNE',
  'JULY',
  'AUGUST',
  'SEPTEMBER',
  'OCTOBER',
  'NOVEMBER',
  'DECEMBER'
];

/** A month name or its three-letter abbreviation, 1–12, or null. */
function monthNumber(name: string): number | null {
  const upper = name.toUpperCase();
  const at = MONTH_NAMES.findIndex(full => full === upper || full.slice(0, 3) === upper);
  return at === -1 ? null : at + 1;
}

/**
 * What typing a date gives you: a number, and the format to show it
 * with.
 *
 * The format travels with the value because typing `2026-09-24` and
 * getting 46,289 would be absurd, and because the *shape somebody
 * typed* is the only evidence available about how they want it shown.
 * Type it with slashes and you get it back with slashes.
 */
export interface TypedDate {
  readonly serial: number;
  /** `null` when only a time was typed, which needs no date pattern. */
  readonly date: DatePattern | null;
  /** `null` when no time was typed. */
  readonly time: TimePattern | null;
}

const ISO = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const SLASHED = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/;
const DAY_FIRST = /^(\d{1,2})[ -]([A-Za-z]{3,9})\.?,?[ -](\d{4})$/;
const MONTH_FIRST = /^([A-Za-z]{3,9})\.?[ -](\d{1,2}),?[ -](\d{4})$/;
const CLOCK = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp])\.?[Mm]\.?)?$/;

/**
 * A typed date, a typed time, or null for anything else.
 *
 * Three decisions worth stating, because each one is a place a
 * spreadsheet can be quietly wrong.
 *
 * **A year is required.** Excel reads `1/2` as the second of January
 * *this year*, which makes the value depend on the day it was typed
 * and silently wrong when the file is reopened in January. Requiring
 * the year costs four keystrokes and removes the class.
 *
 * **`3/4/2026` is March the fourth.** Somebody has to lose: with no
 * locale there is no evidence in the text, and month-first is what an
 * unlocalised spreadsheet does. The ambiguity only exists when both
 * numbers are 12 or under — `24/9/2026` is unambiguous and is read as
 * the day first, whatever the general rule says, because the other
 * reading is not a date at all.
 *
 * **Two-digit years follow the spreadsheet convention**: 00–29 are
 * 2000s and 30–99 are 1900s, which is what every other sheet does and
 * is therefore the only answer that round-trips.
 */
export function parseTypedDate(text: string): TypedDate | null {
  const trimmed = text.trim();
  if (trimmed === '') {
    return null;
  }

  // A date and a time together, split on the gap between them.
  const gap = trimmed.search(/\s+(?=\d{1,2}:)/);
  if (gap !== -1) {
    const day = parseDatePart(trimmed.slice(0, gap));
    const clock = parseTimePart(trimmed.slice(gap).trim());
    if (day !== null && clock !== null) {
      return { serial: day.serial + clock.fraction, date: day.pattern, time: clock.pattern };
    }
    return null;
  }

  const day = parseDatePart(trimmed);
  if (day !== null) {
    return { serial: day.serial, date: day.pattern, time: null };
  }
  const clock = parseTimePart(trimmed);
  if (clock !== null) {
    return { serial: clock.fraction, date: null, time: clock.pattern };
  }
  return null;
}

function parseDatePart(text: string): { serial: number; pattern: DatePattern } | null {
  const iso = ISO.exec(text);
  if (iso !== null) {
    return build(Number(iso[1]), Number(iso[2]), Number(iso[3]), 'ymd');
  }

  const slashed = SLASHED.exec(text);
  if (slashed !== null) {
    const first = Number(slashed[1]);
    const second = Number(slashed[2]);
    const year = fullYear(Number(slashed[3]), slashed[3].length);
    // The unambiguous case wins over the general rule: a first number
    // past twelve can only be a day.
    const dayFirst = first > 12;
    return dayFirst ? build(year, second, first, 'dmy') : build(year, first, second, 'mdy');
  }

  const dayFirst = DAY_FIRST.exec(text);
  if (dayFirst !== null) {
    const month = monthNumber(dayFirst[2]);
    return month === null ? null : build(Number(dayFirst[3]), month, Number(dayFirst[1]), 'dmy');
  }

  const monthFirst = MONTH_FIRST.exec(text);
  if (monthFirst !== null) {
    const month = monthNumber(monthFirst[1]);
    return month === null ? null : build(Number(monthFirst[3]), month, Number(monthFirst[2]), 'mdy');
  }

  return null;
}

/**
 * Validated rather than rolled over, which is the opposite of `DATE`.
 *
 * `DATE(2026, 13, 1)` is deliberately January 2027 because that is how
 * `EOMONTH` is built. Somebody *typing* `2026-13-01` has made a
 * mistake, and a cell that silently became next January would hide it.
 */
function build(year: number, month: number, day: number, pattern: DatePattern): { serial: number; pattern: DatePattern } | null {
  if (month < 1 || month > 12 || day < 1 || day > 31 || year < 1 || year > 9999) {
    return null;
  }
  const serial = serialOfDate(year, month, day);
  const back = dateOfSerial(serial);
  // A rolled-over day — the 31st of a thirty-day month — comes back as
  // a different date than it went in as, which is how this notices.
  if (back.year !== year || back.month !== month || back.day !== day) {
    return null;
  }
  return { serial, pattern };
}

function fullYear(value: number, digits: number): number {
  if (digits > 2) {
    return value;
  }
  return value <= 29 ? 2000 + value : 1900 + value;
}

function parseTimePart(text: string): { fraction: number; pattern: TimePattern } | null {
  const clock = CLOCK.exec(text);
  if (clock === null) {
    return null;
  }
  let hours = Number(clock[1]);
  const minutes = Number(clock[2]);
  const seconds = clock[3] === undefined ? 0 : Number(clock[3]);
  const half = clock[4]?.toUpperCase();
  if (half !== undefined) {
    if (hours < 1 || hours > 12) {
      return null;
    }
    hours = half === 'A' ? hours % 12 : (hours % 12) + 12;
  }
  if (hours > 23 || minutes > 59 || seconds > 59) {
    return null;
  }
  return {
    fraction: serialOfTime(hours, minutes, seconds),
    pattern: clock[3] === undefined ? 'hm' : 'hms'
  };
}
