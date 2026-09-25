import { dateOfSerial, serialOfDate, serialOfTime, timeOfSerial, weekdayOf } from './Dates';
import { arity, checked, integerAt, numberAt, numbersOf, textAt, type SheetFunction } from './FunctionKit';
import { isError, VALUE } from './Values';

/**
 * Dates, as arithmetic on the serials `Dates.ts` defines.
 *
 * Nothing here has a date type, because there isn't one: every
 * function takes and returns a number, and it is the *format* on the
 * cell that makes one of those numbers read as a day. That is what
 * lets `EDATE(A1, 3)` be added to, compared and summed like any other
 * number, and it is why this file is thirty lines of calendar and no
 * abstraction.
 *
 * `TODAY` and `NOW` are volatile — see `Functions.VOLATILE` — and take
 * the time from the context rather than from the clock directly, so a
 * spec can say what day it is.
 */

export const DATE_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** Today's date, with no time on it. */
  TODAY(args, ctx) {
    return arity(args, 0, 0) ?? Math.floor(ctx.now());
  },

  /** And the moment, with one. */
  NOW(args, ctx) {
    return arity(args, 0, 0) ?? ctx.now();
  },

  /**
   * A date from its three parts.
   *
   * Out-of-range parts roll over — `DATE(2026, 13, 1)` is January
   * 2027 — which is the mechanism `EDATE` and `EOMONTH` are built
   * out of rather than a leniency. Typing `2026-13-01` into a cell is
   * still refused; see `Dates.parseTypedDate`.
   */
  DATE(args) {
    const wrong = checked(args, 3, 3);
    if (wrong !== null) {
      return wrong;
    }
    const year = integerAt(args, 0);
    if (isError(year)) {
      return year;
    }
    const month = integerAt(args, 1);
    if (isError(month)) {
      return month;
    }
    const day = integerAt(args, 2);
    if (isError(day)) {
      return day;
    }
    const serial = serialOfDate(year, month, day);
    return Number.isFinite(serial) && serial >= 0 ? serial : VALUE;
  },

  TIME(args) {
    const wrong = checked(args, 3, 3);
    if (wrong !== null) {
      return wrong;
    }
    const hours = integerAt(args, 0);
    if (isError(hours)) {
      return hours;
    }
    const minutes = integerAt(args, 1);
    if (isError(minutes)) {
      return minutes;
    }
    const seconds = integerAt(args, 2);
    if (isError(seconds)) {
      return seconds;
    }
    const fraction = serialOfTime(hours, minutes, seconds);
    // Wrapped into a single day, as Excel's `TIME` does: 25:00 is
    // 01:00, because a time of day is what the function returns.
    return ((fraction % 1) + 1) % 1;
  },

  YEAR: part(serial => dateOfSerial(serial).year),
  MONTH: part(serial => dateOfSerial(serial).month),
  DAY: part(serial => dateOfSerial(serial).day),
  HOUR: part(serial => timeOfSerial(serial).hours),
  MINUTE: part(serial => timeOfSerial(serial).minutes),
  SECOND: part(serial => timeOfSerial(serial).seconds),

  /** Sunday is 1 unless a second argument says otherwise. */
  WEEKDAY(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const serial = numberAt(args, 0);
    if (isError(serial)) {
      return serial;
    }
    if (serial < 0) {
      return VALUE;
    }
    const type = args.length > 1 ? integerAt(args, 1) : 1;
    if (isError(type)) {
      return type;
    }
    const sundayOne = weekdayOf(serial);
    switch (type) {
      case 1:
        return sundayOne;
      // Monday is 1, which is how the rest of the world numbers days.
      case 2:
        return ((sundayOne + 5) % 7) + 1;
      // The same week, counted from zero.
      case 3:
        return (sundayOne + 5) % 7;
      default:
        return VALUE;
    }
  },

  /** The same day a number of months away, clamped to the month's end. */
  EDATE(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const serial = numberAt(args, 0);
    if (isError(serial)) {
      return serial;
    }
    const months = integerAt(args, 1);
    if (isError(months)) {
      return months;
    }
    if (serial < 0) {
      return VALUE;
    }
    const { year, month, day } = dateOfSerial(serial);
    // The 31st of January plus one month is the 28th of February, not
    // the 3rd of March: the day is clamped to what the target month
    // has, which is what `EDATE` means and what rolling over would
    // get wrong.
    const lastDay = dateOfSerial(serialOfDate(year, month + months + 1, 0)).day;
    return serialOfDate(year, month + months, Math.min(day, lastDay));
  },

  /**
   * The last day of the month a number of months away.
   *
   * Day zero of the month after, which is the trick the whole of this
   * file's calendar arithmetic rests on.
   */
  EOMONTH(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const serial = numberAt(args, 0);
    if (isError(serial)) {
      return serial;
    }
    const months = integerAt(args, 1);
    if (isError(months)) {
      return months;
    }
    if (serial < 0) {
      return VALUE;
    }
    const { year, month } = dateOfSerial(serial);
    return serialOfDate(year, month + months + 1, 0);
  },

  /**
   * Whole working days between two dates, both ends included.
   *
   * Saturdays and Sundays are the weekend and there is no holiday
   * list, which is the honest version of the function: a holiday
   * argument that silently ignored the holidays would be worse than
   * not having one.
   */
  NETWORKDAYS(args) {
    const wrong = checked(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const from = numberAt(args, 0);
    if (isError(from)) {
      return from;
    }
    const to = numberAt(args, 1);
    if (isError(to)) {
      return to;
    }
    if (from < 0 || to < 0) {
      return VALUE;
    }
    const holidays = args.length > 2 ? numbersOf([args[2]]) : [];
    if (isError(holidays)) {
      return holidays;
    }
    const skipped = new Set(holidays.map(Math.floor));
    const first = Math.floor(Math.min(from, to));
    const last = Math.floor(Math.max(from, to));
    let days = 0;
    for (let at = first; at <= last; at++) {
      const weekday = weekdayOf(at);
      if (weekday === 1 || weekday === 7 || skipped.has(at)) {
        continue;
      }
      days++;
    }
    // Counted backwards when the dates were given that way round,
    // which is what Excel does and what makes the sign meaningful.
    return from > to ? -days : days;
  },

  /**
   * The distance between two dates in whole units.
   *
   * The unit is text — `"Y"`, `"M"`, `"D"` — which is a wart and is
   * the function's actual interface everywhere it exists. The
   * remainder units (`"YM"`, `"MD"`, `"YD"`) are the ones people use
   * to write "3 years and 2 months" and are why it survives at all.
   */
  DATEDIF(args) {
    const wrong = checked(args, 3, 3);
    if (wrong !== null) {
      return wrong;
    }
    const from = numberAt(args, 0);
    if (isError(from)) {
      return from;
    }
    const to = numberAt(args, 1);
    if (isError(to)) {
      return to;
    }
    const unit = textAt(args, 2);
    if (isError(unit)) {
      return unit;
    }
    // Backwards is an error rather than a negative number, which is
    // Excel's behaviour and catches the commonest misuse: the two
    // arguments in the wrong order.
    if (from > to || from < 0) {
      return VALUE;
    }
    const start = dateOfSerial(from);
    const end = dateOfSerial(to);
    const wholeMonths =
      (end.year - start.year) * 12 + (end.month - start.month) - (end.day < start.day ? 1 : 0);

    switch (unit.toUpperCase()) {
      case 'D':
        return Math.floor(to) - Math.floor(from);
      case 'M':
        return wholeMonths;
      case 'Y':
        return Math.floor(wholeMonths / 12);
      // Months, ignoring the years.
      case 'YM':
        return wholeMonths % 12;
      // Days, ignoring the months.
      case 'MD': {
        const lastDay = dateOfSerial(serialOfDate(end.year, end.month, 0)).day;
        return end.day >= start.day ? end.day - start.day : end.day + lastDay - start.day;
      }
      // Days, ignoring the years: the distance from the most recent
      // anniversary of the start date, which is in the end's year
      // unless that date has not come round yet.
      case 'YD': {
        const notYet = start.month > end.month || (start.month === end.month && start.day > end.day);
        return Math.floor(to) - serialOfDate(end.year - (notYet ? 1 : 0), start.month, start.day);
      }
      default:
        return VALUE;
    }
  }
};

/**
 * The six that pull one field out of a serial.
 *
 * Written once because they differ only in which field, and a
 * negative serial is `#VALUE!` in all of them: there are no dates
 * before the epoch to have a year.
 */
function part(field: (serial: number) => number): SheetFunction {
  return args => {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const serial = numberAt(args, 0);
    if (isError(serial)) {
      return serial;
    }
    return serial < 0 ? VALUE : field(serial);
  };
}
