import { formatWith, type NumberFormat } from './Format';
import { checked, integerAt, scalar, textAt, valuesOf, type SheetFunction } from './FunctionKit';
import { isError, toBoolean, toText, VALUE } from './Values';

/**
 * Text.
 *
 * Two things worth knowing before reading any of it.
 *
 * **Positions count from one**, everywhere — `MID(text, 1, 3)` is the
 * first three characters. Every function here converts at its edge
 * and none of them leaks a zero-based index, because a spreadsheet
 * that counted from zero in one function out of sixteen would be
 * wrong in the way that takes an afternoon to find.
 *
 * **A position past the end is not an error.** `LEFT("ab", 9)` is
 * `"ab"`, not `#VALUE!`. Only a *negative* count is an error, and
 * only `FIND` and `SEARCH` report not-finding as one — because there
 * the answer is a position and there isn't one.
 */

/**
 * The patterns `TEXT` understands.
 *
 * Excel takes a format *language* here, and this application
 * deliberately does not have one — see `NumberFormat` in `Format.ts`,
 * where the reasoning is set out. What it has instead is the named
 * formats, so `TEXT` accepts the pattern strings that name them and
 * says `#VALUE!` to anything else.
 *
 * That is a real limit and it is the honest version of one: a
 * half-implemented pattern language that silently ignored the parts
 * it could not do would produce text that is wrong rather than
 * absent, which is worse in a cell nobody is checking.
 */
const PATTERNS: Readonly<Record<string, NumberFormat>> = {
  '0': { kind: 'number', places: 0, thousands: false },
  '0.0': { kind: 'number', places: 1, thousands: false },
  '0.00': { kind: 'number', places: 2, thousands: false },
  '#,##0': { kind: 'number', places: 0, thousands: true },
  '#,##0.00': { kind: 'number', places: 2, thousands: true },
  '0%': { kind: 'percent', places: 0 },
  '0.0%': { kind: 'percent', places: 1 },
  '0.00%': { kind: 'percent', places: 2 },
  '0.00E+00': { kind: 'scientific', places: 2 },
  '$#,##0': { kind: 'currency', places: 0, symbol: '$' },
  '$#,##0.00': { kind: 'currency', places: 2, symbol: '$' },
  'YYYY-MM-DD': { kind: 'date', pattern: 'ymd' },
  'D MMM YYYY': { kind: 'date', pattern: 'dmy' },
  'MMM D, YYYY': { kind: 'date', pattern: 'mdy' },
  'HH:MM': { kind: 'time', pattern: 'hm' },
  'HH:MM:SS': { kind: 'time', pattern: 'hms' },
  'YYYY-MM-DD HH:MM': { kind: 'datetime', date: 'ymd', time: 'hm' },
  GENERAL: { kind: 'general' }
};

export const TEXT_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  LEFT(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const count = args.length > 1 ? integerAt(args, 1) : 1;
    if (isError(count)) {
      return count;
    }
    return count < 0 ? VALUE : text.slice(0, count);
  },

  RIGHT(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const count = args.length > 1 ? integerAt(args, 1) : 1;
    if (isError(count)) {
      return count;
    }
    // `slice(-0)` is the whole string, which is the one case where
    // the obvious spelling is wrong.
    return count < 0 ? VALUE : count === 0 ? '' : text.slice(-count);
  },

  MID(args) {
    const wrong = checked(args, 3, 3);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const start = integerAt(args, 1);
    if (isError(start)) {
      return start;
    }
    const count = integerAt(args, 2);
    if (isError(count)) {
      return count;
    }
    if (start < 1 || count < 0) {
      return VALUE;
    }
    return text.slice(start - 1, start - 1 + count);
  },

  LEN(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    return isError(text) ? text : text.length;
  },

  /** Case-sensitive, and `#VALUE!` when it is not there. */
  FIND(args) {
    const wrong = checked(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    return locate(args, false);
  },

  /** The same question asked case-insensitively. */
  SEARCH(args) {
    const wrong = checked(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    return locate(args, true);
  },

  /**
   * Spaces off both ends, and runs inside collapsed to one.
   *
   * The inner collapse is the part people forget is there, and it is
   * what makes `TRIM` useful on names pasted out of a report rather
   * than merely tidy.
   */
  TRIM(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    return isError(text) ? text : text.trim().replace(/ +/g, ' ');
  },

  UPPER(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    return isError(text) ? text : text.toUpperCase();
  },

  LOWER(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    return isError(text) ? text : text.toLowerCase();
  },

  /** Each word's first letter up and the rest down. */
  PROPER(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    // A "word" starts after anything that is not a letter, which is
    // why `o'brien` becomes `O'Brien` here as it does in Excel.
    return text.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (_, before, letter) => before + letter.toUpperCase());
  },

  /** Every occurrence, or just the n'th when one is named. */
  SUBSTITUTE(args) {
    const wrong = checked(args, 3, 4);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const from = textAt(args, 1);
    if (isError(from)) {
      return from;
    }
    const to = textAt(args, 2);
    if (isError(to)) {
      return to;
    }
    if (from === '') {
      return text;
    }
    if (args.length < 4) {
      return text.split(from).join(to);
    }
    const which = integerAt(args, 3);
    if (isError(which)) {
      return which;
    }
    if (which < 1) {
      return VALUE;
    }
    let at = -1;
    for (let seen = 0; seen < which; seen++) {
      at = text.indexOf(from, at + 1);
      if (at === -1) {
        return text;
      }
    }
    return text.slice(0, at) + to + text.slice(at + from.length);
  },

  /** By position, which is what makes it different from SUBSTITUTE. */
  REPLACE(args) {
    const wrong = checked(args, 4, 4);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const start = integerAt(args, 1);
    if (isError(start)) {
      return start;
    }
    const count = integerAt(args, 2);
    if (isError(count)) {
      return count;
    }
    const replacement = textAt(args, 3);
    if (isError(replacement)) {
      return replacement;
    }
    if (start < 1 || count < 0) {
      return VALUE;
    }
    return text.slice(0, start - 1) + replacement + text.slice(start - 1 + count);
  },

  REPT(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const times = integerAt(args, 1);
    if (isError(times)) {
      return times;
    }
    if (times < 0) {
      return VALUE;
    }
    // A cap, because `REPT("x", 1e9)` is a gigabyte in a cell and the
    // application worker never comes back from it.
    if (text.length * times > 32_767) {
      return VALUE;
    }
    return text.repeat(times);
  },

  /**
   * Joined with a delimiter, optionally skipping the gaps.
   *
   * The skipping is the whole point: a column of addresses with
   * missing second lines joins into `"12 High St, , Leeds"` without
   * it, which is exactly the output nobody wants.
   */
  TEXTJOIN(args) {
    const wrong = checked(args, 3, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const delimiter = textAt(args, 0);
    if (isError(delimiter)) {
      return delimiter;
    }
    const skipEmpty = toBoolean(scalar(args, 1));
    if (isError(skipEmpty)) {
      return skipEmpty;
    }
    const pieces: string[] = [];
    for (const value of valuesOf(args.slice(2))) {
      const piece = toText(value);
      if (isError(piece)) {
        return piece;
      }
      if (skipEmpty && piece === '') {
        continue;
      }
      pieces.push(piece);
    }
    return pieces.join(delimiter);
  },

  /** Text back to a number, and `#VALUE!` when it is not one. */
  VALUE(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = scalar(args, 0);
    if (typeof value === 'number') {
      return value;
    }
    const text = toText(value);
    if (isError(text)) {
      return text;
    }
    const trimmed = text.trim();
    if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(trimmed)) {
      return VALUE;
    }
    return Number(trimmed);
  },

  /**
   * Phase 9's formatter, called from inside a formula.
   *
   * The seam runs the other way from the one dates made: there a
   * typed value chose a format, and here a formula asks the formatter
   * for text. Both are the same rule — the format axis is where
   * values become readable — approached from opposite sides.
   */
  TEXT(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const pattern = textAt(args, 1);
    if (isError(pattern)) {
      return pattern;
    }
    const format = PATTERNS[pattern.toUpperCase()];
    return format === undefined ? VALUE : formatWith(scalar(args, 0), format);
  },

  /** One character from its code, and the code from one character. */
  CHAR(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const code = integerAt(args, 0);
    if (isError(code)) {
      return code;
    }
    return code < 1 || code > 0x10_ffff ? VALUE : String.fromCodePoint(code);
  },

  CODE(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const text = textAt(args, 0);
    if (isError(text)) {
      return text;
    }
    const code = text.codePointAt(0);
    return code === undefined ? VALUE : code;
  },

  /** True when two strings are the same, case and all. */
  EXACT(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const left = textAt(args, 0);
    if (isError(left)) {
      return left;
    }
    const right = textAt(args, 1);
    return isError(right) ? right : left === right;
  },

  /** Joins its arguments, which `&` also does one pair at a time. */
  CONCATENATE(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    let joined = '';
    for (const value of valuesOf(args)) {
      const piece = toText(value);
      if (isError(piece)) {
        return piece;
      }
      joined += piece;
    }
    return joined;
  }
};

/** `FIND` and `SEARCH`, which differ only in whether case matters. */
function locate(args: Parameters<SheetFunction>[0], foldCase: boolean) {
  const needle = textAt(args, 0);
  if (isError(needle)) {
    return needle;
  }
  const haystack = textAt(args, 1);
  if (isError(haystack)) {
    return haystack;
  }
  const start = args.length > 2 ? integerAt(args, 2) : 1;
  if (isError(start)) {
    return start;
  }
  if (start < 1 || start > haystack.length + 1) {
    return VALUE;
  }
  const at = foldCase
    ? haystack.toLowerCase().indexOf(needle.toLowerCase(), start - 1)
    : haystack.indexOf(needle, start - 1);
  // Not found is `#VALUE!` because the answer is a position and there
  // is not one. It is the one place in this file where absence is an
  // error rather than an empty string.
  return at === -1 ? VALUE : at + 1;
}
