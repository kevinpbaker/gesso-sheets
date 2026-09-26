import { arity, checked, checkedScalars, numberAt, numbersOf, rangeAt, type SheetFunction } from './FunctionKit';
import { DIV0, isError, NUM, VALUE, type CellError } from './Values';

/**
 * Arithmetic.
 *
 * Two conventions run through the file and neither is JavaScript's.
 *
 * **Rounding goes away from zero on a tie**, so `-2.5` rounds to `-3`.
 * `Math.round` gives `-2`, and a column of figures that rounded
 * negatives towards zero would not add up the way the person checking
 * it expects.
 *
 * **A result that is not a real number is `#NUM!`**, not `NaN`, as it
 * is in Excel: `SQRT(-1)` and `LN(0)` were given numbers of the right
 * kind that are out of range. What must never happen is a `NaN`
 * reaching a cell: it compares false with itself and poisons
 * everything downstream in silence.
 */

/** Away from zero on a tie, as a spreadsheet rounds. */
export function roundHalfAway(value: number, places: number): number {
  const factor = 10 ** Math.trunc(places);
  // To fifteen significant digits first, which is as far as Excel
  // believes a number: 0.05 is stored as 0.04999…, and scaled by ten it
  // is 0.4999…, which rounds to nothing — where every spreadsheet says
  // ROUND(0.05, 1) is 0.1, because nobody typed 0.04999….
  const scaled = Number((value * factor).toPrecision(15));
  const rounded = scaled < 0 ? -Math.round(-scaled) : Math.round(scaled);
  return rounded / factor;
}

/** A result, or `#NUM!` when the arithmetic left the real numbers. */
function real(value: number) {
  return Number.isFinite(value) ? value : NUM;
}

export const MATH_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  SQRT(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value < 0 ? NUM : Math.sqrt(value);
  },

  POWER(args) {
    const wrong = checkedScalars(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const base = numberAt(args, 0);
    if (isError(base)) {
      return base;
    }
    const exponent = numberAt(args, 1);
    return isError(exponent) ? exponent : real(base ** exponent);
  },

  /**
   * The remainder, with the sign of the *divisor*.
   *
   * `MOD(-3, 2)` is 1 in every spreadsheet and -1 in JavaScript, and
   * the spreadsheet is the one worth agreeing with: the whole use of
   * `MOD` is cycling through a repeating pattern, which the negative
   * answer breaks at zero.
   */
  MOD(args) {
    const wrong = checkedScalars(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const divisor = numberAt(args, 1);
    if (isError(divisor)) {
      return divisor;
    }
    return divisor === 0 ? DIV0 : value - divisor * Math.floor(value / divisor);
  },

  /** Towards negative infinity, which is what makes `INT(-2.5)` -3. */
  INT(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.floor(value);
  },

  /** Towards zero, which is the other one, and why both exist. */
  TRUNC(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const places = args.length > 1 ? numberAt(args, 1) : 0;
    if (isError(places)) {
      return places;
    }
    const factor = 10 ** Math.trunc(places);
    return Math.trunc(value * factor) / factor;
  },

  /** Up to the next multiple, away from zero. */
  CEILING(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return toMultiple(args, Math.ceil, 0);
  },

  /** Down to the multiple below, towards zero. */
  FLOOR(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return toMultiple(args, Math.floor, DIV0);
  },

  SIGN(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.sign(value);
  },

  EXP(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : real(Math.exp(value));
  },

  LN(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value <= 0 ? NUM : Math.log(value);
  },

  /** Base ten unless a base is given, as every spreadsheet has it. */
  LOG(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const base = args.length > 1 ? numberAt(args, 1) : 10;
    if (isError(base)) {
      return base;
    }
    if (value <= 0 || base <= 0 || base === 1) {
      return NUM;
    }
    return Math.log(value) / Math.log(base);
  },

  LOG10(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value <= 0 ? NUM : Math.log10(value);
  },

  /**
   * A number in [0, 1).
   *
   * Volatile: see `Functions.VOLATILE`. It takes its randomness from
   * the context rather than from `Math.random` so that a spec can
   * pin it, which is the only way to assert anything about it.
   */
  RAND(args, ctx) {
    return arity(args, 0, 0) ?? ctx.random();
  },

  /** A whole number in a closed range, both ends included. */
  RANDBETWEEN(args, ctx) {
    const wrong = checkedScalars(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const low = numberAt(args, 0);
    if (isError(low)) {
      return low;
    }
    const high = numberAt(args, 1);
    if (isError(high)) {
      return high;
    }
    const first = Math.ceil(low);
    const last = Math.floor(high);
    return first > last ? NUM : first + Math.floor(ctx.random() * (last - first + 1));
  },

  /**
   * The ranges multiplied cell by cell, then added up.
   *
   * Anything that is not a number counts as zero rather than as an
   * error, which is Excel's rule and the one that makes
   * `SUMPRODUCT(A2:A9, B2:B9)` survive a blank row in the middle of
   * the table. Ranges of different sizes are `#VALUE!`: there is no
   * sensible pairing, and guessing one would quietly total the wrong
   * cells.
   */
  SUMPRODUCT(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    // A single value is an array of one, which is how Excel reads
    // `SUMPRODUCT(3, B10, C9)`: three times B10 times C9.
    const ranges = args.map(arg =>
      arg.kind === 'range' ? arg : { kind: 'range' as const, values: [arg.value], rows: 1, columns: 1 }
    );
    const shape = ranges[0];
    if (ranges.some(range => range.rows !== shape.rows || range.columns !== shape.columns)) {
      return VALUE;
    }
    let total = 0;
    for (let at = 0; at < shape.values.length; at++) {
      let product = 1;
      for (const range of ranges) {
        const value = range.values[at];
        // An error anywhere in the arrays is the answer, as it is in
        // Excel; text and blanks count as nothing.
        if (isError(value)) {
          return value;
        }
        product *= typeof value === 'number' ? value : 0;
      }
      total += product;
    }
    return total;
  },

  /** The whole numbers up to one, multiplied together. */
  FACT(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const whole = Math.trunc(value);
    if (whole < 0 || whole > 170) {
      // Past 170 the answer is larger than a double can hold, and
      // `Infinity` in a cell is worse than saying so.
      return NUM;
    }
    let total = 1;
    for (let at = 2; at <= whole; at++) {
      total *= at;
    }
    return total;
  },

  /** Every number in the arguments, multiplied. */
  PRODUCT(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    // Nothing to multiply is zero, not the empty product: Excel's
    // answer, and the one that does not turn an empty column into a 1
    // that looks like data.
    return numbers.length === 0 ? 0 : numbers.reduce((total, value) => total * value, 1);
  },

  PI(args) {
    return arity(args, 0, 0) ?? Math.PI;
  },

  ROUNDUP(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return awayFromZero(args, Math.ceil);
  },

  ROUNDDOWN(args) {
    const wrong = checkedScalars(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return awayFromZero(args, Math.floor);
  },

  /** Whether a number, truncated, is even; and odd. Text is `#VALUE!`. */
  ISEVEN(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.trunc(value) % 2 === 0;
  },

  ISODD(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.abs(Math.trunc(value)) % 2 === 1;
  },

  /**
   * A number rounded to the nearest multiple, a tie away from zero.
   * The number and the multiple must share a sign, as Excel insists,
   * and a pair that does not is `#NUM!`.
   */
  MROUND(args) {
    const wrong = checkedScalars(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const multiple = numberAt(args, 1);
    if (isError(multiple)) {
      return multiple;
    }
    if (multiple === 0) {
      return 0;
    }
    if (value !== 0 && Math.sign(value) !== Math.sign(multiple)) {
      return NUM;
    }
    // The same sign, so the quotient is positive and rounding it up
    // on a tie is rounding away from zero. Nudged first, so that
    // 10 / 0.1 — which is 99.99999… in binary — rounds to the hundred
    // it means.
    return Math.round(value / multiple + 1e-9) * multiple;
  },

  /** The whole part of a division, as `INT` would give it for positives and `TRUNC` for all. */
  QUOTIENT(args) {
    const wrong = checkedScalars(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const numerator = numberAt(args, 0);
    if (isError(numerator)) {
      return numerator;
    }
    const denominator = numberAt(args, 1);
    if (isError(denominator)) {
      return denominator;
    }
    return denominator === 0 ? DIV0 : Math.trunc(numerator / denominator);
  },

  /** Rounded away from zero to the next even integer; and odd. */
  EVEN(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    return (value < 0 ? -1 : 1) * Math.ceil(Math.abs(value) / 2) * 2;
  },

  ODD(args) {
    const wrong = checkedScalars(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const up = Math.ceil(Math.abs(value));
    return (value < 0 ? -1 : 1) * (up % 2 === 1 ? up : up + 1);
  }
};

/** `CEILING` and `FLOOR`, which differ only in which way they go. */
/**
 * `onZero` is what a step of zero gives, which is where the two differ
 * in Excel: `CEILING(x, 0)` is 0 and `FLOOR(x, 0)` divides by it.
 */
function toMultiple(args: Parameters<SheetFunction>[0], round: (value: number) => number, onZero: number | CellError) {
  const value = numberAt(args, 0);
  if (isError(value)) {
    return value;
  }
  const step = args.length > 1 ? numberAt(args, 1) : 1;
  if (isError(step)) {
    return step;
  }
  if (step === 0) {
    return value === 0 ? 0 : onZero;
  }
  // Worked on the magnitude and signed back, so both functions mean
  // the same thing either side of zero.
  const sign = value < 0 ? -1 : 1;
  return sign * round(Math.abs(value) / Math.abs(step)) * Math.abs(step);
}

/** `ROUNDUP` and `ROUNDDOWN`, which are the same trick at a scale. */
function awayFromZero(args: Parameters<SheetFunction>[0], round: (value: number) => number) {
  const value = numberAt(args, 0);
  if (isError(value)) {
    return value;
  }
  const places = args.length > 1 ? numberAt(args, 1) : 0;
  if (isError(places)) {
    return places;
  }
  const factor = 10 ** Math.trunc(places);
  const sign = value < 0 ? -1 : 1;
  return (sign * round(Math.abs(value) * factor)) / factor;
}
