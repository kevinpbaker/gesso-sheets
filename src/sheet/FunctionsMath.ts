import { arity, checked, numberAt, numbersOf, rangeAt, type SheetFunction } from './FunctionKit';
import { DIV0, isError, VALUE } from './Values';

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
 * **A result that is not a real number is `#VALUE!`**, not `NaN`.
 * `SQRT(-1)` and `LN(0)` are `#NUM!` in Excel; there is no `#NUM!`
 * here — see `Values.ts` — and `#VALUE!` says the true thing, which is
 * that a value of the wrong kind went in. What must never happen is a
 * `NaN` reaching a cell: it compares false with itself and poisons
 * everything downstream in silence.
 */

/** Away from zero on a tie, as a spreadsheet rounds. */
export function roundHalfAway(value: number, places: number): number {
  const factor = 10 ** Math.trunc(places);
  const scaled = value * factor;
  const rounded = scaled < 0 ? -Math.round(-scaled) : Math.round(scaled);
  return rounded / factor;
}

/** A result, or `#VALUE!` when the arithmetic left the real numbers. */
function real(value: number) {
  return Number.isFinite(value) ? value : VALUE;
}

export const MATH_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  SQRT(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value < 0 ? VALUE : Math.sqrt(value);
  },

  POWER(args) {
    const wrong = checked(args, 2, 2);
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
    const wrong = checked(args, 2, 2);
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
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.floor(value);
  },

  /** Towards zero, which is the other one, and why both exist. */
  TRUNC(args) {
    const wrong = checked(args, 1, 2);
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
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return toMultiple(args, Math.ceil);
  },

  /** Down to the multiple below, towards zero. */
  FLOOR(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return toMultiple(args, Math.floor);
  },

  SIGN(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.sign(value);
  },

  EXP(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : real(Math.exp(value));
  },

  LN(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value <= 0 ? VALUE : Math.log(value);
  },

  /** Base ten unless a base is given, as every spreadsheet has it. */
  LOG(args) {
    const wrong = checked(args, 1, 2);
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
      return VALUE;
    }
    return Math.log(value) / Math.log(base);
  },

  LOG10(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : value <= 0 ? VALUE : Math.log10(value);
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
    const wrong = checked(args, 2, 2);
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
    return first > last ? VALUE : first + Math.floor(ctx.random() * (last - first + 1));
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
    const ranges = [];
    for (let at = 0; at < args.length; at++) {
      const range = rangeAt(args, at);
      if (isError(range)) {
        return range;
      }
      ranges.push(range);
    }
    const shape = ranges[0];
    if (ranges.some(range => range.rows !== shape.rows || range.columns !== shape.columns)) {
      return VALUE;
    }
    let total = 0;
    for (let at = 0; at < shape.values.length; at++) {
      let product = 1;
      for (const range of ranges) {
        const value = range.values[at];
        product *= typeof value === 'number' ? value : 0;
      }
      total += product;
    }
    return total;
  },

  /** The whole numbers up to one, multiplied together. */
  FACT(args) {
    const wrong = checked(args, 1, 1);
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
      return VALUE;
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
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return awayFromZero(args, Math.ceil);
  },

  ROUNDDOWN(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    return awayFromZero(args, Math.floor);
  }
};

/** `CEILING` and `FLOOR`, which differ only in which way they go. */
function toMultiple(args: Parameters<SheetFunction>[0], round: (value: number) => number) {
  const value = numberAt(args, 0);
  if (isError(value)) {
    return value;
  }
  const step = args.length > 1 ? numberAt(args, 1) : 1;
  if (isError(step)) {
    return step;
  }
  if (step === 0) {
    return value === 0 ? 0 : VALUE;
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
