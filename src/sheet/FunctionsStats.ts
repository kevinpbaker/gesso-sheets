import { checked, integerAt, numberAt, numbersOf, scalar, valuesOf, type SheetFunction } from './FunctionKit';
import { DIV0, isError, NA, VALUE, type CellError } from './Values';

/**
 * Statistics.
 *
 * One decision governs half the file: **the sample, not the
 * population**. `STDEV` and `VAR` divide by n−1, because that is what
 * the unsuffixed names mean in every spreadsheet, and somebody who
 * wants the population form reaches for `STDEVP`. Getting this
 * backwards gives answers that are close enough to look right and
 * wrong in exactly the cases anyone is checking.
 *
 * The other one: **an empty set is an error, not zero**. A `MEDIAN`
 * of nothing has no answer, and zero would be one.
 */

/** Sorted ascending, which four of these need before they start. */
function sortedNumbers(args: Parameters<SheetFunction>[0]): number[] | CellError {
  const numbers = numbersOf(args);
  if (isError(numbers)) {
    return numbers;
  }
  return [...numbers].sort((a, b) => a - b);
}

function sumOf(numbers: readonly number[]): number {
  let total = 0;
  for (const value of numbers) {
    total += value;
  }
  return total;
}

/** The sum of squared deviations, which both spread functions need. */
function sumSquaredDeviations(numbers: readonly number[]): number {
  const mean = sumOf(numbers) / numbers.length;
  let total = 0;
  for (const value of numbers) {
    total += (value - mean) ** 2;
  }
  return total;
}

export const STATS_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  MEDIAN(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const sorted = sortedNumbers(args);
    if (isError(sorted)) {
      return sorted;
    }
    if (sorted.length === 0) {
      return VALUE;
    }
    const middle = Math.floor(sorted.length / 2);
    // An even count has two middles and the median is between them.
    return sorted.length % 2 === 1 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  },

  /**
   * The value that appears most often.
   *
   * `#N/A` when every value appears once, which is Excel's answer and
   * the honest one: there is no mode, and returning the first number
   * would invent one. Ties go to whichever reached its count first,
   * which is also Excel.
   */
  MODE(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    const counts = new Map<number, number>();
    let best: number | null = null;
    let bestCount = 1;
    for (const value of numbers) {
      const count = (counts.get(value) ?? 0) + 1;
      counts.set(value, count);
      if (count > bestCount) {
        bestCount = count;
        best = value;
      }
    }
    return best === null ? NA : best;
  },

  /** The sample standard deviation: n−1, as the bare name means. */
  STDEV(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    // One value has no spread to speak of, and n−1 would be a
    // division by zero — which is exactly what Excel reports.
    return numbers.length < 2 ? DIV0 : Math.sqrt(sumSquaredDeviations(numbers) / (numbers.length - 1));
  },

  VAR(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    return numbers.length < 2 ? DIV0 : sumSquaredDeviations(numbers) / (numbers.length - 1);
  },

  /** The population forms, for when the data really is everything. */
  STDEVP(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    return numbers.length === 0 ? DIV0 : Math.sqrt(sumSquaredDeviations(numbers) / numbers.length);
  },

  VARP(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    return numbers.length === 0 ? DIV0 : sumSquaredDeviations(numbers) / numbers.length;
  },

  /**
   * How many cells hold anything at all.
   *
   * Errors count. A broken cell is a cell somebody filled in, and a
   * `COUNTA` that skipped it would disagree with what the screen
   * shows. Blanks do not count, and neither `COUNTA` nor `COUNT`
   * propagates an error — counting is the one thing still possible
   * over a range with a broken cell in it.
   */
  COUNTA(args) {
    return valuesOf(args).filter(value => value !== null).length;
  },

  /** And how many hold nothing, which is not the same as `''`. */
  COUNTBLANK(args) {
    return valuesOf(args).filter(value => value === null).length;
  },

  /** The k'th largest, counting from one. */
  LARGE(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    return nth(args, false);
  },

  SMALL(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    return nth(args, true);
  },

  /**
   * Where a value comes in the list, counting from one.
   *
   * Descending by default — rank 1 is the largest — which is the
   * convention every spreadsheet uses and the opposite of what the
   * word suggests to a programmer. A non-zero third argument counts
   * the other way.
   *
   * Ties share the better rank and the ranks after them skip, so two
   * firsts are followed by a third. That is what a league table does
   * and what Excel does.
   */
  RANK(args) {
    const wrong = checked(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const pool = numbersOf([args[1]]);
    if (isError(pool)) {
      return pool;
    }
    const ascending = args.length > 2 ? numberAt(args, 2) : 0;
    if (isError(ascending)) {
      return ascending;
    }
    if (!pool.includes(value)) {
      return NA;
    }
    const better = pool.filter(other => (ascending !== 0 ? other < value : other > value)).length;
    return better + 1;
  },

  /**
   * The value a given fraction of the way through the data.
   *
   * Interpolated between the two neighbouring values, which is
   * Excel's `PERCENTILE.INC` and what the bare name has always meant.
   */
  PERCENTILE(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const sorted = sortedNumbers([args[0]]);
    if (isError(sorted)) {
      return sorted;
    }
    const fraction = numberAt(args, 1);
    if (isError(fraction)) {
      return fraction;
    }
    if (sorted.length === 0 || fraction < 0 || fraction > 1) {
      return VALUE;
    }
    const at = fraction * (sorted.length - 1);
    const below = Math.floor(at);
    const above = Math.ceil(at);
    return below === above ? sorted[below] : sorted[below] + (at - below) * (sorted[above] - sorted[below]);
  },

  /** The halfway point of the halves, which is what a quartile is. */
  QUARTILE(args) {
    const wrong = checked(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const quarter = integerAt(args, 1);
    if (isError(quarter)) {
      return quarter;
    }
    if (quarter < 0 || quarter > 4) {
      return VALUE;
    }
    return STATS_FUNCTIONS.PERCENTILE([args[0], { kind: 'value', value: quarter / 4 }], NO_CONTEXT);
  }
};

/** `LARGE` and `SMALL`, which are one function looked at from two ends. */
function nth(args: Parameters<SheetFunction>[0], fromTheSmall: boolean) {
  const sorted = sortedNumbers([args[0]]);
  if (isError(sorted)) {
    return sorted;
  }
  const k = integerAt(args, 1);
  if (isError(k)) {
    return k;
  }
  if (k < 1 || k > sorted.length) {
    return NA;
  }
  return fromTheSmall ? sorted[k - 1] : sorted[sorted.length - k];
}

/**
 * `QUARTILE` calls `PERCENTILE`, which takes a context it never uses.
 *
 * Rather than thread one through, this is the context for a call that
 * cannot reach the clock or the dice — and if either were ever
 * reached through here it would throw rather than quietly answer with
 * something made up.
 */
const NO_CONTEXT = {
  now(): number {
    throw new Error('this function cannot ask the time');
  },
  random(): number {
    throw new Error('this function cannot ask for a random number');
  }
};
