import { criterionOf } from './Criteria';
import { arity, rangeAt, scalar, type Argument, type SheetFunction } from './FunctionKit';
import { DIV0, isError, VALUE, type CellError, type CellValue } from './Values';

/**
 * The aggregates that ask a question first.
 *
 * Six functions and one idea: walk a range, keep the cells a
 * criterion matches, and do something with what is left. The
 * criterion grammar is `Criteria.ts` — shared, because six copies of
 * "does this cell match" is six chances to disagree about whether
 * `"<>"` matches a blank.
 *
 * The argument orders are an inconsistency inherited from Excel and
 * kept on purpose. `SUMIF` puts the range it tests *first* and the
 * range it adds last; `SUMIFS` puts the range it adds first and the
 * pairs after it. They disagree, everybody trips over it once, and a
 * spreadsheet that quietly fixed it would break every formula anyone
 * pasted in from somewhere else.
 */

/** The positions in a range that match, which is all six need. */
function matching(range: Extract<Argument, { kind: 'range' }>, criterion: CellValue): number[] {
  const test = criterionOf(criterion);
  const hits: number[] = [];
  for (let at = 0; at < range.values.length; at++) {
    if (test.matches(range.values[at])) {
      hits.push(at);
    }
  }
  return hits;
}

/** Positions matching every pair of (range, criterion), as `*IFS` needs. */
function matchingAll(args: readonly Argument[], from: number): number[] | CellError {
  let hits: number[] | null = null;
  let shape: { rows: number; columns: number } | null = null;

  for (let at = from; at + 1 < args.length; at += 2) {
    const range = rangeAt(args, at);
    if (isError(range)) {
      return range;
    }
    // Every criteria range has to line up with the others, or the
    // positions they agree on are positions in different tables.
    shape ??= { rows: range.rows, columns: range.columns };
    if (range.rows !== shape.rows || range.columns !== shape.columns) {
      return VALUE;
    }
    const found = new Set(matching(range, scalar(args, at + 1)));
    hits = hits === null ? [...found] : hits.filter(position => found.has(position));
  }
  return hits ?? [];
}

/**
 * The cells a set of positions names, in the range that supplies the
 * values.
 *
 * `SUMIF(A2:A9, ">10", B2:B9)` tests one range and adds another, and
 * the two are matched by *position* rather than by address: Excel
 * takes the sum range's top-left corner and walks the same shape from
 * there, so a sum range written as `B2` alone still works. Matching
 * by position is what makes that true here too.
 */
function pick(values: readonly CellValue[], positions: readonly number[]): number[] {
  const numbers: number[] = [];
  for (const at of positions) {
    const value = values[at];
    if (typeof value === 'number') {
      numbers.push(value);
    }
  }
  return numbers;
}

function total(numbers: readonly number[]): number {
  let sum = 0;
  for (const value of numbers) {
    sum += value;
  }
  return sum;
}

export const CONDITIONAL_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** `SUMIF(range, criteria, [sum_range])` — tested range first. */
  SUMIF(args) {
    const wrong = arity(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const tested = rangeAt(args, 0);
    if (isError(tested)) {
      return tested;
    }
    const summed = args.length > 2 ? rangeAt(args, 2) : tested;
    if (isError(summed)) {
      return summed;
    }
    return total(pick(summed.values, matching(tested, scalar(args, 1))));
  },

  /** `SUMIFS(sum_range, range1, criteria1, …)` — added range first. */
  SUMIFS(args) {
    const wrong = arity(args, 3, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const summed = rangeAt(args, 0);
    if (isError(summed)) {
      return summed;
    }
    const hits = matchingAll(args, 1);
    if (isError(hits)) {
      return hits;
    }
    return total(pick(summed.values, hits));
  },

  COUNTIF(args) {
    const wrong = arity(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const tested = rangeAt(args, 0);
    if (isError(tested)) {
      return tested;
    }
    return matching(tested, scalar(args, 1)).length;
  },

  COUNTIFS(args) {
    const wrong = arity(args, 2, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const hits = matchingAll(args, 0);
    return isError(hits) ? hits : hits.length;
  },

  AVERAGEIF(args) {
    const wrong = arity(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const tested = rangeAt(args, 0);
    if (isError(tested)) {
      return tested;
    }
    const averaged = args.length > 2 ? rangeAt(args, 2) : tested;
    if (isError(averaged)) {
      return averaged;
    }
    const numbers = pick(averaged.values, matching(tested, scalar(args, 1)));
    // Nothing matched is `#DIV/0!`, as an empty `AVERAGE` is: zero
    // would be an answer, and there is no answer.
    return numbers.length === 0 ? DIV0 : total(numbers) / numbers.length;
  },

  AVERAGEIFS(args) {
    const wrong = arity(args, 3, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const averaged = rangeAt(args, 0);
    if (isError(averaged)) {
      return averaged;
    }
    const hits = matchingAll(args, 1);
    if (isError(hits)) {
      return hits;
    }
    const numbers = pick(averaged.values, hits);
    return numbers.length === 0 ? DIV0 : total(numbers) / numbers.length;
  },

  MAXIFS(args) {
    const wrong = arity(args, 3, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    return extremeIf(args, true);
  },

  MINIFS(args) {
    const wrong = arity(args, 3, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    return extremeIf(args, false);
  }
};

function extremeIf(args: readonly Argument[], largest: boolean): CellValue {
  const searched = rangeAt(args, 0);
  if (isError(searched)) {
    return searched;
  }
  const hits = matchingAll(args, 1);
  if (isError(hits)) {
    return hits;
  }
  const numbers = pick(searched.values, hits);
  // Nothing matched is zero, which is what Excel's `MAXIFS` answers —
  // unlike `AVERAGEIFS`, where zero would be a claim about a mean.
  if (numbers.length === 0) {
    return 0;
  }
  return largest ? Math.max(...numbers) : Math.min(...numbers);
}
