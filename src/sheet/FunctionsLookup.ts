import { criterionOf } from './Criteria';
import { arity, integerAt, rangeAt, scalar, type Argument, type SheetFunction } from './FunctionKit';
import { compareValues, isError, NA, VALUE, type CellValue } from './Values';

/**
 * Finding a value by looking something else up.
 *
 * `OFFSET` and `INDIRECT` are *not* here. They compute a reference
 * rather than read one, so the dependency graph cannot know their
 * edges before they run — which makes them the evaluator's problem
 * and the most dangerous thing in this phase. See `Evaluator.call`
 * and `Sheet.evaluateCell`.
 *
 * ## Not finding something is `#N/A`
 *
 * This is the reason the sixth error code exists. A `VLOOKUP` that
 * found nothing has not failed: the formula is fine, the table simply
 * does not have that row. Calling it `#VALUE!` would say the formula
 * is wrong and send somebody to debug a formula that is correct.
 *
 * ## The default is the dangerous one
 *
 * `VLOOKUP`'s fourth argument defaults to **approximate**, which is
 * Excel's single worst default: on unsorted data it returns a wrong
 * answer rather than an error, quietly, and the sheet looks fine.
 * Copying it is nevertheless right, because every `VLOOKUP` anybody
 * pastes in was written against that default and changing it would
 * make those formulas silently return something else.
 */

/**
 * The value being looked up, or the error that came in as it.
 *
 * Errors travel through a lookup: `VLOOKUP(#DIV/0!, …)` is
 * `#DIV/0!`, not `#N/A`. The difference matters more here than
 * almost anywhere, because `#N/A` means "this table has no such row"
 * and would send somebody to check the table when the thing that
 * actually broke is three cells upstream.
 *
 * Only the *lookup value* is checked, not the table. A broken cell in
 * a corner of the table somewhere is not this formula's problem, and
 * propagating it would make one bad cell poison every lookup that
 * happened to span it.
 */
function lookupValue(args: readonly Argument[]): CellValue {
  return scalar(args, 0);
}

/** A range's cell at a row and column inside it, both from zero. */
function cellIn(range: Extract<Argument, { kind: 'range' }>, row: number, column: number): CellValue {
  return range.values[row * range.columns + column] ?? null;
}

/**
 * The position of a value in a list, exactly or by the largest that
 * does not exceed it.
 *
 * The approximate search is a binary search and assumes the list is
 * sorted, which is the assumption the caller made by asking for it.
 * Unsorted input gives a wrong answer rather than an error — that is
 * what the mode means, and it is why the exact mode is the one to
 * reach for.
 */
function positionOf(values: readonly CellValue[], wanted: CellValue, exact: boolean): number {
  if (exact) {
    const test = criterionOf(wanted);
    for (let at = 0; at < values.length; at++) {
      if (test.matches(values[at])) {
        return at;
      }
    }
    return -1;
  }

  let low = 0;
  let high = values.length - 1;
  let best = -1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const order = compareValues(values[middle], wanted);
    if (isError(order)) {
      // A blank or a mismatched kind in a sorted column: step past it
      // rather than give up on the whole search.
      low = middle + 1;
      continue;
    }
    if (order <= 0) {
      best = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return best;
}

export const LOOKUP_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** `VLOOKUP(value, table, column, [approximate])`, column from one. */
  VLOOKUP(args) {
    const wrong = arity(args, 3, 4);
    if (wrong !== null) {
      return wrong;
    }
    const wanted = lookupValue(args);
    if (isError(wanted)) {
      return wanted;
    }
    const table = rangeAt(args, 1);
    if (isError(table)) {
      return table;
    }
    const column = integerAt(args, 2);
    if (isError(column)) {
      return column;
    }
    if (column < 1 || column > table.columns) {
      return VALUE;
    }
    const approximate = args.length < 4 || truthy(scalar(args, 3));
    const first: CellValue[] = [];
    for (let row = 0; row < table.rows; row++) {
      first.push(cellIn(table, row, 0));
    }
    const at = positionOf(first, wanted, !approximate);
    return at === -1 ? NA : cellIn(table, at, column - 1);
  },

  /** The same, turned ninety degrees. */
  HLOOKUP(args) {
    const wrong = arity(args, 3, 4);
    if (wrong !== null) {
      return wrong;
    }
    const wanted = lookupValue(args);
    if (isError(wanted)) {
      return wanted;
    }
    const table = rangeAt(args, 1);
    if (isError(table)) {
      return table;
    }
    const row = integerAt(args, 2);
    if (isError(row)) {
      return row;
    }
    if (row < 1 || row > table.rows) {
      return VALUE;
    }
    const approximate = args.length < 4 || truthy(scalar(args, 3));
    const first: CellValue[] = [];
    for (let column = 0; column < table.columns; column++) {
      first.push(cellIn(table, 0, column));
    }
    const at = positionOf(first, wanted, !approximate);
    return at === -1 ? NA : cellIn(table, row - 1, at);
  },

  /**
   * The cell at a row and column of a range, both counting from one.
   *
   * A zero means "all of it", which in a sheet that does not spill
   * can only be honoured when the range is one row or one column
   * wide — `INDEX(A1:A9, 0)` is the whole column and there is nowhere
   * to put it, so it is the first cell. A one-dimensional range takes
   * a single index, which is the form `MATCH` feeds.
   */
  INDEX(args) {
    const wrong = arity(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const range = rangeAt(args, 0);
    if (isError(range)) {
      return range;
    }
    const first = integerAt(args, 1);
    if (isError(first)) {
      return first;
    }
    if (args.length === 2) {
      // One index into a single row or column, counted along it.
      if (range.rows !== 1 && range.columns !== 1) {
        return VALUE;
      }
      if (first < 1 || first > range.values.length) {
        return VALUE;
      }
      return range.values[first - 1] ?? null;
    }
    const second = integerAt(args, 2);
    if (isError(second)) {
      return second;
    }
    const row = first === 0 ? 1 : first;
    const column = second === 0 ? 1 : second;
    if (row < 1 || row > range.rows || column < 1 || column > range.columns) {
      return VALUE;
    }
    return cellIn(range, row - 1, column - 1);
  },

  /**
   * Where a value is in a list, counting from one.
   *
   * The third argument is the odd one in the whole library: **1 means
   * approximate ascending, 0 means exact, −1 means approximate
   * descending**, and it defaults to 1. Nobody remembers that zero is
   * the exact one, which is why `MATCH(x, range, 0)` is written out
   * in full in every sheet that works.
   */
  MATCH(args) {
    const wrong = arity(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const wanted = lookupValue(args);
    if (isError(wanted)) {
      return wanted;
    }
    const range = rangeAt(args, 1);
    if (isError(range)) {
      return range;
    }
    const mode = args.length > 2 ? integerAt(args, 2) : 1;
    if (isError(mode)) {
      return mode;
    }
    if (mode === 0) {
      const at = positionOf(range.values, wanted, true);
      return at === -1 ? NA : at + 1;
    }
    if (mode < 0) {
      // Descending: the smallest value that is at least the target.
      for (let at = range.values.length - 1; at >= 0; at--) {
        const order = compareValues(range.values[at], wanted);
        if (!isError(order) && order >= 0) {
          return at + 1;
        }
      }
      return NA;
    }
    const at = positionOf(range.values, wanted, false);
    return at === -1 ? NA : at + 1;
  },

  /**
   * `XLOOKUP(value, searched, returned, [ifMissing], [mode])`.
   *
   * The modern one, and better in the two ways that matter: the
   * default is **exact**, and there is somewhere to put the answer for
   * when it is not found, so `IFNA` wrapped around a `VLOOKUP` stops
   * being the idiom. Mode −1 and 1 fall back to the nearest smaller
   * or larger value.
   */
  XLOOKUP(args) {
    const wrong = arity(args, 3, 5);
    if (wrong !== null) {
      return wrong;
    }
    const wanted = lookupValue(args);
    if (isError(wanted)) {
      return wanted;
    }
    const searched = rangeAt(args, 1);
    if (isError(searched)) {
      return searched;
    }
    const returned = rangeAt(args, 2);
    if (isError(returned)) {
      return returned;
    }
    if (searched.values.length !== returned.values.length) {
      return VALUE;
    }
    const mode = args.length > 4 ? integerAt(args, 4) : 0;
    if (isError(mode)) {
      return mode;
    }
    let at = positionOf(searched.values, wanted, true);
    if (at === -1 && mode !== 0) {
      at = nearest(searched.values, wanted, mode < 0);
    }
    if (at === -1) {
      return args.length > 3 ? scalar(args, 3) : NA;
    }
    return returned.values[at] ?? null;
  },

  /** The n'th of its arguments, counting from one. */
  CHOOSE(args) {
    const wrong = arity(args, 2, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const which = integerAt(args, 0);
    if (isError(which)) {
      return which;
    }
    return which < 1 || which >= args.length ? VALUE : scalar(args, which);
  },

  /** How many rows and columns a range covers. */
  ROWS(args) {
    const wrong = arity(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const range = rangeAt(args, 0);
    return isError(range) ? range : range.rows;
  },

  COLUMNS(args) {
    const wrong = arity(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const range = rangeAt(args, 0);
    return isError(range) ? range : range.columns;
  }
};

/** The closest value below or above, for `XLOOKUP`'s fallback modes. */
function nearest(values: readonly CellValue[], wanted: CellValue, below: boolean): number {
  let best = -1;
  let bestOrder: number | null = null;
  for (let at = 0; at < values.length; at++) {
    const order = compareValues(values[at], wanted);
    if (isError(order)) {
      continue;
    }
    if (below ? order > 0 : order < 0) {
      continue;
    }
    if (bestOrder === null || (below ? order > bestOrder : order < bestOrder)) {
      bestOrder = order;
      best = at;
    }
  }
  return best;
}

/**
 * `VLOOKUP`'s fourth argument, which is a boolean written as anything.
 *
 * Omitted means approximate; `FALSE` and `0` mean exact. A blank cell
 * passed in means exact, which is Excel and is the safer reading of
 * an argument somebody left empty on purpose.
 */
function truthy(value: CellValue): boolean {
  if (value === null) {
    return false;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  return typeof value === 'string' ? value.toUpperCase() !== 'FALSE' : false;
}
