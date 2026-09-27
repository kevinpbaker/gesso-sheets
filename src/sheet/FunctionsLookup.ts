import { criterionOf } from './Criteria';
import { arity, arrayOrValue, integerAt, rangeAt, scalar, type Argument, type ArrayValue, type SheetFunction } from './FunctionKit';
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
    const value = values[middle];
    // A blank is stepped past, like a mismatched kind: a blank compares
    // equal to "" and below every word, so taken as a value it looked
    // like the largest thing not above any text wanted.
    const order = value === null ? VALUE : compareValues(value, wanted);
    if (isError(order)) {
      // A blank or a mismatched kind in a sorted column: step past it
      // rather than give up on the whole search.
      low = middle + 1;
      continue;
    }
    if (order === 0) {
      // Found, and Excel stops here — at the last of a run of equal
      // values — rather than searching on to the right, which in a list
      // that is not quite sorted walks past the answer to a worse one.
      let last = middle;
      while (last + 1 < values.length && compareValues(values[last + 1], wanted) === 0) {
        last++;
      }
      return last;
    }
    if (order < 0) {
      best = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return best;
}

/** The last exact match, for a search from the end. */
function lastPositionOf(values: readonly CellValue[], wanted: CellValue): number {
  for (let at = values.length - 1; at >= 0; at--) {
    if (positionOf([values[at]], wanted, true) === 0) {
      return at;
    }
  }
  return -1;
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
    // One value is an array of one, as in Excel: `INDEX(5, 1)` is 5. An
    // array that came out one cell wide and one tall arrives as a value,
    // so without this `INDEX(BYCOL(A1:A3, …), 1)` was `#VALUE!`.
    const given = args[0];
    const range =
      given?.kind === 'value' && !isError(given.value)
        ? { kind: 'range' as const, values: [given.value], rows: 1, columns: 1 }
        : rangeAt(args, 0);
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
    const wrong = arity(args, 3, 6);
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
    // The result may be wider than the searched column (or taller than
    // the searched row), and the answer is the whole row it finds, which
    // spills; see `found`.
    const down = searched.columns === 1;
    const matches = down ? returned.rows === searched.values.length : returned.columns === searched.values.length;
    if (!matches) {
      return VALUE;
    }
    const mode = args.length > 4 ? integerAt(args, 4) : 0;
    if (isError(mode)) {
      return mode;
    }
    // 1 searches first to last and -1 last to first; 2 and -2 are a
    // binary search over data sorted up or down.
    const search = args.length > 5 ? integerAt(args, 5) : 1;
    if (isError(search)) {
      return search;
    }
    if (![1, -1, 2, -2].includes(search)) {
      return VALUE;
    }
    if (search === 2 || search === -2) {
      const at = bisected(searched.values, wanted, search === -2, mode);
      if (at === -1) {
        return args.length > 3 ? scalar(args, 3) : NA;
      }
      return found(returned, at, down);
    }
    let at = search === -1 ? lastPositionOf(searched.values, wanted) : positionOf(searched.values, wanted, true);
    if (at === -1 && mode !== 0) {
      at = nearest(searched.values, wanted, mode < 0);
    }
    if (at === -1) {
      return args.length > 3 ? scalar(args, 3) : NA;
    }
    return found(returned, at, down);
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

/**
 * `XLOOKUP`'s binary search, as Excel runs it: a true bisection that
 * trusts the order it was promised, so on data that is not sorted it
 * lands where Excel lands rather than on what a scan would find. With
 * no exact match, `mode` 1 takes the next larger value and -1 the next
 * smaller — the two sides of where the search stopped.
 */
function bisected(values: readonly CellValue[], wanted: CellValue, descending: boolean, mode: number): number {
  let low = 0;
  let high = values.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const compared = compareValues(values[middle], wanted);
    const order = isError(compared) ? 1 : descending ? -compared : compared;
    if (order === 0) {
      return middle;
    }
    if (order < 0) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  // `low` is the first entry past the wanted value in the search's
  // order and `high` the last before it; which is larger depends on
  // which way the data runs.
  const larger = descending ? high : low;
  const smaller = descending ? low : high;
  const at = mode === 1 ? larger : mode === -1 ? smaller : -1;
  return at >= 0 && at < values.length ? at : -1;
}

/**
 * The row (or column) of the returned range that a lookup landed on:
 * one value for a range one wide, and an array that spills for a wider
 * one, as Excel's `XLOOKUP` gives back a whole record.
 */
function found(returned: Extract<Argument, { kind: 'range' }>, at: number, down: boolean): CellValue | ArrayValue {
  if (down) {
    const row = returned.values.slice(at * returned.columns, (at + 1) * returned.columns);
    return arrayOrValue(1, row.length, row);
  }
  const column: CellValue[] = [];
  for (let row = 0; row < returned.rows; row++) {
    column.push(returned.values[row * returned.columns + at] ?? null);
  }
  return arrayOrValue(column.length, 1, column);
}

/** The closest value below or above, for `XLOOKUP`'s fallback modes. */
function nearest(values: readonly CellValue[], wanted: CellValue, below: boolean): number {
  // The candidates are compared with each other, not with what is
  // wanted: `compareValues` answers -1, 0 or 1, so ranking by it made
  // every candidate on the right side a tie, and the first one won —
  // which is the nearest only when the data happens to be sorted.
  let best = -1;
  for (let at = 0; at < values.length; at++) {
    const order = compareValues(values[at], wanted);
    if (isError(order) || (below ? order > 0 : order < 0)) {
      continue;
    }
    if (best === -1) {
      best = at;
      continue;
    }
    const against = compareValues(values[at], values[best]);
    if (!isError(against) && (below ? against > 0 : against < 0)) {
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
