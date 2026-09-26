import { arity, arrayOrValue, integerAt, type Argument, type ArrayValue, type SheetFunction } from './FunctionKit';
import { compareValues, isError, NA, toBoolean, VALUE, type CellValue } from './Values';

/**
 * Functions whose answer is an array: it spills from the formula's
 * cell into the cells beside and below it — see `Workbook.settle`.
 * Read anywhere one value is wanted, an array is its top-left value.
 *
 * `FILTER` with nothing left says `#N/A` rather than Excel's `#CALC!`,
 * an error code this sheet does not have; the third argument, which
 * says what to show instead, is the way round it in both.
 */

/** An argument as a grid of values, whether it was a range, an array or one value. */
interface Grid {
  readonly rows: number;
  readonly columns: number;
  readonly values: readonly CellValue[];
}

function gridAt(args: readonly Argument[], index: number): Grid {
  const arg = args[index];
  if (arg === undefined) {
    return { rows: 0, columns: 0, values: [] };
  }
  return arg.kind === 'range'
    ? { rows: arg.rows, columns: arg.columns, values: arg.values }
    : { rows: 1, columns: 1, values: [arg.value] };
}

const cellOf = (grid: Grid, row: number, column: number): CellValue => grid.values[row * grid.columns + column] ?? null;

function rowsOfGrid(grid: Grid): CellValue[][] {
  const rows: CellValue[][] = [];
  for (let row = 0; row < grid.rows; row++) {
    rows.push(grid.values.slice(row * grid.columns, (row + 1) * grid.columns));
  }
  return rows;
}

function transposed(grid: Grid): Grid {
  const values: CellValue[] = [];
  for (let column = 0; column < grid.columns; column++) {
    for (let row = 0; row < grid.rows; row++) {
      values.push(cellOf(grid, row, column));
    }
  }
  return { rows: grid.columns, columns: grid.rows, values };
}

function fromRows(rows: readonly (readonly CellValue[])[], columns: number): CellValue | ArrayValue {
  return arrayOrValue(rows.length, columns, rows.flat());
}

/** A whole-row key, so two rows with the same values in them are the same row. */
function rowKey(row: readonly CellValue[]): string {
  return JSON.stringify(row.map(value => (typeof value === 'string' ? value.toLowerCase() : value)));
}

/** A flag argument: absent is false, and anything else is read as `IF` would read it. */
function flagAt(args: readonly Argument[], index: number): boolean | ReturnType<typeof toBoolean> {
  const arg = args[index];
  if (arg === undefined) {
    return false;
  }
  return toBoolean(arg.kind === 'range' ? (arg.single ?? arg.values[0] ?? null) : arg.value);
}

export const ARRAY_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** Rows as columns and columns as rows. */
  TRANSPOSE(args) {
    const wrong = arity(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const grid = transposed(gridAt(args, 0));
    return grid.rows * grid.columns === 0 ? null : arrayOrValue(grid.rows, grid.columns, grid.values);
  },

  /** `SEQUENCE(5)` is 1 to 5 down a column; rows, columns, a start and a step. */
  SEQUENCE(args) {
    const wrong = arity(args, 1, 4);
    if (wrong !== null) {
      return wrong;
    }
    const numbers: number[] = [];
    for (const [index, fallback] of [
      [0, 1],
      [1, 1],
      [2, 1],
      [3, 1]
    ] as const) {
      if (index >= args.length || (args[index].kind === 'value' && (args[index] as { value: CellValue }).value === null)) {
        numbers.push(fallback);
        continue;
      }
      const value = index < 2 ? integerAt(args, index) : integerOrNumber(args, index);
      if (isError(value)) {
        return value;
      }
      numbers.push(value);
    }
    const [rows, columns, start, step] = numbers;
    if (rows < 1 || columns < 1 || rows * columns > 1_000_000) {
      return VALUE;
    }
    const values: number[] = [];
    for (let at = 0; at < rows * columns; at++) {
      values.push(start + at * step);
    }
    return arrayOrValue(rows, columns, values);
  },

  /** The rows (or columns) of an array where a test is true. */
  FILTER(args) {
    const wrong = arity(args, 2, 3);
    if (wrong !== null) {
      return wrong;
    }
    const grid = gridAt(args, 0);
    const include = gridAt(args, 1);
    const byRow = include.rows === grid.rows && include.columns === 1;
    const byColumn = include.columns === grid.columns && include.rows === 1 && !byRow;
    if (!byRow && !byColumn) {
      return VALUE;
    }
    const kept: CellValue[][] = [];
    const source = byRow ? rowsOfGrid(grid) : rowsOfGrid(transposed(grid));
    for (const [at, row] of source.entries()) {
      const test = toBoolean(include.values[at] ?? null);
      if (isError(test)) {
        return test;
      }
      if (test) {
        kept.push(row);
      }
    }
    if (kept.length === 0) {
      if (args.length < 3) {
        return NA;
      }
      const fallback = args[2];
      return fallback.kind === 'range' ? arrayOrValue(fallback.rows, fallback.columns, fallback.values) : fallback.value;
    }
    const width = byRow ? grid.columns : grid.rows;
    if (byRow) {
      return fromRows(kept, width);
    }
    const back = transposed({ rows: kept.length, columns: width, values: kept.flat() });
    return arrayOrValue(back.rows, back.columns, back.values);
  },

  /** An array in order: by its first column, ascending, unless told otherwise. */
  SORT(args) {
    const wrong = arity(args, 1, 4);
    if (wrong !== null) {
      return wrong;
    }
    const byColumn = flagAt(args, 3);
    if (isError(byColumn)) {
      return byColumn;
    }
    const grid = byColumn ? transposed(gridAt(args, 0)) : gridAt(args, 0);
    const index = args.length > 1 && !(args[1].kind === 'value' && args[1].value === null) ? integerAt(args, 1) : 1;
    const order = args.length > 2 && !(args[2].kind === 'value' && args[2].value === null) ? integerAt(args, 2) : 1;
    if (isError(index)) {
      return index;
    }
    if (isError(order)) {
      return order;
    }
    if (index < 1 || index > grid.columns || (order !== 1 && order !== -1)) {
      return VALUE;
    }
    // Stable, as Excel's is: rows that tie keep the order they came in.
    const rows = rowsOfGrid(grid)
      .map((row, at) => ({ row, at }))
      .sort((a, b) => {
        const compared = compareValues(a.row[index - 1] ?? null, b.row[index - 1] ?? null);
        const by = isError(compared) ? 0 : compared * order;
        return by !== 0 ? by : a.at - b.at;
      })
      .map(entry => entry.row);
    if (!byColumn) {
      return fromRows(rows, grid.columns);
    }
    const back = transposed({ rows: rows.length, columns: grid.columns, values: rows.flat() });
    return arrayOrValue(back.rows, back.columns, back.values);
  },

  /** Each distinct row once, in the order first seen — or only the rows seen once. */
  UNIQUE(args) {
    const wrong = arity(args, 1, 3);
    if (wrong !== null) {
      return wrong;
    }
    const byColumn = flagAt(args, 1);
    const once = flagAt(args, 2);
    if (isError(byColumn)) {
      return byColumn;
    }
    if (isError(once)) {
      return once;
    }
    const grid = byColumn ? transposed(gridAt(args, 0)) : gridAt(args, 0);
    const counts = new Map<string, { row: CellValue[]; count: number }>();
    for (const row of rowsOfGrid(grid)) {
      const key = rowKey(row);
      const seen = counts.get(key);
      if (seen === undefined) {
        counts.set(key, { row, count: 1 });
      } else {
        seen.count++;
      }
    }
    const rows = [...counts.values()].filter(entry => !once || entry.count === 1).map(entry => entry.row);
    if (rows.length === 0) {
      return NA;
    }
    if (!byColumn) {
      return fromRows(rows, grid.columns);
    }
    const back = transposed({ rows: rows.length, columns: grid.columns, values: rows.flat() });
    return arrayOrValue(back.rows, back.columns, back.values);
  }
};

function integerOrNumber(args: readonly Argument[], index: number): number | ReturnType<typeof integerAt> {
  const arg = args[index];
  const value = arg.kind === 'range' ? (arg.single ?? arg.values[0] ?? null) : arg.value;
  if (typeof value === 'number') {
    return value;
  }
  return integerAt(args, index);
}
