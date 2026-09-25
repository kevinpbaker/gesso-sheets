/**
 * Addresses: the packed key the store is built on, and the A1
 * references people type.
 *
 * A cell's identity is one number. A `Map` keyed by `"3,7"` costs a
 * string allocation and a hash of it on every lookup, and a recalc
 * does a lookup per edge of the dependency graph; keyed by
 * `row * MAX_COLUMNS + column` it is an integer key, which every
 * engine stores unboxed. The sheet is bounded so that the packing is
 * exact — 1,048,576 × 16,384 is 1.7e10, comfortably inside the safe
 * integer range — and the bounds are the ones spreadsheets already
 * have, so nothing is given up by taking them.
 */

/** Columns A through XFD, as a spreadsheet has. */
export const MAX_COLUMNS = 16_384;
/** Rows 1 through 1,048,576. */
export const MAX_ROWS = 1_048_576;

/** A cell's identity: one integer, from a zero-based row and column. */
export function cellKey(row: number, column: number): number {
  return row * MAX_COLUMNS + column;
}

export function rowOf(key: number): number {
  return Math.floor(key / MAX_COLUMNS);
}

export function columnOf(key: number): number {
  return key % MAX_COLUMNS;
}

export function inBounds(row: number, column: number): boolean {
  return Number.isInteger(row) && Number.isInteger(column) && row >= 0 && row < MAX_ROWS && column >= 0 && column < MAX_COLUMNS;
}

/**
 * A reference as it was written, which is not the same as where it
 * points.
 *
 * The `$` flags are carried rather than resolved because Phase 5's
 * fill handle is the whole reason they exist: extending `=A1+$B$1`
 * down has to adjust the first and leave the second, and it can only
 * do that if the parse remembered which was which.
 */
export interface CellRef {
  readonly row: number;
  readonly column: number;
  readonly rowAbsolute: boolean;
  readonly columnAbsolute: boolean;
}

export interface RangeRef {
  readonly start: CellRef;
  readonly end: CellRef;
}

/** `0` → `A`, `25` → `Z`, `26` → `AA`. */
export function columnName(index: number): string {
  let name = '';
  let remaining = index;
  while (remaining >= 0) {
    name = String.fromCharCode(65 + (remaining % 26)) + name;
    remaining = Math.floor(remaining / 26) - 1;
  }
  return name;
}

/**
 * `A` → `0`, `AA` → `26`, and null for anything that is not a column.
 *
 * Bijective base-26: there is no zero digit, so `AA` is 26 rather than
 * 0, which is why this is not `parseInt` with a radix.
 */
export function columnIndex(name: string): number | null {
  if (name.length === 0) {
    return null;
  }
  let index = 0;
  for (const character of name) {
    const code = character.toUpperCase().charCodeAt(0);
    if (code < 65 || code > 90) {
      return null;
    }
    index = index * 26 + (code - 64);
  }
  return index - 1;
}

const REFERENCE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

/**
 * An A1 reference, or null when the text is not one.
 *
 * Null rather than an error because the caller — the parser — has
 * another reading available: a bare word that is not a reference is a
 * function name or a mistake, and only the parser knows which.
 * Out-of-bounds is a different answer again, and deliberately not made
 * here: `ZZZ9999999` parses to a well-formed reference that points off
 * the sheet, and turning that into `#REF!` is evaluation's job.
 */
export function parseRef(text: string): CellRef | null {
  const match = REFERENCE.exec(text);
  if (match === null) {
    return null;
  }
  const column = columnIndex(match[2]);
  if (column === null) {
    return null;
  }
  const row = Number(match[4]) - 1;
  if (row < 0) {
    return null;
  }
  return {
    row,
    column,
    rowAbsolute: match[3] === '$',
    columnAbsolute: match[1] === '$'
  };
}

export function formatRef(ref: CellRef): string {
  const column = `${ref.columnAbsolute ? '$' : ''}${columnName(ref.column)}`;
  const row = `${ref.rowAbsolute ? '$' : ''}${ref.row + 1}`;
  return `${column}${row}`;
}

export function formatRange(range: RangeRef): string {
  return `${formatRef(range.start)}:${formatRef(range.end)}`;
}

/** A reference with no `$`, which is what a bare address means. */
export function relativeRef(row: number, column: number): CellRef {
  return { row, column, rowAbsolute: false, columnAbsolute: false };
}

/**
 * Every key in a range, in row-major order.
 *
 * The corners are normalised, so `B2:A1` covers what `A1:B2` does —
 * dragging a selection upwards produces the first and means the
 * second.
 */
export function* rangeKeys(range: RangeRef): Generator<number> {
  const firstRow = Math.min(range.start.row, range.end.row);
  const lastRow = Math.max(range.start.row, range.end.row);
  const firstColumn = Math.min(range.start.column, range.end.column);
  const lastColumn = Math.max(range.start.column, range.end.column);
  for (let row = firstRow; row <= lastRow; row++) {
    for (let column = firstColumn; column <= lastColumn; column++) {
      yield cellKey(row, column);
    }
  }
}

/** How many cells a range covers, without walking it. */
export function rangeSize(range: RangeRef): number {
  const rows = Math.abs(range.end.row - range.start.row) + 1;
  const columns = Math.abs(range.end.column - range.start.column) + 1;
  return rows * columns;
}

/**
 * An address as somebody types it into the name box: `B7`, `a1:c9`,
 * `$A$1`.
 *
 * Returns a range because a single cell is one — `B7` is `B7:B7` —
 * which saves every caller a second shape to handle. Null for
 * anything that is not an address at all, which is how the name box
 * knows to leave the selection where it is rather than jumping to
 * somewhere it invented.
 *
 * Deliberately separate from `parseRef`: that one is the parser's,
 * and it answers about a fragment inside a formula. This one is about
 * a whole string, so `B7 ` with a stray space is an address and
 * `B7+1` is not.
 */
export function parseAddress(text: string): RangeRef | null {
  const trimmed = text.trim();
  if (trimmed === '') {
    return null;
  }
  const halves = trimmed.split(':');
  if (halves.length > 2) {
    return null;
  }
  const start = wholeRef(halves[0]);
  if (start === null) {
    return null;
  }
  if (halves.length === 1) {
    return { start, end: start };
  }
  const end = wholeRef(halves[1]);
  return end === null ? null : { start, end };
}

/** `parseRef`, but the whole string has to be the reference. */
function wholeRef(text: string): CellRef | null {
  const ref = parseRef(text.trim());
  return ref !== null && formatRef(ref).toUpperCase() === text.trim().toUpperCase() ? ref : null;
}
