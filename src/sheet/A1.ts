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
  /**
   * The sheet written in front of it, as it was typed, or undefined
   * for a reference to the sheet the formula is on.
   *
   * The **name** and not an id, which is the decision this field is.
   * An id would make renaming free and deleting obvious, and would
   * make the stored formula disagree with what the cell puts back in
   * the editor — because `Cell.input` is exactly what somebody typed
   * and nothing regenerates it. So the name travels in the text, and
   * a rename rewrites the formulas that use it, the way an inserted
   * row rewrites the references past it.
   *
   * Compared case-insensitively, because sheet names are.
   */
  readonly sheet?: string;
}

/**
 * A sheet's name as it is written in front of a reference.
 *
 * Quoted when it is not a plain word, because `Q3 Budget!A1` would
 * tokenize as two words with a space between them and the second half
 * would be read as a reference to a sheet called `Budget`. Inside the
 * quotes a quote doubles, which is the rule the tokenizer reads back.
 */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** `Sheet2!`, or nothing at all when the reference stays home. */
function sheetPrefix(ref: CellRef): string {
  return ref.sheet === undefined ? '' : `${quoteSheetName(ref.sheet)}!`;
}

export interface RangeRef {
  readonly start: CellRef;
  readonly end: CellRef;
  /**
   * `A:A` — every row of those columns, however many there turn out
   * to be.
   *
   * A flag rather than a span from row 0 to row 1,048,575, because
   * the two are not the same claim. A span is a rectangle somebody
   * drew and it means those cells; this means *the column*, and it
   * has to go on meaning the column after a row is added at the
   * bottom. The difference shows up in three places, each of which
   * would otherwise be a million of something: the dependency graph
   * watches the column instead of storing an edge per cell, the
   * evaluator reads only as far as the sheet is used, and `rangeKeys`
   * is never handed one of these at all.
   */
  readonly wholeColumn?: boolean;
}

/** `A:A`, as a range that means the column rather than a rectangle. */
export function wholeColumnRange(first: number, last: number): RangeRef {
  return {
    start: { row: 0, column: first, rowAbsolute: true, columnAbsolute: false },
    end: { row: MAX_ROWS - 1, column: last, rowAbsolute: true, columnAbsolute: false },
    wholeColumn: true
  };
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
  return `${sheetPrefix(ref)}${bareRef(ref)}`;
}

/** The address without the sheet in front of it. */
function bareRef(ref: CellRef): string {
  const column = `${ref.columnAbsolute ? '$' : ''}${columnName(ref.column)}`;
  const row = `${ref.rowAbsolute ? '$' : ''}${ref.row + 1}`;
  return `${column}${row}`;
}

/**
 * A range, with its sheet written once.
 *
 * `Sheet2!A1:B9` rather than `Sheet2!A1:Sheet2!B9`. Both parse, and
 * the second is what a printer built out of two `formatRef` calls
 * would produce — which is how a fill of a cross-sheet formula would
 * have slowly filled the text with repetitions of the sheet name.
 */
export function formatRange(range: RangeRef): string {
  const prefix = sheetPrefix(range.start);
  /**
   * `A:A` goes back as `A:A`.
   *
   * Spelled out it is `A$1:A$1048576`, which is a different claim —
   * a rectangle somebody drew rather than the column — and it is what
   * a fill used to turn `=SUM(A:A)` into the moment it was dragged
   * one cell sideways.
   */
  if (range.wholeColumn === true) {
    return `${prefix}${columnName(range.start.column)}:${columnName(range.end.column)}`;
  }
  return `${prefix}${bareRef(range.start)}:${bareRef(range.end)}`;
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
  if (range.wholeColumn === true) {
    // A million keys, and every caller that could reach here has a
    // better answer available. Throwing is louder than a hang.
    throw new Error('a whole-column reference has no key list; watch the column instead');
  }
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
