/**
 * What a cell holds, and how one kind becomes another.
 *
 * Six error values, and no seventh. They are tagged objects rather
 * than strings because a cell may legitimately hold the *text*
 * `#REF!` — someone pasted it from a report — and a sheet that cannot
 * tell that apart from the error is a sheet that cannot be trusted
 * about either.
 *
 * It was five until Phase 11. `#N/A` is the sixth, and it arrived
 * with the lookups because it is the answer to a question they are
 * constantly asked: `VLOOKUP` that found nothing has not failed, and
 * calling that `#VALUE!` would say the formula is wrong when the
 * formula is fine and the table simply does not have that row.
 * `IFNA` exists to catch exactly this one and nothing else, which is
 * only a coherent function if the code is its own.
 *
 * `#NUM!` is deliberately still absent. `SQRT(-1)` and `LN(0)` give
 * `#VALUE!` here: a value of the wrong kind was passed, which is what
 * `#VALUE!` says, and a seventh code earns less than it costs.
 *
 * `null` is an empty cell, and it is not the same as `''`. An empty
 * cell is zero in arithmetic and the empty string in text, which is
 * what the coercions below do; a cell holding `''` is text that
 * happens to be empty, and `COUNT` treats them differently.
 */

export type ErrorCode = '#REF!' | '#DIV/0!' | '#NAME?' | '#VALUE!' | '#CIRC!' | '#N/A';

export interface CellError {
  readonly kind: 'error';
  readonly code: ErrorCode;
}

/** A literal, an error, or an empty cell. */
export type CellValue = number | string | boolean | CellError | null;

function error(code: ErrorCode): CellError {
  return Object.freeze({ kind: 'error', code });
}

/** A reference that does not point at a cell. */
export const REF = error('#REF!');
/** Division by zero, and the empty AVERAGE. */
export const DIV0 = error('#DIV/0!');
/** A name the sheet does not know: an unknown function, mostly. */
export const NAME = error('#NAME?');
/** A value of the wrong kind, and a formula that does not parse. */
export const VALUE = error('#VALUE!');
/** A cell that depends, however indirectly, on itself. */
export const CIRC = error('#CIRC!');
/** A lookup that found nothing. Not a failure — an absence. */
export const NA = error('#N/A');

/**
 * Takes `unknown` rather than `CellValue` because it is the guard the
 * whole engine uses to unwrap an "answer or error" return, and those
 * are not all cell values: `numbersOf` returns `number[] | CellError`
 * and wants the same one-line check as everything else.
 */
export function isError(value: unknown): value is CellError {
  return typeof value === 'object' && value !== null && (value as CellError).kind === 'error';
}

/**
 * A cell's value as a number.
 *
 * Returns the error rather than throwing, because every caller is in
 * the middle of evaluating something and an error is a value here, not
 * an exception. Text that looks like a number is one — `="5"+1` is 6 in
 * every spreadsheet — and text that does not is `#VALUE!`.
 */
export function toNumber(value: CellValue): number | CellError {
  if (isError(value)) {
    return value;
  }
  if (value === null) {
    return 0;
  }
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }
  const trimmed = value.trim();
  if (trimmed === '') {
    return 0;
  }
  const parsed = Number(trimmed);
  return Number.isNaN(parsed) ? VALUE : parsed;
}

export function toText(value: CellValue): string | CellError {
  if (isError(value)) {
    return value;
  }
  if (value === null) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  return typeof value === 'number' ? formatNumber(value) : value;
}

export function toBoolean(value: CellValue): boolean | CellError {
  if (isError(value)) {
    return value;
  }
  if (value === null) {
    return false;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  const upper = value.trim().toUpperCase();
  if (upper === 'TRUE') {
    return true;
  }
  if (upper === 'FALSE') {
    return false;
  }
  return VALUE;
}

/**
 * What the screen shows for a value.
 *
 * This is the last step before the wire in Phase 2 — the contract
 * carries display strings, not values — so the rule about where
 * formatting happens is enforced by this being the only place that
 * turns a number into text for a person.
 */
export function formatValue(value: CellValue): string {
  if (isError(value)) {
    return value.code;
  }
  if (value === null) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  return typeof value === 'number' ? formatNumber(value) : value;
}

/**
 * A number as a person reads it.
 *
 * Binary floating point is the whole problem: `0.1 + 0.2` is
 * 0.30000000000000004, and a spreadsheet that showed that would be
 * reporting its own arithmetic rather than the user's. Fifteen
 * significant digits is what spreadsheets settle on — enough that no
 * honest result is altered, few enough that the representation error
 * never surfaces.
 */
export function formatNumber(value: number): string {
  if (!Number.isFinite(value)) {
    return value > 0 ? '∞' : Number.isNaN(value) ? 'NaN' : '-∞';
  }
  if (Number.isInteger(value) && Math.abs(value) < 1e15) {
    return String(value);
  }
  return String(Number(value.toPrecision(15)));
}

/** Orders two values for `<`, `>` and friends, or reports a mismatch. */
export function compareValues(left: CellValue, right: CellValue): number | CellError {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }
  // Numbers sort before text and text before booleans, which is the
  // order spreadsheets use when the two sides are not the same kind.
  const leftRank = rank(left);
  const rightRank = rank(right);
  if (leftRank !== rightRank) {
    return leftRank < rightRank ? -1 : 1;
  }
  if (leftRank === 0) {
    const a = left === null ? 0 : (left as number);
    const b = right === null ? 0 : (right as number);
    return a === b ? 0 : a < b ? -1 : 1;
  }
  if (leftRank === 1) {
    // Case-insensitive, as a spreadsheet's comparison is.
    const a = (left as string).toUpperCase();
    const b = (right as string).toUpperCase();
    return a === b ? 0 : a < b ? -1 : 1;
  }
  const a = left as boolean;
  const b = right as boolean;
  return a === b ? 0 : a ? 1 : -1;
}

function rank(value: Exclude<CellValue, CellError>): 0 | 1 | 2 {
  if (value === null || typeof value === 'number') {
    return 0;
  }
  return typeof value === 'string' ? 1 : 2;
}
