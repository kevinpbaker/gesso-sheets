import { columnIndex, inBounds, keyOn, parseRef, rangeKeys, rowOf, type CellRef, type RangeRef } from './A1';
import type { Ast, BinaryOperator, CallNode } from './Ast';
import { FUNCTIONS, isSheetFunction, LIFTED, liveContext, type Argument, type FunctionContext } from './Functions';
import { arrayOrValue, isArray, type ArrayValue } from './FunctionKit';
import {
  CALC,
  compareValues,
  DIV0,
  isError,
  NA,
  NAME,
  NUM,
  REF,
  toBoolean,
  toNumber,
  toText,
  VALUE,
  type CellError,
  type CellValue
} from './Values';

/** Where an expression reads its cells from. */
export interface EvaluationContext {
  valueAt(key: number): CellValue;
  /**
   * The clock and the dice, for the four volatile functions.
   *
   * Optional so that a spec can evaluate an expression against a bare
   * `Map` without inventing a time, which most of them want to do.
   */
  functions?: FunctionContext;
  /**
   * How far down the sheet has ever been written, which is how far a
   * whole-column reference reads.
   *
   * `=SUM(A:A)` means the column, and the column is 1,048,576 cells
   * long; reading them all would take a second per evaluation to add
   * up a million nothings. A high-water mark is enough because the
   * cells past it are empty by definition, and erring high only costs
   * a few blank reads — erring low would silently drop data.
   */
  usedRows?: number;
  /**
   * The range a name stands for, for named ranges.
   *
   * Optional, so a spec can evaluate against a bare `Map` without a
   * name table — which is what most of them want, and what the
   * evaluator's own specs have always done.
   */
  rangeForName?(name: string): RangeRef | null;
  /**
   * The workbook around this sheet, when there is one.
   *
   * Absent means a workbook of one sheet — which is what every spec
   * that evaluates an expression against a bare `Map` means, and what
   * the application meant for twelve phases. A reference that names a
   * sheet then has nowhere to resolve to and reads `#REF!`, which is
   * the same answer it gets for a sheet that was deleted.
   */
  book?: WorkbookContext;
  /**
   * Which sheet the formula being evaluated lives on.
   *
   * An unqualified `A1` means A1 *here*, so this is half of what a
   * reference resolves through. Defaults to the first sheet.
   */
  onSheet?: number;
  /**
   * The cell whose formula this is, for `ROW()` and `COLUMN()` with no
   * argument — the only functions whose answer is where they are.
   * Absent for an expression that is not in a cell, which then has no
   * row to give and says `#VALUE!`.
   */
  at?: { readonly row: number; readonly column: number };
  /**
   * Whether a row is hidden by hand, filtered out, or neither.
   *
   * `SUBTOTAL`'s whole reason to exist: `SUBTOTAL(9, …)` under a
   * filter totals what the filter shows. Rows are the document's, not
   * the sheet's, so this is asked rather than known.
   */
  rowState?(sheet: number, row: number): 'hidden' | 'filtered' | null;
  /** Whether a cell's formula calls `SUBTOTAL`, which a `SUBTOTAL` skips. */
  isSubtotal?(key: number): boolean;
  /**
   * How far the array in a cell spills, for `A1#`: null when the cell
   * holds no array that spilled. Absent, nothing spills.
   */
  spillAt?(key: number): { readonly rows: number; readonly columns: number } | null;
  /**
   * The formula a name holds, for a name that holds one rather than a
   * range: `TaxRate` as `=0.2`, `Double` as `=LAMBDA(x, x*2)`. Absent,
   * or null, and the name is a range or nothing.
   */
  formulaForName?(name: string): Ast | null;
  /**
   * The names a `LET` or a `LAMBDA`'s parameters have bound, innermost
   * last. Set by the evaluator as it goes in; nothing outside it
   * builds one.
   */
  locals?: Scope;
  /** How many `LAMBDA`s deep this evaluation is; see `MAX_DEPTH`. */
  depth?: number;
}

/**
 * A function a formula made: `LAMBDA(x, y, x*y)` evaluated.
 *
 * It closes over the names that were bound where it was made, so
 * `LET(rate, 0.2, LAMBDA(x, x*rate))` hands back a function that still
 * knows the rate wherever it ends up being called. It is a value inside
 * the evaluator only: a cell that would hold one holds `#CALC!`, as in
 * Excel, because there is nothing a cell could show for it.
 */
export interface LambdaValue {
  readonly kind: 'lambda';
  /** The parameters, upper-cased as names are looked up. */
  readonly params: readonly string[];
  readonly body: Ast;
  readonly scope: Scope;
}

/** What a name inside a formula can stand for. */
export type Bound = CellValue | ArrayValue | LambdaValue;

/** Names bound by `LET` and `LAMBDA`, upper-cased. */
export type Scope = ReadonlyMap<string, Bound>;

const NO_LOCALS: Scope = new Map();

export function isLambda(value: unknown): value is LambdaValue {
  return typeof value === 'object' && value !== null && (value as { kind?: unknown }).kind === 'lambda';
}

/**
 * How deep a `LAMBDA` may call itself.
 *
 * Recursion is how Excel's users loop, so it is allowed — and a
 * recursion with no floor is a tab that stops answering, so past this
 * it is `#NUM!`. Deep enough for a real loop over a list, shallow
 * enough that the stack it takes is far from the engine's own limit.
 */
export const MAX_DEPTH = 512;

/** A test for the cells a read should leave out; see `evaluateSubtotal`. */
type Skip = (key: number, sheet: number, row: number) => boolean;

/** What the evaluator needs to know about the sheets around it. */
export interface WorkbookContext {
  /** The index of a sheet by name, case-insensitively, or null. */
  sheetFor(name: string): number | null;
  /** How far down a given sheet has been written; see `usedRows`. */
  usedRowsOf(sheet: number): number;
}

/**
 * The key a reference points at, or null when it points nowhere.
 *
 * Null is `#REF!`: a reference naming a sheet the workbook does not
 * have. That is what a formula holds after the sheet it read was
 * deleted, and Excel answers it the same way.
 */
function keyFor(ref: CellRef, context: EvaluationContext): number | null {
  const sheet = sheetIndex(ref.sheet, context);
  return sheet === null ? null : keyOn(sheet, ref.row, ref.column);
}

function sheetIndex(name: string | undefined, context: EvaluationContext): number | null {
  if (name === undefined) {
    return context.onSheet ?? 0;
  }
  return context.book?.sheetFor(name) ?? null;
}

/**
 * One expression, against values that are already correct.
 *
 * The evaluator is deliberately ignorant of recalculation order: it
 * reads whatever the context holds and assumes it is current, because
 * `Sheet` evaluates in topological order and so it always is. That
 * separation is what keeps this a pure function of the tree and the
 * store, testable on its own with a `Map`.
 */
export function evaluate(node: Ast, context: EvaluationContext): CellValue {
  switch (node.kind) {
    case 'number':
      return node.value;
    case 'text':
      return node.value;
    case 'boolean':
      return node.value;
    case 'error':
      return { kind: 'error', code: node.code };
    case 'ref':
      return readRef(node.ref, context);
    case 'range': {
      // A range where one value is wanted — `=A1:A9*2` — gives the one
      // implicit intersection picks; see `readRange`.
      const read = readRange(node.range, context);
      return read.single !== undefined ? read.single : (read.values[0] ?? null);
    }
    case 'unary':
      if (node.op === '@') {
        return implicit(node.operand, context);
      }
      return negate(node.op, evaluate(node.operand, context));
    case 'call': {
      const bound = boundCall(node, context);
      if (bound !== undefined) {
        return scalarOf(bound);
      }
      const result = call(node.name, node.args, context);
      return isArray(result) ? (result.values[0] ?? null) : result;
    }
    case 'invoke':
      return scalarOf(invoke(node.callee, node.args, context));
    case 'binary':
      return combine(node.op, evaluate(node.left, context), evaluate(node.right, context));
  }
}

/**
 * An expression as a formula's own cell sees it: arrays and all.
 *
 * `evaluate` answers one value, which is what a rule, a validation and
 * every argument that wants a number need. A cell can hold more: a
 * range, or arithmetic over one — `=B2:D4 + E2:G4` — is an array, and
 * the cell it is written in spills it across the cells beside it, as
 * Excel has done since dynamic arrays. `@` asks for the one value
 * instead, by implicit intersection; it is what a formula from an older
 * `.xlsx` is given wherever its author meant one value.
 */
export function evaluateArray(node: Ast, context: EvaluationContext): CellValue | ArrayValue {
  switch (node.kind) {
    case 'range': {
      const read = readRange(node.range, context);
      return read.rows * read.columns === 0 ? null : arrayOrValue(read.rows, read.columns, read.values);
    }
    case 'unary': {
      if (node.op === '@') {
        return implicit(node.operand, context);
      }
      const operand = evaluateArray(node.operand, context);
      return isArray(operand) ? mapArray(operand, value => negate(node.op, value)) : negate(node.op, operand);
    }
    case 'binary': {
      const left = evaluateArray(node.left, context);
      const right = evaluateArray(node.right, context);
      if (!isArray(left) && !isArray(right)) {
        return combine(node.op, left, right);
      }
      return lifted(left, right, (a, b) => combine(node.op, a, b));
    }
    case 'call': {
      const bound = boundCall(node, context);
      if (bound !== undefined) {
        return settled(bound);
      }
      if (node.name === 'IF') {
        return arrayIf(node.args, context);
      }
      if (node.name === 'ANCHORARRAY') {
        const range = spillRange(node.args, context);
        if (isError(range)) {
          return range;
        }
        const read = readRange(range, context);
        return arrayOrValue(read.rows, read.columns, read.values);
      }
      // A computed reference is as much a range as a written one.
      if (node.name === 'OFFSET' || node.name === 'INDIRECT') {
        const range = node.name === 'OFFSET' ? offsetRange(node.args, context) : indirectRange(node.args, context);
        if (range === null) {
          return REF;
        }
        if (isError(range)) {
          return range;
        }
        const read = readRange(range, context);
        return read.rows * read.columns === 0 ? null : arrayOrValue(read.rows, read.columns, read.values);
      }
      if (node.args.length === 0 && !isSheetFunction(node.name)) {
        const named = context.rangeForName?.(node.name) ?? null;
        if (named !== null) {
          const read = readRange(named, context);
          return read.rows * read.columns === 0 ? null : arrayOrValue(read.rows, read.columns, read.values);
        }
      }
      if (LIFTED.has(node.name)) {
        return liftedCall(node.name, node.args, context);
      }
      return call(node.name, node.args, context);
    }
    case 'invoke':
      return settled(invoke(node.callee, node.args, context));
    default:
      return evaluate(node, context);
  }
}

/**
 * A function of single values, handed an array: run once per cell.
 *
 * The shapes follow the operators' rule — a single value goes with
 * every cell, a row with every row, a column with every column — so
 * `=ROUND(A1:A9, B1:B9)` rounds each value to its own places and
 * `=ROUND(A1:A9, 2)` all of them to two. Handed no array at all, it is
 * the one call it always was.
 */
function liftedCall(name: string, args: readonly Ast[], context: EvaluationContext): CellValue | ArrayValue {
  const evaluated = args.map(arg => argumentOf(arg, context));
  const wide = (arg: Argument): arg is Extract<Argument, { kind: 'range' }> =>
    arg.kind === 'range' && arg.rows * arg.columns > 1;
  if (!evaluated.some(wide)) {
    const result = FUNCTIONS[name](evaluated, context.functions ?? liveContext());
    return result;
  }
  const shapes = evaluated.map(arg => (wide(arg) ? { kind: 'array' as const, rows: arg.rows, columns: arg.columns, values: arg.values } : null));
  const rows = Math.max(...shapes.map(shape => shape?.rows ?? 1));
  const columns = Math.max(...shapes.map(shape => shape?.columns ?? 1));
  const values: CellValue[] = [];
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      const one = evaluated.map((arg, at): Argument => {
        const shape = shapes[at];
        return shape === null ? arg : { kind: 'value', value: elementOf(shape, row, column) };
      });
      const result = FUNCTIONS[name](one, context.functions ?? liveContext());
      values.push(isArray(result) ? (result.values[0] ?? null) : result);
    }
  }
  return { kind: 'array', rows, columns, values };
}

/** `@`: one value where a range or an array might be, by implicit intersection. */
function implicit(node: Ast, context: EvaluationContext): CellValue {
  if (node.kind === 'call' && context.locals?.has(node.name) === true) {
    return evaluate(node, context);
  }
  if (node.kind === 'range') {
    const read = readRange(node.range, context);
    return read.single !== undefined ? read.single : (read.values[0] ?? null);
  }
  if (node.kind === 'call' && node.args.length === 0 && !isSheetFunction(node.name)) {
    const named = context.rangeForName?.(node.name) ?? null;
    if (named !== null) {
      const read = readRange(named, context);
      return read.single !== undefined ? read.single : (read.values[0] ?? null);
    }
  }
  return evaluate(node, context);
}

function negate(op: '-' | '+' | '@', operand: CellValue): CellValue {
  // A leading `+` changes nothing, text included: `=+A1` is A1 in
  // Excel, whatever A1 holds. Only `-` needs a number.
  if (op !== '-') {
    return operand;
  }
  const number = toNumber(operand);
  return isError(number) ? number : -number;
}

function mapArray(array: ArrayValue, each: (value: CellValue) => CellValue): ArrayValue {
  return { kind: 'array', rows: array.rows, columns: array.columns, values: array.values.map(each) };
}

/**
 * Two operands, at least one an array, taken cell by cell.
 *
 * Excel's rule for the shapes: a single value goes with every cell, a
 * single row with every row and a single column with every column; the
 * result is as large as the larger in each direction, and where the
 * smaller has nothing to pair it is `#N/A`.
 */
function lifted(
  left: CellValue | ArrayValue,
  right: CellValue | ArrayValue,
  each: (a: CellValue, b: CellValue) => CellValue
): ArrayValue {
  const rows = Math.max(rowsOf(left), rowsOf(right));
  const columns = Math.max(columnsOf(left), columnsOf(right));
  const values: CellValue[] = [];
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      values.push(each(elementOf(left, row, column), elementOf(right, row, column)));
    }
  }
  return { kind: 'array', rows, columns, values };
}

const rowsOf = (value: CellValue | ArrayValue): number => (isArray(value) ? value.rows : 1);
const columnsOf = (value: CellValue | ArrayValue): number => (isArray(value) ? value.columns : 1);

function elementOf(value: CellValue | ArrayValue, row: number, column: number): CellValue {
  if (!isArray(value)) {
    return value;
  }
  const r = value.rows === 1 ? 0 : row;
  const c = value.columns === 1 ? 0 : column;
  return r < value.rows && c < value.columns ? (value.values[r * value.columns + c] ?? null) : NA;
}

/**
 * `IF` over an array condition: each cell takes its own branch, which
 * is what `=IF(B2:B9>100, "big", "small")` means. Over one condition it
 * is as lazy as it always was, and the branch taken may be an array.
 */
function arrayIf(args: readonly Ast[], context: EvaluationContext): CellValue | ArrayValue {
  if (args.length < 2 || args.length > 3) {
    return VALUE;
  }
  const condition = evaluateArray(args[0], context);
  if (!isArray(condition)) {
    const test = toBoolean(condition);
    if (isError(test)) {
      return test;
    }
    if (test) {
      return evaluateArray(args[1], context);
    }
    return args.length === 3 ? evaluateArray(args[2], context) : false;
  }
  const yes = evaluateArray(args[1], context);
  const no = args.length === 3 ? evaluateArray(args[2], context) : false;
  const rows = Math.max(condition.rows, rowsOf(yes), rowsOf(no));
  const columns = Math.max(condition.columns, columnsOf(yes), columnsOf(no));
  const values: CellValue[] = [];
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      const test = toBoolean(elementOf(condition, row, column));
      values.push(isError(test) ? test : test ? elementOf(yes, row, column) : elementOf(no, row, column));
    }
  }
  return { kind: 'array', rows, columns, values };
}

function readRef(ref: CellRef, context: EvaluationContext): CellValue {
  if (!inBounds(ref.row, ref.column)) {
    return REF;
  }
  const key = keyFor(ref, context);
  return key === null ? REF : context.valueAt(key);
}

/** A range's values and its shape, which the lookups need. */
function readRange(range: RangeRef, context: EvaluationContext, skip?: Skip): Extract<Argument, { kind: 'range' }> {
  if (!inBounds(range.start.row, range.start.column) || !inBounds(range.end.row, range.end.column)) {
    return { kind: 'range', values: [REF], rows: 1, columns: 1 };
  }
  const sheet = sheetIndex(range.start.sheet, context);
  if (sheet === null) {
    return { kind: 'range', values: [REF], rows: 1, columns: 1 };
  }
  const read = range.wholeColumn === true ? usedPartOf(range, sheet, context) : range;
  if (read === null) {
    // A whole column of a sheet nothing has been written to.
    const columns = Math.abs(range.end.column - range.start.column) + 1;
    return { kind: 'range', values: [], rows: 0, columns };
  }
  const values: CellValue[] = [];
  const single = intersection(read, sheet, context);
  if (skip !== undefined) {
    // A read with holes in it has no shape to report, and the one
    // caller that skips — `SUBTOTAL` — hands the values to aggregates
    // that do not ask for one.
    for (const key of rangeKeys(read, sheet)) {
      if (!skip(key, sheet, rowOf(key))) {
        values.push(context.valueAt(key));
      }
    }
    return { kind: 'range', values, rows: values.length, columns: 1 };
  }
  for (const key of rangeKeys(read, sheet)) {
    values.push(context.valueAt(key));
  }
  return {
    kind: 'range',
    values,
    rows: Math.abs(read.end.row - read.start.row) + 1,
    columns: Math.abs(read.end.column - read.start.column) + 1,
    ...(single === undefined ? {} : { single })
  };
}

/**
 * The one value a range gives where one value is wanted: implicit
 * intersection, which is Excel's rule for `=A1:A9*2` and for a range
 * handed to `ABS` or `LEN`.
 *
 * A single cell is itself. A column gives the cell in the formula's own
 * row, and a row the cell in its own column — which is why `=B1:B9*2`
 * typed down column C doubles the B beside each row — and a range that
 * does not pass through the formula, or is two-dimensional, is
 * `#VALUE!`. Without a formula to stand in (an expression evaluated in
 * no cell), the first cell stands in, as it always did.
 */
function intersection(range: RangeRef, sheet: number, context: EvaluationContext): CellValue | undefined {
  const firstRow = Math.min(range.start.row, range.end.row);
  const lastRow = Math.max(range.start.row, range.end.row);
  const firstColumn = Math.min(range.start.column, range.end.column);
  const lastColumn = Math.max(range.start.column, range.end.column);
  if (firstRow === lastRow && firstColumn === lastColumn) {
    return context.valueAt(keyOn(sheet, firstRow, firstColumn));
  }
  const at = context.at;
  if (at === undefined) {
    return undefined;
  }
  if (firstColumn === lastColumn && at.row >= firstRow && at.row <= lastRow) {
    return context.valueAt(keyOn(sheet, at.row, firstColumn));
  }
  if (firstRow === lastRow && at.column >= firstColumn && at.column <= lastColumn) {
    return context.valueAt(keyOn(sheet, firstRow, at.column));
  }
  return VALUE;
}

/** A whole-column reference cut down to the rows that can hold anything. */
function usedPartOf(range: RangeRef, sheet: number, context: EvaluationContext): RangeRef | null {
  const used = context.book?.usedRowsOf(sheet) ?? context.usedRows ?? 0;
  if (used <= 0) {
    return null;
  }
  return {
    start: range.start,
    end: { ...range.end, row: Math.min(range.end.row, used - 1) }
  };
}

/**
 * A call, with the five special forms handled before the table.
 *
 * Three of them are lazy and two of them compute references. See
 * `Functions.ts`, where the reasoning for each is set out; this is
 * only where they are dispatched.
 */
function call(name: string, args: readonly Ast[], context: EvaluationContext): CellValue | ArrayValue {
  switch (name) {
    case 'IF':
      return evaluateIf(args, context);
    case 'IFS':
      return evaluateIfs(args, context);
    case 'SWITCH':
      return evaluateSwitch(args, context);
    case 'INDIRECT':
      return evaluateIndirect(args, context);
    case 'OFFSET':
      return evaluateOffset(args, context);
    case 'ROW':
      return evaluatePosition(args, context, 'row');
    case 'COLUMN':
      return evaluatePosition(args, context, 'column');
    case 'SUBTOTAL':
      return evaluateSubtotal(args, context);
    case 'ANCHORARRAY': {
      const range = spillRange(args, context);
      return isError(range) ? range : (readRange(range, context).values[0] ?? null);
    }
    default:
      break;
  }

  if (!isSheetFunction(name)) {
    /**
     * A bare word the sheet has no function for may be a named range.
     *
     * Checked *after* the functions, because a function name is not
     * available to be a name — `Names.nameProblem` refuses anything
     * that already means something — and checking this way round
     * means a sheet can never shadow its own library.
     *
     * Only a word with no brackets: `Sales()` is somebody calling a
     * function that does not exist, and saying `#NAME?` to it is the
     * true answer.
     */
    if (args.length === 0) {
      const named = context.rangeForName?.(name) ?? null;
      if (named !== null) {
        return readRange(named, context).values[0] ?? null;
      }
    }
    return NAME;
  }
  const evaluated: Argument[] = args.map(arg => argumentOf(arg, context));
  return FUNCTIONS[name](evaluated, context.functions ?? liveContext());
}

/**
 * `ROW` and `COLUMN`: where a reference is, counted from one.
 *
 * Special forms because they read a reference's *position* rather than
 * its value, and a function in the table is only ever handed values.
 * With no argument they answer for the cell the formula is in, which
 * is what `=ROW()` numbering a list down the side is. A range answers
 * for its first row or column, which is what Excel shows when the
 * array it would return has nowhere to spill.
 */
function evaluatePosition(args: readonly Ast[], context: EvaluationContext, axis: 'row' | 'column'): CellValue {
  if (args.length > 1) {
    return VALUE;
  }
  if (args.length === 0) {
    return context.at === undefined ? VALUE : context.at[axis] + 1;
  }
  const node = args[0];
  if (node.kind === 'ref') {
    return node.ref[axis] + 1;
  }
  const range =
    node.kind === 'range'
      ? node.range
      : node.kind === 'call' && node.args.length === 0
        ? (context.rangeForName?.(node.name) ?? null)
        : null;
  if (range === null) {
    return VALUE;
  }
  return Math.min(range.start[axis], range.end[axis]) + 1;
}

/**
 * Which aggregate `SUBTOTAL`'s first argument names: 1 to 11, and
 * 101 to 111 for the same ones leaving out hidden rows as well.
 */
const SUBTOTALS = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV', 'STDEVP', 'SUM', 'VAR', 'VARP'];

/**
 * `SUBTOTAL`: an aggregate that leaves things out, which is the point.
 *
 * Three things, all Excel's. Rows a filter took away are never
 * counted, so a total under a filtered list is the total of what the
 * filter shows. Rows hidden by hand are counted by 1–11 and not by
 * 101–111, which is the only difference between the two sets. And a
 * cell that is itself a `SUBTOTAL` is never counted, so a grand total
 * over a column of subtotals adds up the data once rather than twice.
 *
 * A special form because all three are questions about *cells* — which
 * row, whose formula — and a function in the table is handed values.
 */
function evaluateSubtotal(args: readonly Ast[], context: EvaluationContext): CellValue {
  if (args.length < 2) {
    return VALUE;
  }
  const which = toNumber(evaluate(args[0], context));
  if (isError(which)) {
    return which;
  }
  const code = Math.trunc(which);
  const name = SUBTOTALS[(code > 100 ? code - 100 : code) - 1];
  if (name === undefined) {
    return VALUE;
  }
  const leaveHidden = code > 100;
  const skip: Skip = (key, sheet, row) => {
    if (context.isSubtotal?.(key) === true) {
      return true;
    }
    const state = context.rowState?.(sheet, row) ?? null;
    return state === 'filtered' || (leaveHidden && state === 'hidden');
  };
  const evaluated = args.slice(1).map(arg => {
    if (arg.kind === 'ref') {
      return readRange({ start: arg.ref, end: arg.ref }, context, skip);
    }
    return argumentOf(arg, context, skip);
  });
  const result = FUNCTIONS[name](evaluated, context.functions ?? liveContext());
  return isArray(result) ? (result.values[0] ?? null) : result;
}

/**
 * `IF` has to be lazy or it is a trap.
 *
 * `=IF(B1=0, "n/a", A1/B1)` is the way everyone writes a guarded
 * division, and a sheet that evaluated both branches would answer
 * `#DIV/0!` to the formula written specifically to avoid it.
 */
function evaluateIf(args: readonly Ast[], context: EvaluationContext): CellValue {
  if (args.length < 2 || args.length > 3) {
    return VALUE;
  }
  const condition = toBoolean(evaluate(args[0], context));
  if (isError(condition)) {
    return condition;
  }
  if (condition) {
    return evaluate(args[1], context);
  }
  return args.length === 3 ? evaluate(args[2], context) : false;
}

/**
 * The first branch whose test passes, out of any number of pairs.
 *
 * What a chain of nested `IF`s is trying to say, and lazy for the
 * same reason: the later branches are exactly the ones that would
 * divide by zero if the earlier test is what stopped them.
 */
function evaluateIfs(args: readonly Ast[], context: EvaluationContext): CellValue {
  if (args.length < 2 || args.length % 2 !== 0) {
    return VALUE;
  }
  for (let at = 0; at + 1 < args.length; at += 2) {
    const condition = toBoolean(evaluate(args[at], context));
    if (isError(condition)) {
      return condition;
    }
    if (condition) {
      return evaluate(args[at + 1], context);
    }
  }
  // No branch matched and none was named as the fallback. `#N/A` is
  // Excel's answer and the right one: the sheet was asked a question
  // it has no case for.
  return { kind: 'error', code: '#N/A' };
}

/**
 * One value against a list of cases, with an optional default.
 *
 * `SWITCH(A1, 1, "one", 2, "two", "other")` — an odd number of
 * arguments after the subject means the last is the fallback.
 */
function evaluateSwitch(args: readonly Ast[], context: EvaluationContext): CellValue {
  if (args.length < 3) {
    return VALUE;
  }
  const subject = evaluate(args[0], context);
  if (isError(subject)) {
    return subject;
  }
  let at = 1;
  for (; at + 1 < args.length; at += 2) {
    const candidate = evaluate(args[at], context);
    if (isError(candidate)) {
      return candidate;
    }
    const order = compareValues(subject, candidate);
    if (!isError(order) && order === 0) {
      return evaluate(args[at + 1], context);
    }
  }
  // One argument left over is the default.
  return at < args.length ? evaluate(args[at], context) : { kind: 'error', code: '#N/A' };
}

/**
 * A reference built out of text, at the moment it runs.
 *
 * The dangerous one, and the reason `Sheet` re-derives a formula's
 * precedents after evaluating it: nothing in `=INDIRECT("A" & B1)`
 * mentions the cell it ends up reading, so the graph built from the
 * *tree* has no edge to it. A cell that is stale and never woken is
 * the worst bug a spreadsheet can have, because it is silent.
 */
function evaluateIndirect(args: readonly Ast[], context: EvaluationContext): CellValue {
  const range = indirectRange(args, context);
  if (range === null) {
    return REF;
  }
  if (isError(range)) {
    return range;
  }
  return range.start === range.end
    ? readRef(range.start, context)
    : (readRange(range, context).values[0] ?? null);
}

/** The rectangle an `INDIRECT` names, the error that stopped it, or null for `#REF!`. */
function indirectRange(args: readonly Ast[], context: EvaluationContext): RangeRef | CellError | null {
  if (args.length !== 1) {
    return VALUE;
  }
  const text = toText(evaluate(args[0], context));
  if (isError(text)) {
    return text;
  }
  return referenceOf(text.trim());
}

/**
 * A range moved, and optionally resized, from a starting point.
 *
 * `OFFSET(A1, 2, 0, 3, 1)` is `A3:A5`. Like `INDIRECT` it names cells
 * that are nowhere in the formula's text, so its edges are re-derived
 * after it runs.
 *
 * It returns a single value here, because a range only means
 * something to the function that receives it and this returns to an
 * expression. `SUM(OFFSET(...))` is the form people want and it needs
 * `OFFSET` to be a range argument, which is `argumentOf`'s job below.
 */
function evaluateOffset(args: readonly Ast[], context: EvaluationContext): CellValue {
  const range = offsetRange(args, context);
  if (range === null) {
    return REF;
  }
  if (isError(range)) {
    return range;
  }
  return readRange(range, context).values[0] ?? null;
}

/** The rectangle an `OFFSET` names, the error that stopped it, or null for `#REF!`. */
function offsetRange(args: readonly Ast[], context: EvaluationContext): RangeRef | CellError | null {
  if (args.length < 3 || args.length > 5) {
    return VALUE;
  }
  const anchor = args[0];
  const base: RangeRef | null =
    anchor.kind === 'ref' ? { start: anchor.ref, end: anchor.ref } : anchor.kind === 'range' ? anchor.range : null;
  if (base === null) {
    return VALUE;
  }
  // The base's top-left corner, and its size as the default size: in
  // Excel `OFFSET(K7:L8, 0, 0)` is K7:L8, not K7.
  const start = {
    ...base.start,
    row: Math.min(base.start.row, base.end.row),
    column: Math.min(base.start.column, base.end.column)
  };
  const baseHeight = Math.abs(base.end.row - base.start.row) + 1;
  const baseWidth = Math.abs(base.end.column - base.start.column) + 1;
  const numbers: number[] = [];
  for (let at = 1; at < args.length; at++) {
    const value = toNumber(evaluate(args[at], context));
    if (isError(value)) {
      return value;
    }
    numbers.push(Math.trunc(value));
  }
  const [downBy, acrossBy, height = baseHeight, width = baseWidth] = numbers;
  if (height < 1 || width < 1) {
    return VALUE;
  }
  const row = start.row + downBy;
  const column = start.column + acrossBy;
  if (!inBounds(row, column) || !inBounds(row + height - 1, column + width - 1)) {
    return null;
  }
  return {
    start: { ...start, row, column, rowAbsolute: true, columnAbsolute: true },
    end: { ...start, row: row + height - 1, column: column + width - 1, rowAbsolute: true, columnAbsolute: true }
  };
}

/**
 * `A1#`: the rectangle A1's array fills, from A1. `#REF!` for a cell
 * that holds no array, or one that did not spill, as in Excel.
 */
function spillRange(args: readonly Ast[], context: EvaluationContext): RangeRef | CellError {
  const anchor = args[0];
  if (args.length !== 1 || anchor?.kind !== 'ref') {
    return VALUE;
  }
  const key = keyFor(anchor.ref, context);
  const spill = key === null ? null : (context.spillAt?.(key) ?? null);
  if (spill === null) {
    return REF;
  }
  const start = { ...anchor.ref, rowAbsolute: true, columnAbsolute: true };
  return { start, end: { ...start, row: start.row + spill.rows - 1, column: start.column + spill.columns - 1 } };
}

/** `A1`, `A1:B9` or a column letter pair, as a rectangle. */
function referenceOf(text: string): RangeRef | null {
  const halves = text.split(':');
  if (halves.length > 2) {
    return null;
  }
  const start = parseRef(halves[0]);
  if (start === null) {
    return wholeColumns(halves);
  }
  if (halves.length === 1) {
    return { start, end: start };
  }
  const end = parseRef(halves[1]);
  return end === null ? null : { start, end };
}

/** `INDIRECT("A:A")`, which names a column rather than a cell. */
function wholeColumns(halves: readonly string[]): RangeRef | null {
  if (halves.length !== 2) {
    return null;
  }
  const first = columnIndex(halves[0].replace('$', ''));
  const last = columnIndex(halves[1].replace('$', ''));
  if (first === null || last === null) {
    return null;
  }
  return {
    start: { row: 0, column: first, rowAbsolute: true, columnAbsolute: true },
    end: { row: 0, column: last, rowAbsolute: true, columnAbsolute: true }
  };
}

/**
 * A range argument stays a range; everything else is one value.
 *
 * `OFFSET` is the exception that earns its own branch: it *is* a
 * range, computed rather than written, and `SUM(OFFSET(A1, 0, 0, 5,
 * 1))` is the whole reason anybody uses it.
 */
function argumentOf(node: Ast, context: EvaluationContext, skip?: Skip): Argument {
  if (node.kind === 'range') {
    return readRange(node.range, context, skip);
  }
  /**
   * A reference is a range of one, as Excel has it: `SUM(A1, B1)`
   * leaves out text in A1 exactly as `SUM(A1:B1)` does, where a literal
   * `SUM("x")` is a claim that "x" is a number. Handed over as a value,
   * a text cell next to the figures made the whole total `#VALUE!`.
   */
  if (node.kind === 'ref') {
    if (!inBounds(node.ref.row, node.ref.column)) {
      return { kind: 'value', value: REF };
    }
    return keyFor(node.ref, context) === null
      ? { kind: 'value', value: REF }
      : readRange({ start: node.ref, end: node.ref }, context, skip);
  }
  /**
   * A named range is a range argument, which is the whole point of
   * naming one: `=SUM(Sales)` has to sum the range, not take its
   * first cell.
   */
  if (node.kind === 'call' && node.args.length === 0 && !isSheetFunction(node.name) && context.locals?.has(node.name) !== true) {
    const named = context.rangeForName?.(node.name) ?? null;
    if (named !== null) {
      return readRange(named, context, skip);
    }
  }
  if (node.kind === 'call' && node.name === 'OFFSET') {
    const range = offsetRange(node.args, context);
    if (range === null) {
      return { kind: 'value', value: REF };
    }
    return isError(range) ? { kind: 'value', value: range } : readRange(range, context);
  }
  /**
   * Anything else that comes out an array — `A1:A9*B1:B9`, `--(B5:B20)`
   * — is handed over as one, which is what `SUMPRODUCT` over arithmetic
   * and `SUM(A1:A3*B1:B3)` need. A formula from an older file that
   * meant one value there says `@`, and gets one.
   */
  const value = evaluateArray(node, context);
  return isArray(value)
    ? { kind: 'range', values: value.values, rows: value.rows, columns: value.columns }
    : { kind: 'value', value };
}

/** Two values and an operator: every binary operator's arithmetic and comparison. */
function combine(op: BinaryOperator, left: CellValue, right: CellValue): CellValue {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }

  if (op === '&') {
    const a = toText(left);
    if (isError(a)) {
      return a;
    }
    const b = toText(right);
    return isError(b) ? b : a + b;
  }

  if (op === '=' || op === '<>' || op === '<' || op === '<=' || op === '>' || op === '>=') {
    const order = compareValues(left, right);
    if (isError(order)) {
      return order;
    }
    switch (op) {
      case '=':
        return order === 0;
      case '<>':
        return order !== 0;
      case '<':
        return order < 0;
      case '<=':
        return order <= 0;
      case '>':
        return order > 0;
      default:
        return order >= 0;
    }
  }

  const a = toNumber(left);
  if (isError(a)) {
    return a;
  }
  const b = toNumber(right);
  if (isError(b)) {
    return b;
  }
  switch (op) {
    case '+':
      return a + b;
    case '-':
      return a - b;
    case '*':
      return a * b;
    case '/':
      return b === 0 ? DIV0 : a / b;
    default: {
      // Excel's answers for the powers that have none: 0^0 and a
      // negative base to a fractional power are out of range, 0 to a
      // negative power divides by zero, and an overflow is out of range.
      // A NaN must never reach a cell; it compares false with itself.
      if (a === 0 && b === 0) {
        return NUM;
      }
      if (a === 0 && b < 0) {
        return DIV0;
      }
      const power = a ** b;
      return Number.isFinite(power) ? power : NUM;
    }
  }
}

// ---------------------------------------------------------------------
// LET, LAMBDA, and the helpers that take one
// ---------------------------------------------------------------------

/** The helpers that are handed a `LAMBDA` and call it. */
const HELPERS: ReadonlySet<string> = new Set(['MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY']);

/**
 * A call that binds or calls a name the formula made, or undefined for
 * every other call.
 *
 * The one place the evaluator looks before its usual paths: a name a
 * `LET` bound shadows the workbook's names, a name that holds a formula
 * is that formula, and `LET`, `LAMBDA` and the helpers answer with
 * values that may be functions. Everything else — which is every call
 * in a sheet that uses none of this — comes back undefined, having cost
 * a lookup in an empty scope and a set test.
 */
function boundCall(node: CallNode, context: EvaluationContext): Bound | undefined {
  const local = context.locals?.get(node.name);
  if (local !== undefined) {
    // `x`, or `f(3)` where a LET bound `f` to a LAMBDA.
    if (node.word !== undefined) {
      return local;
    }
    if (isLambda(local)) {
      return apply(local, node.args, context);
    }
    // `ROW(row)` inside `LAMBDA(row, …)`: brackets after a name bound to
    // a value mean the function of that name, if there is one.
    if (isSheetFunction(node.name)) {
      return undefined;
    }
    return isError(local) ? local : VALUE;
  }
  switch (node.name) {
    case 'LET':
      return evaluateLet(node.args, context);
    case 'LAMBDA':
      return makeLambda(node.args, context);
    default:
      break;
  }
  if (HELPERS.has(node.name)) {
    return helper(node.name, node.args, context);
  }
  if (isSheetFunction(node.name)) {
    return undefined;
  }
  const formula = context.formulaForName?.(node.name) ?? null;
  if (formula === null) {
    return undefined;
  }
  const value = namedValue(formula, context);
  if (node.word !== undefined) {
    return value;
  }
  return isLambda(value) ? apply(value, node.args, context) : isError(value) ? value : VALUE;
}

/**
 * What a name that holds a formula stands for, worked out where it is read.
 *
 * In the caller's workbook and on the caller's sheet, so an unqualified
 * `A1` in it means A1 of the sheet reading it — the rule a name holding a
 * range already follows — but with none of the caller's `LET` names in
 * scope: a name means the same thing wherever it is used. One level
 * deeper, so a name defined in terms of itself stops.
 */
function namedValue(formula: Ast, context: EvaluationContext): Bound {
  const depth = (context.depth ?? 0) + 1;
  if (depth > MAX_DEPTH) {
    return NUM;
  }
  return evaluateBound(formula, { ...context, locals: NO_LOCALS, depth });
}

/** An expression as anything it can be: a value, an array, or a function. */
export function evaluateBound(node: Ast, context: EvaluationContext): Bound {
  if (node.kind === 'call') {
    return boundCall(node, context) ?? evaluateArray(node, context);
  }
  if (node.kind === 'invoke') {
    return invoke(node.callee, node.args, context);
  }
  return evaluateArray(node, context);
}

/** `(…)(args)`: whatever the callee comes to, called. */
function invoke(callee: Ast, args: readonly Ast[], context: EvaluationContext): Bound {
  const fn = evaluateBound(callee, context);
  if (isLambda(fn)) {
    return apply(fn, args, context);
  }
  return isError(fn) ? fn : VALUE;
}

/** A bound value where one cell's worth is wanted. */
function scalarOf(value: Bound): CellValue {
  if (isLambda(value)) {
    return CALC;
  }
  return isArray(value) ? (value.values[0] ?? null) : value;
}

/** A bound value where a cell's answer is wanted: a function is `#CALC!`. */
function settled(value: Bound): CellValue | ArrayValue {
  return isLambda(value) ? CALC : value;
}

/** The name a `LET` or `LAMBDA` binds, or null when the argument is not one. */
function bindingName(node: Ast): string | null {
  if (node.kind !== 'call' || node.word === undefined || node.args.length > 0) {
    return null;
  }
  // A function's name is allowed — `LAMBDA(row, SUM(row))` is how Excel's
  // own examples read — because a bare word looks here first and a call
  // with brackets still finds the function.
  return node.name;
}

/**
 * `LET(name, value, …, result)`.
 *
 * Left to right, each name visible to every value after it and to the
 * result; a name bound again later in the same `LET` is refused, as in
 * Excel, rather than quietly meaning its last value.
 */
function evaluateLet(args: readonly Ast[], context: EvaluationContext): Bound {
  if (args.length < 3 || args.length % 2 === 0) {
    return VALUE;
  }
  const scope = new Map(context.locals ?? NO_LOCALS);
  const inner: EvaluationContext = { ...context, locals: scope };
  const own = new Set<string>();
  for (let at = 0; at + 1 < args.length; at += 2) {
    const name = bindingName(args[at]);
    if (name === null || own.has(name)) {
      return VALUE;
    }
    own.add(name);
    scope.set(name, evaluateBound(args[at + 1], inner));
  }
  return evaluateBound(args[args.length - 1], inner);
}

/** `LAMBDA(x, y, body)`: the function, closed over the names in scope. */
function makeLambda(args: readonly Ast[], context: EvaluationContext): Bound {
  if (args.length === 0) {
    return VALUE;
  }
  const params: string[] = [];
  for (const arg of args.slice(0, -1)) {
    const name = bindingName(arg);
    if (name === null || params.includes(name)) {
      return VALUE;
    }
    params.push(name);
  }
  return { kind: 'lambda', params, body: args[args.length - 1], scope: context.locals ?? NO_LOCALS };
}

/** A function called with the expressions written for it. */
function apply(fn: LambdaValue, args: readonly Ast[], context: EvaluationContext): Bound {
  if (args.length !== fn.params.length) {
    return VALUE;
  }
  return applyValues(
    fn,
    args.map(arg => evaluateBound(arg, context)),
    context
  );
}

/**
 * A function called with values: its parameters bound over the scope it
 * was made in, its body evaluated there.
 *
 * The workbook half of the context is the caller's, so the body reads
 * the cells of the sheet it is running on.
 */
function applyValues(fn: LambdaValue, values: readonly Bound[], context: EvaluationContext): Bound {
  const depth = (context.depth ?? 0) + 1;
  if (depth > MAX_DEPTH) {
    return NUM;
  }
  const scope = new Map(fn.scope);
  for (let at = 0; at < fn.params.length; at++) {
    scope.set(fn.params[at], values[at] ?? null);
  }
  try {
    return evaluateBound(fn.body, { ...context, locals: scope, depth });
  } catch (error) {
    // The engine's own stack, reached before the depth above — a body
    // that recurses through many frames at each level. The same answer.
    if (error instanceof RangeError) {
      return NUM;
    }
    throw error;
  }
}

/** A helper's function argument, or the error it came to instead. */
function lambdaArg(node: Ast | undefined, params: number, context: EvaluationContext): LambdaValue | CellError {
  if (node === undefined) {
    return VALUE;
  }
  const fn = evaluateBound(node, context);
  if (!isLambda(fn)) {
    return isError(fn) ? fn : VALUE;
  }
  return fn.params.length === params ? fn : VALUE;
}

/** An argument as an array, a single value being an array of one. */
function arrayArg(node: Ast, context: EvaluationContext): ArrayValue | CellError {
  const value = evaluateBound(node, context);
  if (isLambda(value)) {
    return CALC;
  }
  if (isArray(value)) {
    return value;
  }
  return { kind: 'array', rows: 1, columns: 1, values: [value] };
}

/**
 * One cell's worth out of a function's answer, for a helper that builds
 * an array a cell at a time. An array where one value belongs is
 * `#CALC!`, which is Excel's answer: arrays of arrays are not a thing a
 * sheet can show.
 */
function oneValue(value: Bound): CellValue {
  if (isLambda(value)) {
    return CALC;
  }
  if (isArray(value)) {
    return value.rows === 1 && value.columns === 1 ? (value.values[0] ?? null) : CALC;
  }
  return value;
}

/** Excel's ceiling on what one helper may build, the same as `SEQUENCE`'s. */
const MAX_CELLS = 1_000_000;

/**
 * `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` and `MAKEARRAY`: the reason a
 * `LAMBDA` is worth writing. Each takes its function last.
 */
function helper(name: string, args: readonly Ast[], context: EvaluationContext): Bound {
  switch (name) {
    case 'MAP': {
      // MAP(array, …, fn): the function called on each position, with
      // one value from every array; shapes pair as the operators' do.
      if (args.length < 2) {
        return VALUE;
      }
      const arrays: ArrayValue[] = [];
      for (const arg of args.slice(0, -1)) {
        const array = arrayArg(arg, context);
        if (isError(array)) {
          return array;
        }
        arrays.push(array);
      }
      const fn = lambdaArg(args[args.length - 1], arrays.length, context);
      if (isError(fn)) {
        return fn;
      }
      const rows = Math.max(...arrays.map(array => array.rows));
      const columns = Math.max(...arrays.map(array => array.columns));
      const values: CellValue[] = [];
      for (let row = 0; row < rows; row++) {
        for (let column = 0; column < columns; column++) {
          values.push(oneValue(applyValues(fn, arrays.map(array => elementOf(array, row, column)), context)));
        }
      }
      return arrayOrValue(rows, columns, values);
    }
    case 'REDUCE':
    case 'SCAN': {
      // REDUCE(initial, array, fn(total, value)): the array folded into
      // one answer, row by row. SCAN keeps every step, in the array's shape.
      if (args.length !== 2 && args.length !== 3) {
        return VALUE;
      }
      const initial: Bound = args.length === 3 ? evaluateBound(args[0], context) : null;
      if (isLambda(initial)) {
        return CALC;
      }
      const array = arrayArg(args[args.length - 2], context);
      if (isError(array)) {
        return array;
      }
      const fn = lambdaArg(args[args.length - 1], 2, context);
      if (isError(fn)) {
        return fn;
      }
      let total: Bound = initial;
      const steps: CellValue[] = [];
      for (const value of array.values) {
        total = applyValues(fn, [total, value], context);
        if (name === 'SCAN') {
          steps.push(oneValue(total));
        }
      }
      if (name === 'SCAN') {
        return arrayOrValue(array.rows, array.columns, steps);
      }
      return isLambda(total) ? CALC : total;
    }
    case 'BYROW':
    case 'BYCOL': {
      // BYROW(array, fn(row)): one answer per row, down a column; BYCOL
      // the same across.
      if (args.length !== 2) {
        return VALUE;
      }
      const array = arrayArg(args[0], context);
      if (isError(array)) {
        return array;
      }
      const fn = lambdaArg(args[1], 1, context);
      if (isError(fn)) {
        return fn;
      }
      const values: CellValue[] = [];
      if (name === 'BYROW') {
        for (let row = 0; row < array.rows; row++) {
          const line = array.values.slice(row * array.columns, (row + 1) * array.columns);
          values.push(oneValue(applyValues(fn, [arrayOrValue(1, array.columns, line)], context)));
        }
        return arrayOrValue(array.rows, 1, values);
      }
      for (let column = 0; column < array.columns; column++) {
        const line: CellValue[] = [];
        for (let row = 0; row < array.rows; row++) {
          line.push(array.values[row * array.columns + column] ?? null);
        }
        values.push(oneValue(applyValues(fn, [arrayOrValue(array.rows, 1, line)], context)));
      }
      return arrayOrValue(1, array.columns, values);
    }
    case 'MAKEARRAY': {
      // MAKEARRAY(rows, columns, fn(row, column)), counted from one.
      if (args.length !== 3) {
        return VALUE;
      }
      const rows = toNumber(evaluate(args[0], context));
      if (isError(rows)) {
        return rows;
      }
      const columns = toNumber(evaluate(args[1], context));
      if (isError(columns)) {
        return columns;
      }
      const height = Math.trunc(rows);
      const width = Math.trunc(columns);
      if (height < 1 || width < 1 || height * width > MAX_CELLS) {
        return VALUE;
      }
      const fn = lambdaArg(args[2], 2, context);
      if (isError(fn)) {
        return fn;
      }
      const values: CellValue[] = [];
      for (let row = 1; row <= height; row++) {
        for (let column = 1; column <= width; column++) {
          values.push(oneValue(applyValues(fn, [row, column], context)));
        }
      }
      return arrayOrValue(height, width, values);
    }
    default:
      return NAME;
  }
}
