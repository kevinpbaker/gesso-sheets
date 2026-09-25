import { columnIndex, inBounds, keyOn, parseRef, rangeKeys, type CellRef, type RangeRef } from './A1';
import type { Ast, BinaryOperator } from './Ast';
import { FUNCTIONS, isSheetFunction, liveContext, type Argument, type FunctionContext } from './Functions';
import {
  compareValues,
  DIV0,
  isError,
  NAME,
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
}

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
    case 'range':
      // A range where a single value is wanted. Functions take ranges
      // through `argumentOf`; anything else — `=A1:B2+1` — has no
      // meaning without the spilling a modern sheet does, so it is the
      // range's first cell, which is what a sheet did before it spilled.
      return readRange(node.range, context).values[0] ?? null;
    case 'unary': {
      const operand = evaluate(node.operand, context);
      const number = toNumber(operand);
      if (isError(number)) {
        return number;
      }
      return node.op === '-' ? -number : number;
    }
    case 'call':
      return call(node.name, node.args, context);
    case 'binary':
      return binary(node.op, node.left, node.right, context);
  }
}

function readRef(ref: CellRef, context: EvaluationContext): CellValue {
  if (!inBounds(ref.row, ref.column)) {
    return REF;
  }
  const key = keyFor(ref, context);
  return key === null ? REF : context.valueAt(key);
}

/** A range's values and its shape, which the lookups need. */
function readRange(range: RangeRef, context: EvaluationContext): Extract<Argument, { kind: 'range' }> {
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
  for (const key of rangeKeys(read, sheet)) {
    values.push(context.valueAt(key));
  }
  return {
    kind: 'range',
    values,
    rows: Math.abs(read.end.row - read.start.row) + 1,
    columns: Math.abs(read.end.column - read.start.column) + 1
  };
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
function call(name: string, args: readonly Ast[], context: EvaluationContext): CellValue {
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
  if (args.length !== 1) {
    return VALUE;
  }
  const text = toText(evaluate(args[0], context));
  if (isError(text)) {
    return text;
  }
  const range = referenceOf(text.trim());
  if (range === null) {
    return REF;
  }
  return range.start === range.end
    ? readRef(range.start, context)
    : (readRange(range, context).values[0] ?? null);
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
  const start = anchor.kind === 'ref' ? anchor.ref : anchor.kind === 'range' ? anchor.range.start : null;
  if (start === null) {
    return VALUE;
  }
  const numbers: number[] = [];
  for (let at = 1; at < args.length; at++) {
    const value = toNumber(evaluate(args[at], context));
    if (isError(value)) {
      return value;
    }
    numbers.push(Math.trunc(value));
  }
  const [downBy, acrossBy, height = 1, width = 1] = numbers;
  if (height < 1 || width < 1) {
    return VALUE;
  }
  const row = start.row + downBy;
  const column = start.column + acrossBy;
  if (!inBounds(row, column) || !inBounds(row + height - 1, column + width - 1)) {
    return null;
  }
  return {
    start: { row, column, rowAbsolute: true, columnAbsolute: true },
    end: { row: row + height - 1, column: column + width - 1, rowAbsolute: true, columnAbsolute: true }
  };
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
function argumentOf(node: Ast, context: EvaluationContext): Argument {
  if (node.kind === 'range') {
    return readRange(node.range, context);
  }
  /**
   * A named range is a range argument, which is the whole point of
   * naming one: `=SUM(Sales)` has to sum the range, not take its
   * first cell.
   */
  if (node.kind === 'call' && node.args.length === 0 && !isSheetFunction(node.name)) {
    const named = context.rangeForName?.(node.name) ?? null;
    if (named !== null) {
      return readRange(named, context);
    }
  }
  if (node.kind === 'call' && node.name === 'OFFSET') {
    const range = offsetRange(node.args, context);
    if (range === null) {
      return { kind: 'value', value: REF };
    }
    return isError(range) ? { kind: 'value', value: range } : readRange(range, context);
  }
  return { kind: 'value', value: evaluate(node, context) };
}

function binary(op: BinaryOperator, leftNode: Ast, rightNode: Ast, context: EvaluationContext): CellValue {
  const left = evaluate(leftNode, context);
  const right = evaluate(rightNode, context);
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
      const power = a ** b;
      // `(-8) ** (1/3)` is NaN rather than -2, and a NaN in a cell is
      // a value that compares false with itself and poisons everything
      // downstream silently. An error says so.
      return Number.isNaN(power) ? VALUE : power;
    }
  }
}
