import { cellKey, inBounds, rangeKeys, type CellRef, type RangeRef } from './A1';
import type { Ast, BinaryOperator } from './Ast';
import { FUNCTIONS, isSheetFunction, type Argument } from './Functions';
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
  type CellValue
} from './Values';

/** Where an expression reads its cells from. */
export interface EvaluationContext {
  valueAt(key: number): CellValue;
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
      return readRange(node.range, context)[0] ?? null;
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
  return context.valueAt(cellKey(ref.row, ref.column));
}

function readRange(range: RangeRef, context: EvaluationContext): CellValue[] {
  if (!inBounds(range.start.row, range.start.column) || !inBounds(range.end.row, range.end.column)) {
    return [REF];
  }
  const values: CellValue[] = [];
  for (const key of rangeKeys(range)) {
    values.push(context.valueAt(key));
  }
  return values;
}

/**
 * A call, with `IF` handled before its arguments are touched.
 *
 * `IF` has to be lazy or it is a trap: `=IF(B1=0, "n/a", A1/B1)` is
 * the way everyone writes a guarded division, and a sheet that
 * evaluated both branches would answer `#DIV/0!` to the formula
 * written specifically to avoid it. It is the only function that
 * needs this, which is why it lives here and not in `Functions`.
 */
function call(name: string, args: readonly Ast[], context: EvaluationContext): CellValue {
  if (name === 'IF') {
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

  if (!isSheetFunction(name)) {
    return NAME;
  }
  const evaluated: Argument[] = args.map(arg => argumentOf(arg, context));
  return FUNCTIONS[name](evaluated);
}

/** A range argument stays a range; everything else is one value. */
function argumentOf(node: Ast, context: EvaluationContext): Argument {
  if (node.kind === 'range') {
    return { kind: 'range', values: readRange(node.range, context) };
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
