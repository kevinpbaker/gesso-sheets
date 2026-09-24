import {
  DIV0,
  isError,
  toNumber,
  toText,
  VALUE,
  type CellError,
  type CellValue
} from './Values';

/**
 * An evaluated argument.
 *
 * A range stays a range rather than being flattened by the evaluator,
 * because the difference matters: `COUNT` counts the numbers in a
 * range and ignores its blanks, while `COUNT` of a blank *cell* passed
 * directly counts nothing either — but `SUM("x")` is an error and
 * `SUM(A1:A9)` with text in it is not. Keeping the shape lets each
 * function decide.
 */
export type Argument =
  | { readonly kind: 'value'; readonly value: CellValue }
  | { readonly kind: 'range'; readonly values: readonly CellValue[] };

export type SheetFunction = (args: readonly Argument[]) => CellValue;

/**
 * The first error in the arguments, which every function reports
 * before doing anything else.
 *
 * Errors travel: a sum of a cell holding `#DIV/0!` is `#DIV/0!`, not
 * zero and not `#VALUE!`. Getting this wrong hides the cell that
 * actually broke behind a cell that merely read it.
 */
function firstError(args: readonly Argument[]): CellError | null {
  for (const arg of args) {
    if (arg.kind === 'value') {
      if (isError(arg.value)) {
        return arg.value;
      }
      continue;
    }
    for (const value of arg.values) {
      if (isError(value)) {
        return value;
      }
    }
  }
  return null;
}

/**
 * The numbers an aggregate should see.
 *
 * A range contributes only its numbers — text and blanks in a column
 * of figures are skipped rather than being an error, which is what
 * makes `SUM(A1:A100)` usable on a column with a heading in it. A
 * scalar argument is coerced, so `SUM(A1, "3")` is arithmetic and
 * `SUM(A1, "x")` is `#VALUE!`: passing something directly is a claim
 * that it is a number.
 */
function numbersOf(args: readonly Argument[]): number[] | CellError {
  const numbers: number[] = [];
  for (const arg of args) {
    if (arg.kind === 'range') {
      for (const value of arg.values) {
        if (typeof value === 'number') {
          numbers.push(value);
        }
      }
      continue;
    }
    if (arg.value === null) {
      continue;
    }
    const number = toNumber(arg.value);
    if (isError(number)) {
      return number;
    }
    numbers.push(number);
  }
  return numbers;
}

function scalar(args: readonly Argument[], index: number): CellValue {
  const arg = args[index];
  if (arg === undefined) {
    return null;
  }
  // A range where a value was wanted takes the range's first cell,
  // which is what a spreadsheet does with `=ABS(A1:A3)` in the days
  // before it spilled.
  return arg.kind === 'value' ? arg.value : (arg.values[0] ?? null);
}

function arity(args: readonly Argument[], min: number, max: number): CellError | null {
  return args.length < min || args.length > max ? VALUE : null;
}

export const FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  SUM(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    let total = 0;
    for (const number of numbers) {
      total += number;
    }
    return total;
  },

  AVERAGE(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    // Nothing to average is the canonical `#DIV/0!`, not zero: zero
    // would be an answer, and there is no answer.
    if (numbers.length === 0) {
      return DIV0;
    }
    let total = 0;
    for (const number of numbers) {
      total += number;
    }
    return total / numbers.length;
  },

  MIN(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    return numbers.length === 0 ? 0 : Math.min(...numbers);
  },

  MAX(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const numbers = numbersOf(args);
    if (isError(numbers)) {
      return numbers;
    }
    return numbers.length === 0 ? 0 : Math.max(...numbers);
  },

  /**
   * How many numbers there are.
   *
   * Errors are *not* propagated: counting is the one thing you can
   * still do over a range with a broken cell in it, and a `COUNT` that
   * reported `#DIV/0!` would make a sheet unusable exactly when the
   * count is what you need.
   */
  COUNT(args) {
    let count = 0;
    for (const arg of args) {
      if (arg.kind === 'range') {
        for (const value of arg.values) {
          if (typeof value === 'number') {
            count++;
          }
        }
        continue;
      }
      if (typeof arg.value === 'number') {
        count++;
      }
    }
    return count;
  },

  ROUND(args) {
    const wrong = arity(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const value = toNumber(scalar(args, 0));
    if (isError(value)) {
      return value;
    }
    const places = args.length > 1 ? toNumber(scalar(args, 1)) : 0;
    if (isError(places)) {
      return places;
    }
    const factor = 10 ** Math.trunc(places);
    // Away from zero on a tie, which is what a spreadsheet does and
    // `Math.round` does not: `Math.round(-2.5)` is -2.
    const scaled = value * factor;
    const rounded = scaled < 0 ? -Math.round(-scaled) : Math.round(scaled);
    return rounded / factor;
  },

  ABS(args) {
    const wrong = arity(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    const value = toNumber(scalar(args, 0));
    return isError(value) ? value : Math.abs(value);
  },

  CONCAT(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    let text = '';
    for (const arg of args) {
      const values = arg.kind === 'range' ? arg.values : [arg.value];
      for (const value of values) {
        const piece = toText(value);
        if (isError(piece)) {
          return piece;
        }
        text += piece;
      }
    }
    return text;
  }
};

/** `IF` is absent here on purpose; see `Evaluator`. */
export function isSheetFunction(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(FUNCTIONS, name);
}
