import { isError, toNumber, toText, VALUE, type CellError, type CellValue } from './Values';

/**
 * What every function in the library is built out of.
 *
 * Separate from `Functions.ts` because that file is the registry and
 * this is the vocabulary: the families import these, the registry
 * imports the families, and nothing imports in a circle.
 */

/**
 * An evaluated argument.
 *
 * A range stays a range rather than being flattened by the evaluator,
 * because the difference matters: `SUM("x")` is an error and
 * `SUM(A1:A9)` with text in it is not. Keeping the shape lets each
 * function decide.
 *
 * **And it keeps its dimensions**, which the lookups need and nothing
 * else does. `VLOOKUP` has to know that a range is four columns wide
 * to find the third one; flattened to a list there is no way to ask.
 * Row-major, so the cell at `(r, c)` is `values[r * columns + c]` —
 * the same order `rangeKeys` walks in, which is why no transposition
 * happens anywhere.
 */
export type Argument =
  | { readonly kind: 'value'; readonly value: CellValue }
  | {
      readonly kind: 'range';
      readonly values: readonly CellValue[];
      readonly rows: number;
      readonly columns: number;
    };

/**
 * What a function is allowed to know about the world outside its
 * arguments.
 *
 * Two things, both of which make a function's answer change when
 * nothing it reads has: the clock and the dice. They arrive as a
 * parameter rather than as `Date.now()` and `Math.random()` inside
 * the functions so that a spec can state what `TODAY()` is, which is
 * the only way to assert anything about it at all.
 *
 * `now` is fixed for a whole recalculation, as it is in Excel: two
 * cells calling `NOW()` in the same pass must not disagree about what
 * time it is, or a sheet can contradict itself about its own
 * timestamps.
 */
export interface FunctionContext {
  /** The serial — days since 1899-12-30 — for this recalculation. */
  now(): number;
  random(): number;
}

export type SheetFunction = (args: readonly Argument[], ctx: FunctionContext) => CellValue;

/**
 * The first error in the arguments, which most functions report
 * before doing anything else.
 *
 * Errors travel: a sum of a cell holding `#DIV/0!` is `#DIV/0!`, not
 * zero and not `#VALUE!`. Getting this wrong hides the cell that
 * actually broke behind a cell that merely read it.
 */
export function firstError(args: readonly Argument[]): CellError | null {
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
export function numbersOf(args: readonly Argument[]): number[] | CellError {
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

/** Every value an argument list holds, ranges flattened. */
export function valuesOf(args: readonly Argument[]): CellValue[] {
  const values: CellValue[] = [];
  for (const arg of args) {
    if (arg.kind === 'range') {
      values.push(...arg.values);
    } else {
      values.push(arg.value);
    }
  }
  return values;
}

/**
 * One argument as a single value.
 *
 * A range where a value was wanted takes its first cell, which is
 * what a spreadsheet did with `=ABS(A1:A3)` in the days before it
 * spilled.
 */
export function scalar(args: readonly Argument[], index: number): CellValue {
  const arg = args[index];
  if (arg === undefined) {
    return null;
  }
  return arg.kind === 'value' ? arg.value : (arg.values[0] ?? null);
}

/** One argument as a number, or the error that stopped it. */
export function numberAt(args: readonly Argument[], index: number): number | CellError {
  return toNumber(scalar(args, index));
}

/** One argument as text, or the error that stopped it. */
export function textAt(args: readonly Argument[], index: number): string | CellError {
  return toText(scalar(args, index));
}

/** An integer argument, truncated as a spreadsheet truncates them. */
export function integerAt(args: readonly Argument[], index: number): number | CellError {
  const value = numberAt(args, index);
  return isError(value) ? value : Math.trunc(value);
}

export function arity(args: readonly Argument[], min: number, max: number): CellError | null {
  return args.length < min || args.length > max ? VALUE : null;
}

/** An argument that has to be a range, for the lookups. */
export function rangeAt(args: readonly Argument[], index: number): Extract<Argument, { kind: 'range' }> | CellError {
  const arg = args[index];
  if (arg === undefined || arg.kind !== 'range') {
    return VALUE;
  }
  return arg;
}

/**
 * A guard that runs the arity and error checks together.
 *
 * Almost every function starts with these two lines and getting them
 * the wrong way round changes the answer: `SUM(#REF!, 1, 2, 3)` with
 * a maximum of two arguments is an arity error in Excel, not a
 * `#REF!`, because the call is malformed before its arguments mean
 * anything.
 */
export function checked(args: readonly Argument[], min: number, max: number): CellError | null {
  return arity(args, min, max) ?? firstError(args);
}
