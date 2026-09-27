import { arrayOrValue, isArray, type ArrayValue } from '../sheet/FunctionKit';
import { isSheetFunction } from '../sheet/Functions';
import type { ScriptFunctions } from '../sheet/ScriptFunctions';
import { CALC, isError, NAME, NUM, VALUE, type CellError, type CellValue } from '../sheet/Values';
import type { CellFunctions, FunctionResult, FunctionScalar, FunctionValue } from '../script/CellFunctions';

/**
 * The workbook's script functions, as the evaluator calls them — Part Six.
 *
 * Between the two sides, and nothing else: `CellFunctions` knows the
 * interpreter and plain data, the evaluator knows cell values and
 * arrays, and this turns one into the other.
 *
 * - An argument that is an error is the call's answer, without calling:
 *   `=TAX(#N/A, 2)` is `#N/A`, as it is for nearly every built-in.
 * - An answer that is not a finite number, text, true or false, nothing,
 *   or rows of them is an error in the cell. A throw is `#VALUE!`; a
 *   call past its deadline or its memory, or an answer no cell can hold,
 *   is `#CALC!`. The message goes beside the cell, for the formula bar.
 * - A formula writes a name in capitals, and JavaScript cares about
 *   case, so each function is found by its name upper-cased. Two that
 *   differ only by case, or one named like a built-in, cannot be called,
 *   and `define` says so rather than picking one.
 */
export class SheetFunctions implements ScriptFunctions {
  /** The JavaScript name behind each name a formula can write. */
  private callable = new Map<string, string>();
  /** Why a cell's call failed, by `sheet:row:column`, while it has. */
  private readonly messages = new Map<string, string>();
  private spent = 0;

  constructor(
    private readonly functions: CellFunctions,
    /** How long a slice may spend in scripts before it hands the thread back. */
    private readonly sliceMilliseconds = 8,
    private readonly now: () => number = () => performance.now()
  ) {}

  /**
   * Defines the workbook's function scripts, in order, and says what each
   * made callable or why it made nothing.
   */
  define(
    scripts: readonly { readonly name: string; readonly source: string }[]
  ): { readonly name: string; readonly names: readonly string[]; readonly problem: string }[] {
    const outcomes = this.functions.defineAll(scripts);
    const seen = new Map<string, string>();
    const clashing = new Set<string>();
    for (const outcome of outcomes) {
      for (const name of outcome.names) {
        const upper = name.toUpperCase();
        if (seen.has(upper) && seen.get(upper) !== name) {
          clashing.add(upper);
        }
        seen.set(upper, name);
      }
    }
    this.callable = new Map([...seen].filter(([upper]) => !clashing.has(upper) && !isSheetFunction(upper)));
    this.messages.clear();
    return outcomes.map(outcome => {
      if (outcome.error !== null) {
        return { name: outcome.name, names: [], problem: `It stopped with an error, and defines nothing: ${outcome.error}` };
      }
      const problems: string[] = [];
      const builtIn = outcome.names.filter(name => isSheetFunction(name.toUpperCase()));
      const twice = outcome.names.filter(name => clashing.has(name.toUpperCase()));
      if (builtIn.length > 0) {
        problems.push(`${builtIn.join(', ')} ${builtIn.length === 1 ? 'is' : 'are'} the sheet’s own, which a formula calls instead.`);
      }
      if (twice.length > 0) {
        problems.push(`${twice.join(', ')} ${twice.length === 1 ? 'differs' : 'differ'} from another function only by case, so a formula cannot tell which.`);
      }
      return {
        name: outcome.name,
        names: outcome.names.filter(name => this.callable.get(name.toUpperCase()) === name),
        problem: problems.join(' ')
      };
    });
  }

  has(name: string): boolean {
    return this.callable.has(name);
  }

  /** Why the call in this cell failed, or null when it did not. */
  messageAt(sheet: number, row: number, column: number): string | null {
    return this.messages.get(`${sheet}:${row}:${column}`) ?? null;
  }

  startSlice(): void {
    this.spent = 0;
  }

  overBudget(): boolean {
    return this.spent >= this.sliceMilliseconds;
  }

  call(
    name: string,
    args: readonly (CellValue | ArrayValue)[],
    at: { readonly sheet: number; readonly row: number; readonly column: number } | undefined
  ): CellValue | ArrayValue {
    const script = this.callable.get(name);
    if (script === undefined) {
      return NAME;
    }
    const given: FunctionValue[] = [];
    for (const arg of args) {
      const value = argumentOf(arg);
      if (isError(value)) {
        return value;
      }
      given.push(value);
    }
    const started = this.now();
    const result = this.functions.call(script, given);
    this.spent += this.now() - started;
    const key = at === undefined ? null : `${at.sheet}:${at.row}:${at.column}`;
    const [value, message] = cellValueOf(script, result);
    if (key !== null) {
      if (message === null) {
        this.messages.delete(key);
      } else {
        this.messages.set(key, message);
      }
    }
    return value;
  }
}

/** A cell value or an array as a function is handed it, or the error that stops the call. */
function argumentOf(arg: CellValue | ArrayValue): FunctionValue | CellError {
  if (!isArray(arg)) {
    return arg;
  }
  const rows: FunctionScalar[][] = [];
  for (let row = 0; row < arg.rows; row++) {
    const line: FunctionScalar[] = [];
    for (let column = 0; column < arg.columns; column++) {
      const value = arg.values[row * arg.columns + column] ?? null;
      if (isError(value)) {
        return value;
      }
      line.push(value);
    }
    rows.push(line);
  }
  return rows;
}

/** What a call's result puts in the cell, and why it failed if it did. */
function cellValueOf(name: string, result: FunctionResult): [CellValue | ArrayValue, string | null] {
  if (!result.ok) {
    switch (result.kind) {
      case 'thrown':
        return [VALUE, `${name} threw: ${result.message}`];
      case 'missing':
        return [NAME, result.message];
      default:
        return [CALC, result.message];
    }
  }
  const value = result.value;
  if (!Array.isArray(value)) {
    if (typeof value === 'number' && !Number.isFinite(value)) {
      return [NUM, `${name} returned ${String(value)}, which is not a number a cell can hold.`];
    }
    return [value, null];
  }
  if (value.length === 0) {
    return [CALC, `${name} returned no rows.`];
  }
  const columns = Math.max(...value.map(row => row.length));
  if (columns === 0) {
    return [CALC, `${name} returned rows with nothing in them.`];
  }
  const values: CellValue[] = [];
  for (const row of value) {
    for (let column = 0; column < columns; column++) {
      const cell = row[column] ?? null;
      values.push(typeof cell === 'number' && !Number.isFinite(cell) ? NUM : cell);
    }
  }
  return [arrayOrValue(value.length, columns, values), null];
}
