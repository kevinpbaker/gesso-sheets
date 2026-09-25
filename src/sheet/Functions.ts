import {
  checked,
  firstError,
  numberAt,
  numbersOf,
  scalar,
  valuesOf,
  type Argument,
  type FunctionContext,
  type SheetFunction
} from './FunctionKit';
import { CONDITIONAL_FUNCTIONS } from './FunctionsConditional';
import { DATE_FUNCTIONS } from './FunctionsDate';
import { LOGIC_FUNCTIONS } from './FunctionsLogic';
import { LOOKUP_FUNCTIONS, } from './FunctionsLookup';
import { MATH_FUNCTIONS, roundHalfAway } from './FunctionsMath';
import { STATS_FUNCTIONS } from './FunctionsStats';
import { TEXT_FUNCTIONS } from './FunctionsText';
import { DIV0, isError, toText, type CellValue } from './Values';

export type { Argument, FunctionContext, SheetFunction };

/**
 * The library: every function the sheet knows, in one table.
 *
 * The families are separate files because sixty functions in one is
 * unreadable, and the registry is this one because "what does this
 * sheet know" has to have a single answer. `Functions.spec.ts` is the
 * conformance table over it, and it fails if a name appears here with
 * no case asserting what it does.
 *
 * ## What is not in here
 *
 * Five functions are in `Evaluator.ts` instead, and for two different
 * reasons.
 *
 * **`IF`, `IFS` and `SWITCH` must not evaluate everything.** A
 * function here is handed arguments that are already values, and
 * those three exist specifically to leave one unevaluated:
 * `=IF(B1=0, "n/a", A1/B1)` is how everybody writes a guarded
 * division, and a sheet that evaluated both branches would answer
 * `#DIV/0!` to the formula written to avoid it.
 *
 * **`INDIRECT` and `OFFSET` compute a reference.** They need to
 * *read* cells chosen while they run, which a value-in value-out
 * function cannot do, and they make the dependency graph unable to
 * know their edges beforehand. That is the most dangerous thing in
 * this phase — see `Sheet.evaluateCell`.
 */
const AGGREGATES: Readonly<Record<string, SheetFunction>> = {
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
    for (const value of valuesOf(args)) {
      if (typeof value === 'number') {
        count++;
      }
    }
    return count;
  },

  ROUND(args) {
    const wrong = checked(args, 1, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    if (isError(value)) {
      return value;
    }
    const places = args.length > 1 ? numberAt(args, 1) : 0;
    return isError(places) ? places : roundHalfAway(value, places);
  },

  ABS(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const value = numberAt(args, 0);
    return isError(value) ? value : Math.abs(value);
  },

  CONCAT(args) {
    const error = firstError(args);
    if (error !== null) {
      return error;
    }
    let text = '';
    for (const value of valuesOf(args)) {
      const piece = toText(value);
      if (isError(piece)) {
        return piece;
      }
      text += piece;
    }
    return text;
  }
};

export const FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  ...AGGREGATES,
  ...LOGIC_FUNCTIONS,
  ...MATH_FUNCTIONS,
  ...STATS_FUNCTIONS,
  ...CONDITIONAL_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...DATE_FUNCTIONS
};

/**
 * The functions whose answer changes when nothing they read has.
 *
 * A formula containing one of these is recalculated on every edit
 * anywhere, because there is no edge in the dependency graph that
 * would ever wake it: `=TODAY()` reads nothing, so nothing marks it
 * dirty, and without this it would show the day the file was opened
 * until somebody retyped it.
 *
 * Excel calls this set volatile and treats it the same way. The cost
 * is real and the alternative is a sheet that lies about the date, so
 * the set is kept as small as it can be: these four and nothing else.
 * `INDIRECT` and `OFFSET` are deliberately *not* here — they are
 * handled by re-deriving their edges after each evaluation, which is
 * exact where volatility would be a blunt instrument.
 */
export const VOLATILE: ReadonlySet<string> = new Set(['RAND', 'RANDBETWEEN', 'NOW', 'TODAY']);

/** The names handled in the evaluator rather than by the table. */
export const SPECIAL_FORMS: ReadonlySet<string> = new Set(['IF', 'IFS', 'SWITCH', 'INDIRECT', 'OFFSET']);

/** Whether the sheet knows a name at all, however it is implemented. */
export function isSheetFunction(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(FUNCTIONS, name) || SPECIAL_FORMS.has(name);
}

/** Every name the sheet knows, for the conformance table to walk. */
export function functionNames(): string[] {
  return [...Object.keys(FUNCTIONS), ...SPECIAL_FORMS].sort();
}

/** The default context: the real clock and real dice. */
export function liveContext(): FunctionContext {
  const serial = nowSerial();
  return { now: () => serial, random: () => Math.random() };
}

/**
 * The moment, as a serial.
 *
 * Local rather than UTC, deliberately and unlike everything else in
 * `Dates.ts`: `TODAY()` has to be the date on the wall of the person
 * looking at the screen. A UTC `TODAY()` is yesterday all evening in
 * Auckland, which is the one thing it must never be. Stored dates
 * stay UTC — they are the same day for everybody — and only the
 * reading of *now* is local, because only *now* is a question about
 * where you are standing.
 */
export function nowSerial(): number {
  const at = new Date();
  const midnight = new Date(at.getFullYear(), at.getMonth(), at.getDate()).getTime();
  const days = Math.round((Date.UTC(at.getFullYear(), at.getMonth(), at.getDate()) - Date.UTC(1899, 11, 30)) / 86_400_000);
  return days + (at.getTime() - midnight) / 86_400_000;
}

/** Re-exported so the evaluator can build arguments without a cycle. */
export { scalar };
export type { CellValue };
