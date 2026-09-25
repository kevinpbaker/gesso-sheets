import { arity, checked, scalar, type Argument, type SheetFunction } from './FunctionKit';
import { isError, NA, toBoolean, VALUE, type CellError, type CellValue } from './Values';

/**
 * Logic, minus the three that have to be lazy.
 *
 * `IF`, `IFS` and `SWITCH` are in the evaluator, because a function
 * here is handed arguments that have already been evaluated and those
 * three exist precisely to *not* evaluate one of theirs. See
 * `Evaluator.call`.
 *
 * `IFERROR` and `IFNA` are here and not there, which looks
 * inconsistent and is not. They take an argument that is *expected*
 * to be an error and hand back the other one; evaluating both costs
 * nothing, because the fallback of an `IFERROR` whose first argument
 * is fine is simply discarded. Laziness would buy no correctness, and
 * the rule is that only the functions that need it get it.
 */

/**
 * The booleans a conjunction should see.
 *
 * Text and blanks inside a *range* are skipped, as they are in `SUM`
 * and for the same reason: `AND(A1:A100)` over a column with a
 * heading must not be false because of the heading. A scalar is
 * coerced, so `AND(TRUE, "x")` is `#VALUE!` — passing something
 * directly is a claim that it is a truth value.
 */
function booleansOf(args: readonly Argument[]): boolean[] | CellError {
  const flags: boolean[] = [];
  for (const arg of args) {
    if (arg.kind === 'range') {
      for (const value of arg.values) {
        if (typeof value === 'boolean') {
          flags.push(value);
        } else if (typeof value === 'number') {
          flags.push(value !== 0);
        }
      }
      continue;
    }
    if (arg.value === null) {
      continue;
    }
    const flag = toBoolean(arg.value);
    if (isError(flag)) {
      return flag;
    }
    flags.push(flag);
  }
  return flags;
}

export const LOGIC_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  AND(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const flags = booleansOf(args);
    if (isError(flags)) {
      return flags;
    }
    // Nothing to test is `#VALUE!` rather than the vacuous truth: an
    // `AND` over a range that turned out to hold no truth values is a
    // formula pointing somewhere unexpected, and TRUE would hide it.
    return flags.length === 0 ? VALUE : flags.every(flag => flag);
  },

  OR(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const flags = booleansOf(args);
    if (isError(flags)) {
      return flags;
    }
    return flags.length === 0 ? VALUE : flags.some(flag => flag);
  },

  /** True when an odd number of them are, which is what XOR means. */
  XOR(args) {
    const wrong = checked(args, 1, Number.POSITIVE_INFINITY);
    if (wrong !== null) {
      return wrong;
    }
    const flags = booleansOf(args);
    if (isError(flags)) {
      return flags;
    }
    return flags.length === 0 ? VALUE : flags.filter(flag => flag).length % 2 === 1;
  },

  NOT(args) {
    const wrong = checked(args, 1, 1);
    if (wrong !== null) {
      return wrong;
    }
    const flag = toBoolean(scalar(args, 0));
    return isError(flag) ? flag : !flag;
  },

  /**
   * The first argument unless it broke, in which case the second.
   *
   * Deliberately catches every error including `#CIRC!`, which is the
   * one people are surprised by. Excel does the same, and a version
   * that let circularity through would be a special case somebody has
   * to know about.
   */
  IFERROR(args) {
    const wrong = arity(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = scalar(args, 0);
    return isError(value) ? scalar(args, 1) : value;
  },

  /** The same, for the one error that means "not found". */
  IFNA(args) {
    const wrong = arity(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const value = scalar(args, 0);
    return isError(value) && value.code === '#N/A' ? scalar(args, 1) : value;
  },

  /**
   * Whether a value is one, for the cases `IFERROR` cannot express.
   *
   * Not in the roadmap's list, and here because `IFERROR` is not a
   * substitute for it: a cell that wants to *count* its broken
   * neighbours rather than replace them has no other way to ask.
   */
  ISERROR(args) {
    const wrong = arity(args, 1, 1);
    return wrong !== null ? wrong : isError(scalar(args, 0));
  },

  ISBLANK(args) {
    const wrong = arity(args, 1, 1);
    return wrong !== null ? wrong : scalar(args, 0) === null;
  },

  ISNUMBER(args) {
    const wrong = arity(args, 1, 1);
    return wrong !== null ? wrong : typeof scalar(args, 0) === 'number';
  },

  ISTEXT(args) {
    const wrong = arity(args, 1, 1);
    return wrong !== null ? wrong : typeof scalar(args, 0) === 'string';
  },

  /** The count of values, for a `TRUE()` that reads as a function. */
  TRUE(args) {
    return arity(args, 0, 0) ?? true;
  },

  FALSE(args) {
    return arity(args, 0, 0) ?? false;
  },

  /**
   * `#N/A`, written deliberately.
   *
   * The way a half-built sheet says "this one is not filled in yet"
   * so that the totals below it say so too rather than quietly
   * treating the gap as zero.
   */
  NA(args) {
    return arity(args, 0, 0) ?? NA;
  }
};
