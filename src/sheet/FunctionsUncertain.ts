import { arity, numberAt, type FunctionContext, type SheetFunction } from './FunctionKit';
import { isError, NUM, type CellError } from './Values';

/**
 * Guesses, written as a guess.
 *
 * `=NORMAL(0.02, 0.01)` says the growth is about two per cent, give or
 * take one. Everywhere but a simulation it *is* two per cent — its
 * likeliest value — so the workbook it sits in shows, saves and
 * recalculates exactly as it would with `0.02` typed there. A
 * simulation recalculates the workbook thousands of times with each one
 * drawn from its spread, and what every cell downstream came to in each
 * trial is the answer to "how sure are we of that".
 *
 * Three shapes, the ones a person can state without a statistics
 * course: a bell with a centre and a spread, a range where any value is
 * as likely as any other, and a low, a likeliest and a high.
 */
export const UNCERTAIN_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** A bell: likeliest at `mean`, two thirds of draws within `sd` of it. */
  NORMAL(args, ctx) {
    const wrong = arity(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersFrom(args, 2);
    if (!Array.isArray(numbers)) {
      return numbers;
    }
    const [mean, sd] = numbers;
    if (sd < 0) {
      return NUM;
    }
    const z = standardNormal(ctx);
    return z === null ? mean : mean + sd * z;
  },

  /** Any value from `low` to `high`, each as likely; the middle when not drawing. */
  UNIFORM(args, ctx) {
    const wrong = arity(args, 2, 2);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersFrom(args, 2);
    if (!Array.isArray(numbers)) {
      return numbers;
    }
    const [low, high] = numbers;
    if (high < low) {
      return NUM;
    }
    const u = ctx.sample();
    return u === null ? (low + high) / 2 : low + (high - low) * u;
  },

  /**
   * A low, a likeliest and a high: the estimate people give when asked
   * for one. Drawn from the triangle with those three corners.
   */
  TRIANGULAR(args, ctx) {
    const wrong = arity(args, 3, 3);
    if (wrong !== null) {
      return wrong;
    }
    const numbers = numbersFrom(args, 3);
    if (!Array.isArray(numbers)) {
      return numbers;
    }
    const [low, likeliest, high] = numbers;
    if (!(low <= likeliest && likeliest <= high)) {
      return NUM;
    }
    const u = ctx.sample();
    if (u === null || high === low) {
      return likeliest;
    }
    // The inverse of the triangle's distribution: the left side's
    // area is the share of draws below the peak.
    const split = (likeliest - low) / (high - low);
    return u < split
      ? low + Math.sqrt(u * (high - low) * (likeliest - low))
      : high - Math.sqrt((1 - u) * (high - low) * (high - likeliest));
  }
};

function numbersFrom(args: Parameters<SheetFunction>[0], count: number): number[] | CellError {
  const numbers: number[] = [];
  for (let index = 0; index < count; index++) {
    const value = numberAt(args, index);
    if (isError(value)) {
      return value;
    }
    numbers.push(value);
  }
  return numbers;
}

/**
 * A draw from the standard bell, by Box and Muller's transform of two
 * uniform draws, or null when the context is not drawing. The first
 * uniform is kept away from zero, whose logarithm is not a number.
 */
function standardNormal(ctx: FunctionContext): number | null {
  const u = ctx.sample();
  if (u === null) {
    return null;
  }
  const v = ctx.sample() ?? 0.5;
  return Math.sqrt(-2 * Math.log(Math.max(u, Number.MIN_VALUE))) * Math.cos(2 * Math.PI * v);
}
