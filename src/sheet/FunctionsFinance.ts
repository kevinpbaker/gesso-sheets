import { checked, numberAt, numbersOf, type Argument, type SheetFunction } from './FunctionKit';
import { DIV0, isError, NUM, VALUE, type CellError } from './Values';

/**
 * Loans and savings: the time value of money.
 *
 * The four a budget reaches for. Signs follow the spreadsheet
 * convention — money paid out is negative — so a loan of 10,000 has a
 * negative payment, which is what makes `PMT` and `PV` each other's
 * inverse without a minus sign in the middle. The formulas are the
 * standard annuity ones, and the cases in `Functions.spec.ts` are
 * Microsoft's own documented examples, to the last digit shown.
 */

/** The optional trailing number, or its default. */
function optional(args: readonly Argument[], index: number, fallback: number): number | CellError {
  if (index >= args.length) {
    return fallback;
  }
  const arg = args[index];
  if (arg.kind === 'value' && arg.value === null) {
    return fallback;
  }
  return numberAt(args, index);
}

/** Rate, periods, and the three optional terms, read once for all four. */
function terms(args: readonly Argument[], min: number): { rate: number; periods: number; third: number; fourth: number; due: number } | CellError {
  const wrong = checked(args, min, 5);
  if (wrong !== null) {
    return wrong;
  }
  const rate = numberAt(args, 0);
  if (isError(rate)) {
    return rate;
  }
  const periods = numberAt(args, 1);
  if (isError(periods)) {
    return periods;
  }
  const third = numberAt(args, 2);
  if (isError(third)) {
    return third;
  }
  const fourth = optional(args, 3, 0);
  if (isError(fourth)) {
    return fourth;
  }
  const due = optional(args, 4, 0);
  if (isError(due)) {
    return due;
  }
  return { rate, periods, third, fourth, due: due === 0 ? 0 : 1 };
}

export const FINANCE_FUNCTIONS: Readonly<Record<string, SheetFunction>> = {
  /** The payment each period that pays off `pv` (and leaves `fv`) over `nper` periods. */
  PMT(args) {
    const read = terms(args, 3);
    if (isError(read)) {
      return read;
    }
    const { rate, periods, third: present, fourth: future, due } = read;
    if (periods === 0) {
      return NUM;
    }
    if (rate === 0) {
      return -(present + future) / periods;
    }
    const growth = (1 + rate) ** periods;
    return -(rate * (present * growth + future)) / ((1 + rate * due) * (growth - 1));
  },

  /** What a series of payments, and a sum put in now, will be worth at the end. */
  FV(args) {
    const read = terms(args, 3);
    if (isError(read)) {
      return read;
    }
    const { rate, periods, third: payment, fourth: present, due } = read;
    if (rate === 0) {
      return -(present + payment * periods);
    }
    const growth = (1 + rate) ** periods;
    return -(present * growth + (payment * (1 + rate * due) * (growth - 1)) / rate);
  },

  /** What a series of payments, and a sum at the end, are worth now. */
  PV(args) {
    const read = terms(args, 3);
    if (isError(read)) {
      return read;
    }
    const { rate, periods, third: payment, fourth: future, due } = read;
    if (rate === 0) {
      return -(future + payment * periods);
    }
    const growth = (1 + rate) ** periods;
    return -(future + (payment * (1 + rate * due) * (growth - 1)) / rate) / growth;
  },

  /**
   * Cash flows at the end of each period, discounted to now. The first
   * flow is one period away, as Excel has it — a flow made today is
   * added outside the function, which is the usual mistake and the
   * usual fix.
   */
  NPV(args) {
    const wrong = checked(args, 2, 255);
    if (wrong !== null) {
      return wrong;
    }
    const rate = numberAt(args, 0);
    if (isError(rate)) {
      return rate;
    }
    if (rate === -1) {
      return DIV0;
    }
    const flows = numbersOf(args.slice(1));
    if (isError(flows)) {
      return flows;
    }
    return flows.reduce((total, flow, at) => total + flow / (1 + rate) ** (at + 1), 0);
  }
};
