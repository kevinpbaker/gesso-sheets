import type { RangeRef } from './A1';
import { parseTypedDate, serialOfDate } from './Dates';
import { isError, type CellValue } from './Values';

/**
 * What a cell is allowed to hold.
 *
 * The same shape as a conditional format and resolved the same way:
 * a rule over a range, asked about the cells in the window at publish
 * time. What it adds is a second place it is asked — at **commit**,
 * where a rule can refuse the write rather than only marking it.
 *
 * Marking and refusing are both here because they answer different
 * questions. A list of regions is a rule somebody wants enforced:
 * typing `Noth` should not be allowed to stand. A range of plausible
 * temperatures is a rule somebody wants *noticed*: the reading is
 * what it is, and a sheet that refused to record it would be a sheet
 * that lies. So a rule says which it is and the document does what it
 * says.
 */

export type ValidationRule =
  /**
   * One of these, and nothing else.
   *
   * The only kind with a dropdown, because it is the only kind where
   * the set of acceptable values is small and known — which is also
   * what makes it the kind worth enforcing.
   */
  | { readonly kind: 'list'; readonly values: readonly string[] }
  | { readonly kind: 'number'; readonly min?: number; readonly max?: number; readonly integer?: boolean }
  | { readonly kind: 'text'; readonly maxLength?: number }
  | { readonly kind: 'date'; readonly from?: number; readonly to?: number };

export interface Validation {
  readonly range: RangeRef;
  readonly rule: ValidationRule;
  /**
   * Whether a value that fails is refused or only marked.
   *
   * False by default, which is the safer way round: a rule that
   * silently refuses what somebody typed is worse than one that
   * shows a mark beside it, because the mark can be read and the
   * refusal looks like a broken keyboard.
   */
  readonly strict?: boolean;
  /** What to say when it fails, in words somebody can act on. */
  readonly message?: string;
}

/**
 * Why a value is not allowed, or null when it is.
 *
 * An empty cell is always allowed. A rule says what a value must be,
 * not that there has to be one — emptying a cell is how somebody
 * takes back a mistake, and a validation that refused it would be a
 * cell nobody could clear.
 */
export function validate(rule: ValidationRule, value: CellValue): string | null {
  if (value === null || value === '') {
    return null;
  }
  if (isError(value)) {
    // A formula that is broken is a different problem, and the cell
    // is already saying so.
    return null;
  }
  switch (rule.kind) {
    case 'list': {
      const text = String(value).trim().toUpperCase();
      const allowed = rule.values.some(entry => entry.trim().toUpperCase() === text);
      return allowed ? null : `Not one of: ${rule.values.join(', ')}`;
    }
    case 'number': {
      if (typeof value !== 'number') {
        return 'A number is wanted here.';
      }
      if (rule.integer === true && !Number.isInteger(value)) {
        return 'A whole number is wanted here.';
      }
      if (rule.min !== undefined && value < rule.min) {
        return `Not below ${rule.min}.`;
      }
      if (rule.max !== undefined && value > rule.max) {
        return `Not above ${rule.max}.`;
      }
      return null;
    }
    case 'text': {
      const text = String(value);
      if (rule.maxLength !== undefined && text.length > rule.maxLength) {
        return `At most ${rule.maxLength} characters.`;
      }
      return null;
    }
    case 'date': {
      const serial = serialOf(value);
      if (serial === null) {
        return 'A date is wanted here.';
      }
      if (rule.from !== undefined && serial < rule.from) {
        return 'Earlier than the range allows.';
      }
      if (rule.to !== undefined && serial > rule.to) {
        return 'Later than the range allows.';
      }
      return null;
    }
  }
}

/**
 * A value as a date serial, or null.
 *
 * A number is already one — that is what a date *is* in a
 * spreadsheet — and text is given the same reading somebody typing
 * it would get.
 */
function serialOf(value: CellValue): number | null {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value !== 'string') {
    return null;
  }
  const typed = parseTypedDate(value);
  return typed?.serial ?? null;
}

/** Today, as a serial, for a rule written against it. */
export function todaySerial(now = new Date()): number {
  return serialOfDate(now.getFullYear(), now.getMonth() + 1, now.getDate());
}
