import { cellKey, columnName, rowOf, columnOf } from './A1';
import { isError, type CellValue, type ErrorCode } from './Values';

/**
 * Why a cell is showing an error.
 *
 * An error code is a diagnosis in five characters, and five
 * characters is not enough: `#DIV/0!` in a cell that divides nothing
 * means the division is three cells away, and the person looking at
 * it has to walk the formulas backwards to find out where. That walk
 * is mechanical, which is the argument for the machine doing it.
 *
 * Two questions, and the second is the useful one. *What does the
 * code mean* is a lookup. *Where did it come from* is a search
 * upstream for the cell that produced it rather than merely passed it
 * on — the one whose own precedents are all fine. That is the cell
 * somebody has to go and fix.
 */

export interface Explanation {
  readonly code: ErrorCode;
  /** One line, in the terms of the sheet rather than of the engine. */
  readonly meaning: string;
  /**
   * The cell that produced it, when it is not this one.
   *
   * Null when the cell broke on its own, which is the common case and
   * needs no pointing at.
   */
  readonly blame: { readonly row: number; readonly column: number } | null;
}

const MEANINGS: Readonly<Record<ErrorCode, string>> = {
  '#DIV/0!': 'Something was divided by zero, or by an empty cell.',
  '#VALUE!': 'A value of the wrong kind: text where a number was needed, or a formula that does not parse.',
  '#REF!': 'A reference points at no cell. A row or column it named was deleted.',
  '#NAME?': 'A name the sheet does not know. Usually a misspelt function.',
  '#N/A': 'A lookup found nothing. The formula is fine; the table has no such row.',
  '#CIRC!': 'This cell depends on itself, directly or through others. There is no order to work it out in.'
};

/** What a sheet has to be able to answer for this to work. */
export interface ExplainSource {
  valueAt(key: number): CellValue;
  precedentsOf(row: number, column: number): number[];
}

/**
 * The explanation for a cell, or null when it is not showing an error.
 *
 * The search upstream is breadth-first and stops at the first cell
 * carrying this code whose own precedents carry none — the origin.
 * Bounded, because a sheet can be large and this runs whenever the
 * selection moves: past a few hundred cells the answer stops being
 * worth the walk, and "somewhere upstream" is what it would say
 * anyway.
 */
export function explainCell(sheet: ExplainSource, row: number, column: number, limit = 500): Explanation | null {
  const value = sheet.valueAt(cellKey(row, column));
  if (!isError(value)) {
    return null;
  }
  const code = value.code;
  return { code, meaning: MEANINGS[code], blame: originOf(sheet, row, column, code, limit) };
}

function originOf(
  sheet: ExplainSource,
  row: number,
  column: number,
  code: ErrorCode,
  limit: number
): { row: number; column: number } | null {
  const seen = new Set<number>([cellKey(row, column)]);
  const queue: { row: number; column: number }[] = [{ row, column }];

  for (let at = 0; at < queue.length && seen.size <= limit; at++) {
    const here = queue[at];
    const precedents = sheet.precedentsOf(here.row, here.column);
    // A cell carrying the code whose precedents carry none produced
    // it. Checked before the walk goes further so the *nearest* such
    // cell wins, which is the one somebody can act on.
    let carriedIn = false;
    for (const key of precedents) {
      const value = sheet.valueAt(key);
      if (isError(value) && value.code === code) {
        carriedIn = true;
        if (!seen.has(key)) {
          seen.add(key);
          queue.push({ row: rowOf(key), column: columnOf(key) });
        }
      }
    }
    if (!carriedIn) {
      // `here` is where it started. Reported only when it is not the
      // cell being asked about: pointing somebody at the cell they are
      // already looking at says nothing.
      return here.row === row && here.column === column ? null : here;
    }
  }
  return null;
}

/** `B7`, for putting in a sentence. */
export function addressOf(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`;
}
