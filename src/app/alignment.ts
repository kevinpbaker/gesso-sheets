/**
 * Where an `auto`-aligned cell sits, as the grid works it out from the
 * string and as the value says it should: numbers on the right,
 * logicals and errors in the middle, everything else on the left —
 * which is Excel's rule, and every other spreadsheet's.
 *
 * Shared by both workers, and that is the point of it being a module.
 * The grid decides from the string because the string is all it has;
 * the application worker has the value too, and asks `placeOf` and
 * `guessOf` the same question to find the cells where the string
 * misleads — `$4.50`, `9.6%` and `2026-09-24`, which are numbers, and a
 * Text cell holding `007` or `TRUE`, which are text — so it can send an
 * alignment for those and for nothing else. Two copies of the rule
 * would be two answers to what the grid will do.
 */

export type Place = 'end' | 'center' | 'start';

const LOGICAL_OR_ERROR = /^(TRUE|FALSE|#(REF!|DIV\/0!|NAME\?|VALUE!|CIRC!|N\/A|NUM!))$/;

/** Where the grid puts a display string it knows nothing else about. */
export function guessOf(text: string | null): Place {
  if (text === null || text === '') {
    return 'start';
  }
  if (LOGICAL_OR_ERROR.test(text)) {
    return 'center';
  }
  return Number.isNaN(Number(text.replace(/,/g, ''))) ? 'start' : 'end';
}

/** Where a value belongs: numbers right, logicals and errors centred, text left. */
export function placeOf(value: unknown): Place {
  if (typeof value === 'number') {
    return 'end';
  }
  if (typeof value === 'boolean' || (typeof value === 'object' && value !== null && 'kind' in value)) {
    return 'center';
  }
  return 'start';
}
