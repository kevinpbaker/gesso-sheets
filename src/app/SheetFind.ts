import { cellKey, columnOf, rowOf } from '../sheet/A1';

/**
 * What a search needs, which is less than a document.
 *
 * `display` and not `Sheet.display`: searching values has to find
 * what is *on the screen*, and since Phase 9 that depends on the
 * cell's number format — somebody who can see `$1,234.50` and
 * searches for `1,234` has to find it.
 */
export interface Searchable {
  entries(): Generator<{ row: number; column: number; input: string }>;
  display(row: number, column: number): string;
}

/**
 * Find and replace, on the application worker.
 *
 * It has to be here, and that is worth stating rather than assuming:
 * the render worker knows the thirty rows it has mounted and nothing
 * else, so a find implemented on the render side could only ever
 * search what somebody was already looking at. Every cell in the
 * sheet is on this side of the barrier, so this is the side that can
 * answer. The same sentence rules out `FindBar` and `FindService`
 * from the component set, which search the semantics tree — the right
 * answer for a page, and the wrong one for a document whose cells are
 * mostly not on screen.
 */

export interface FindOptions {
  readonly matchCase: boolean;
  /** The whole cell must equal the query, rather than contain it. */
  readonly wholeCell: boolean;
  /**
   * Search what was typed rather than what is shown.
   *
   * Excel calls these Formulas and Values, and the distinction is
   * real: `=1+2` is found by searching for `1+2` in formulas and by
   * searching for `3` in values, and neither search finds the other.
   */
  readonly inFormulas: boolean;
}

export const DEFAULT_FIND: FindOptions = { matchCase: false, wholeCell: false, inFormulas: true };

/**
 * Every cell matching a query, in reading order.
 *
 * Walks the store rather than the sheet: a million-cell sheet holding
 * forty values costs forty, and the alternative — walking rows and
 * columns — costs a million on a sheet where the answer is nowhere.
 * Sorting afterwards is what puts them back into the order a person
 * reads them in, which is the order Enter walks them in.
 */
export function findMatches(
  sheet: Searchable,
  query: string,
  options: FindOptions,
  rowCount: number,
  columnCount: number
): number[] {
  if (query === '') {
    return [];
  }
  const found: number[] = [];
  for (const cell of sheet.entries()) {
    if (cell.row >= rowCount || cell.column >= columnCount) {
      continue;
    }
    const text = options.inFormulas ? cell.input : sheet.display(cell.row, cell.column);
    if (matches(text, query, options)) {
      found.push(cellKey(cell.row, cell.column));
    }
  }
  // A `Map` has insertion order, not reading order, and a find that
  // jumped around the sheet in the order cells happened to be typed
  // would be unusable even though every match was correct.
  return found.sort((a, b) => a - b);
}

export function matches(text: string, query: string, options: FindOptions): boolean {
  const haystack = options.matchCase ? text : text.toLowerCase();
  const needle = options.matchCase ? query : query.toLowerCase();
  return options.wholeCell ? haystack === needle : haystack.includes(needle);
}

/**
 * The match at or after a cell, wrapping to the top.
 *
 * `after` decides whether the cell the selection is already on counts:
 * opening the find bar should offer the cell you are standing on, and
 * pressing Enter again should move off it.
 */
export function stepTo(found: readonly number[], from: number, after: boolean): number {
  if (found.length === 0) {
    return -1;
  }
  const index = found.findIndex(key => (after ? key > from : key >= from));
  return index === -1 ? found[0] : found[index];
}

/** The match before a cell, wrapping to the bottom. */
export function stepBack(found: readonly number[], from: number): number {
  if (found.length === 0) {
    return -1;
  }
  for (let index = found.length - 1; index >= 0; index--) {
    if (found[index] < from) {
      return found[index];
    }
  }
  return found[found.length - 1];
}

/**
 * One cell's text with the query replaced.
 *
 * Every occurrence inside the cell, not just the first: replacing
 * `a` with `b` in `banana` gives `bbnbnb`, which is what every other
 * spreadsheet and every text editor does. A whole-cell match replaces
 * the lot.
 *
 * Done without `RegExp` on purpose. The query is somebody's typing,
 * and `(` or `*` in it would either throw or match something they did
 * not ask for — a find for `SUM(` has to find `SUM(`.
 */
export function replaceIn(text: string, query: string, replacement: string, options: FindOptions): string {
  if (query === '') {
    return text;
  }
  if (options.wholeCell) {
    return matches(text, query, options) ? replacement : text;
  }
  if (options.matchCase) {
    return text.split(query).join(replacement);
  }
  const lower = text.toLowerCase();
  const needle = query.toLowerCase();
  let out = '';
  let at = 0;
  for (;;) {
    const found = lower.indexOf(needle, at);
    if (found === -1) {
      out += text.slice(at);
      return out;
    }
    out += text.slice(at, found) + replacement;
    at = found + needle.length;
  }
}

/** A found key as the row and column it names. */
export function at(key: number): { row: number; column: number } {
  return { row: rowOf(key), column: columnOf(key) };
}
