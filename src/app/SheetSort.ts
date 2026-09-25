import { compareValues, type CellValue } from '../sheet/Values';
import { rewriteFormula } from '../sheet/Rewrite';
import type { SheetDocument } from './SheetDocument';
import { readRect, type Rect } from './SheetRanges';

/**
 * Sorting a rectangle by one of its columns.
 *
 * **Whole rows move, not one column.** Sorting a column of names and
 * leaving the numbers beside them where they were is the single most
 * destructive thing a spreadsheet can do quietly, and it is what
 * sorting "just this column" means. So a sort takes a rectangle and
 * reorders the rows *of that rectangle*, and every cell in a row
 * travels with it.
 *
 * The ordering is `compareValues`, the same one `<` and `>` use in a
 * formula — numbers before text before booleans, text case-
 * insensitively — so a sort and a comparison never disagree about
 * which of two cells is larger. Empty cells go last whichever way the
 * sort runs, because a blank is an absence rather than a small value
 * and sorting it to the top buries the data under the gaps.
 */
export interface SortOrder {
  /** The column to sort by, absolute. */
  readonly column: number;
  readonly ascending: boolean;
  /**
   * Whether the rectangle's first row is a heading and stays put.
   *
   * Guessed by the caller rather than here: whoever is looking at the
   * sheet can see whether row one is a heading and this cannot.
   */
  readonly hasHeader: boolean;
}

/**
 * Reorders the rows of a rectangle, as one step on the undo stack.
 *
 * Formulas move with their rows and are rewritten by how far they
 * moved, which is `rewriteFormula` — the same machinery a fill uses,
 * for the same reason: a formula that said `=B2*C2` in row 2 has to
 * say `=B7*C7` when it lands in row 7, or it reads somebody else's
 * numbers. A reference *out* of the sorted block is left alone by the
 * `$` somebody wrote, exactly as it would be in a fill.
 */
export function sortRect(document: SheetDocument, rect: Rect, order: SortOrder): void {
  const first = rect.firstRow + (order.hasHeader ? 1 : 0);
  if (first >= rect.lastRow) {
    return;
  }
  const block = readRect(document, { ...rect, firstRow: first });
  /**
   * The formats travel with their rows.
   *
   * Left behind, a sorted table keeps its bold total row where the
   * total *was* and shows it against somebody else's numbers — which
   * is exactly what this did before a browser showed it. Palette ids
   * rather than formats, because an id is a number and the palette is
   * already shared.
   */
  const formats: number[][] = [];
  for (let row = first; row <= rect.lastRow; row++) {
    const line: number[] = [];
    for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
      line.push(document.formats.idAt(row, column));
    }
    formats.push(line);
  }
  const keys: CellValue[] = [];
  for (let row = first; row <= rect.lastRow; row++) {
    keys.push(document.sheet.value(row, order.column));
  }

  const order0 = keys.map((_, index) => index);
  order0.sort((a, b) => {
    const left = keys[a];
    const right = keys[b];
    // A blank sorts last either way: it is an absence, not a small
    // value, and sorting it to the top buries the data under the gaps.
    const leftEmpty = left === null;
    const rightEmpty = right === null;
    if (leftEmpty !== rightEmpty) {
      return leftEmpty ? 1 : -1;
    }
    if (leftEmpty) {
      return a - b;
    }
    const compared = compareValues(left, right);
    // An error compares as nothing; keep those rows where they were
    // rather than inventing an order for them.
    const by = typeof compared === 'number' ? compared : 0;
    // Ties keep the order they had, so sorting by one column and then
    // another gives the second inside the first — which is how people
    // sort by two columns without a dialog for it.
    return by === 0 ? a - b : order.ascending ? by : -by;
  });

  if (order0.every((from, to) => from === to)) {
    return;
  }

  document.transact(() => {
    order0.forEach((from, to) => {
      const line = block[from];
      const rowDelta = to - from;
      line.forEach((input, offset) => {
        const column = rect.firstColumn + offset;
        document.setFormat(first + to, column, document.formats.byId(formats[from][offset]));
        document.setCell(first + to, column, rowDelta === 0 ? input : rewriteFormula(input, rowDelta, 0));
      });
    });
  });
}
