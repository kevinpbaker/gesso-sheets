import { compareValues, type CellValue } from '../sheet/Values';
import { permuteFormula, rowsRead, type ColumnSpan } from '../sheet/Permute';
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
 * Formulas move with their rows, and their references follow the rows
 * *the sort moved* rather than the distance the formula travelled —
 * `permuteFormula`, not `rewriteFormula`. A formula that said `=B2*C2`
 * in row 2 still has to say `=B7*C7` when it lands in row 7, and it
 * does, because rows 2 and 7 are what swapped. But a total that said
 * `=SUM(B2:B4)` keeps saying it wherever it lands, instead of being
 * dragged off the top of the sheet. See `Permute.ts`, which is mostly
 * an explanation of why these are two different operations.
 */
export function sortRect(document: SheetDocument, rect: Rect, order: SortOrder): void {
  const first = rect.firstRow + (order.hasHeader ? 1 : 0);
  if (first >= rect.lastRow) {
    return;
  }
  const columns = { first: rect.firstColumn, last: rect.lastColumn };
  const last = dataFloor(document, first, rect.lastRow, columns);
  if (first >= last) {
    return;
  }
  const block = readRect(document, { ...rect, firstRow: first, lastRow: last });
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
  for (let row = first; row <= last; row++) {
    const line: number[] = [];
    for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
      line.push(document.formats.idAt(row, column));
    }
    formats.push(line);
  }
  const keys: CellValue[] = [];
  for (let row = first; row <= last; row++) {
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

  /**
   * Where each row the sort touched ended up, in sheet rows.
   *
   * Built before anything is written, because every formula is
   * rewritten against the *whole* permutation: a row that did not move
   * may still hold a formula naming a row that did.
   */
  const moved = new Map<number, number>();
  order0.forEach((from, to) => moved.set(first + from, first + to));

  document.transact(() => {
    order0.forEach((from, to) => {
      const line = block[from];
      line.forEach((input, offset) => {
        const column = rect.firstColumn + offset;
        document.setFormat(first + to, column, document.formats.byId(formats[from][offset]));
        document.setCell(first + to, column, permuteFormula(input, moved, columns));
      });
    });
  });
}

/**
 * The last row of a block that is actually data.
 *
 * A table's bottom row is very often a total, and a total is not one of
 * the things being sorted — it is a statement *about* the things being
 * sorted. Sorting it in with them is what wrecked the demo sheet twice:
 * descending by region carried `=SUM(B2:B4)` into the middle of the
 * rows it adds up, which is a circular reference, and no rule about
 * rewriting references can rescue that. The row had to stay put.
 *
 * What marks the row is a *range* over the block — `SUM(B2:B6)` — and
 * not merely reading some other row. The narrowness is the whole
 * design, and a browser is what taught it: every data row of the demo
 * sheet holds `=D2/$D$7*100`, a share of the total, so "reads another
 * row of the block" described all five of them and pinned the entire
 * table. The sort silently did nothing, which is its own kind of bad.
 *
 * A range is the right mark because a range is the thing that cannot
 * follow a permutation. A single reference can: `permuteFormula` sends
 * it wherever its cell went, so a row full of them sorts correctly
 * whatever it points at, including the total below it. A range has to
 * be left where it is — see `Permute.ts` — and leaving it where it is
 * only stays true if the row holding it also stays.
 *
 * Only trailing rows, and only within the sorted columns. A summary in
 * the middle of a table is not a thing people write, and pinning
 * interior rows would silently turn one sort into two.
 */
function dataFloor(document: SheetDocument, first: number, lastRow: number, columns: ColumnSpan): number {
  let last = lastRow;
  while (last > first && aggregatesTheBlock(document, last, first, lastRow, columns)) {
    last--;
  }
  return last;
}

/** Whether a row reads a run of the block's rows, rather than cells of it. */
function aggregatesTheBlock(
  document: SheetDocument,
  row: number,
  first: number,
  lastRow: number,
  columns: ColumnSpan
): boolean {
  for (let column = columns.first; column <= columns.last; column++) {
    for (const span of rowsRead(document.sheet.input(row, column), columns)) {
      if (span.first !== span.last && span.last >= first && span.first <= lastRow) {
        return true;
      }
    }
  }
  return false;
}
