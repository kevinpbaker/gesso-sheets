import { rewriteFormula } from '../sheet/Rewrite';
import { fromTsv, toTsv, type Block } from '../sheet/Tsv';
import type { SheetDocument } from './SheetDocument';
import type { SheetSelection } from './SheetContract';

/** A selection as the rectangle it covers, corners normalised. */
export interface Rect {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

export function rectOf(selection: SheetSelection): Rect {
  return {
    firstRow: Math.min(selection.row, selection.anchorRow),
    lastRow: Math.max(selection.row, selection.anchorRow),
    firstColumn: Math.min(selection.column, selection.anchorColumn),
    lastColumn: Math.max(selection.column, selection.anchorColumn)
  };
}

export function rectSize(rect: Rect): { rows: number; columns: number } {
  return { rows: rect.lastRow - rect.firstRow + 1, columns: rect.lastColumn - rect.firstColumn + 1 };
}

/**
 * The cells of a rectangle, as they were typed.
 *
 * **Typed, not displayed**, and that is the decision this phase turns
 * on. A copy of displayed values round-trips the numbers and loses
 * every formula, so copying a column of totals and pasting it one
 * column over would paste the totals rather than the sums — which is
 * not what a spreadsheet does and not what anybody means. Pasted into
 * another application the formulas arrive as their text, which is what
 * pasting a formula into a text editor gives you anywhere.
 */
export function readRect(document: SheetDocument, rect: Rect): Block {
  const rows: string[][] = [];
  for (let row = rect.firstRow; row <= rect.lastRow; row++) {
    const line: string[] = [];
    for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
      line.push(document.sheet.input(row, column));
    }
    rows.push(line);
  }
  return rows;
}

export function copyRect(document: SheetDocument, rect: Rect): string {
  return toTsv(readRect(document, rect));
}

export function clearRect(document: SheetDocument, rect: Rect): void {
  document.transact(() => {
    for (let row = rect.firstRow; row <= rect.lastRow; row++) {
      for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
        document.setCell(row, column, '');
      }
    }
  });
}

/** Where a block of text came from, when this sheet is what copied it. */
export interface CopyOrigin {
  readonly text: string;
  readonly row: number;
  readonly column: number;
}

/**
 * Writes a block at a corner, moving its formulas with it.
 *
 * References are rewritten only when the text is what this sheet last
 * copied, and then by the distance it moved: pasting `=B2*C2` one row
 * down gives `=B3*C3`, as dragging it would. Text from somewhere else
 * is written as it arrived — a paste out of Excel is values and
 * formula text that mean what they say, and second-guessing where
 * they were copied from would move references that were never
 * relative to this sheet in the first place.
 */
export function pasteBlock(
  document: SheetDocument,
  text: string,
  at: { row: number; column: number },
  origin: CopyOrigin | null
): Rect {
  const block = fromTsv(text);
  if (block.length === 0) {
    return { firstRow: at.row, lastRow: at.row, firstColumn: at.column, lastColumn: at.column };
  }
  const rowDelta = origin !== null && origin.text === text ? at.row - origin.row : 0;
  const columnDelta = origin !== null && origin.text === text ? at.column - origin.column : 0;

  document.transact(() => {
    block.forEach((line, rowOffset) => {
      line.forEach((cell, columnOffset) => {
        document.setCell(
          at.row + rowOffset,
          at.column + columnOffset,
          rowDelta === 0 && columnDelta === 0 ? cell : rewriteFormula(cell, rowDelta, columnDelta)
        );
      });
    });
  });

  return {
    firstRow: at.row,
    lastRow: at.row + block.length - 1,
    firstColumn: at.column,
    lastColumn: at.column + block[0].length - 1
  };
}

/**
 * Extends a selection over a larger rectangle, repeating it.
 *
 * What the fill handle does. The source tiles the target — dragging
 * two cells down four rows gives the pair twice — and each copy is
 * moved by how far it landed from the original, so a column of
 * `=B2*C2` fills as `=B3*C3`, `=B4*C4`. A source cell filled onto
 * itself is left alone rather than rewritten by zero, which keeps the
 * text somebody typed exactly as they typed it.
 */
export function fillRect(document: SheetDocument, source: Rect, target: Rect): void {
  const { rows, columns } = rectSize(source);
  const block = readRect(document, source);

  document.transact(() => {
    for (let row = target.firstRow; row <= target.lastRow; row++) {
      for (let column = target.firstColumn; column <= target.lastColumn; column++) {
        if (
          row >= source.firstRow &&
          row <= source.lastRow &&
          column >= source.firstColumn &&
          column <= source.lastColumn
        ) {
          continue;
        }
        const rowOffset = mod(row - source.firstRow, rows);
        const columnOffset = mod(column - source.firstColumn, columns);
        const from = { row: source.firstRow + rowOffset, column: source.firstColumn + columnOffset };
        document.setCell(row, column, rewriteFormula(block[rowOffset][columnOffset], row - from.row, column - from.column));
      }
    }
  });
}

/** A remainder that is never negative, for tiling upwards or leftwards. */
function mod(value: number, by: number): number {
  return ((value % by) + by) % by;
}

/** The rectangle a fill covers: the source grown to reach a cell. */
export function fillTarget(source: Rect, toRow: number, toColumn: number): Rect {
  return {
    firstRow: Math.min(source.firstRow, toRow),
    lastRow: Math.max(source.lastRow, toRow),
    firstColumn: Math.min(source.firstColumn, toColumn),
    lastColumn: Math.max(source.lastColumn, toColumn)
  };
}
