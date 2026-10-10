import { DEFAULT_FORMAT, type CellFormat } from '../sheet/Format';
import { moveFormula } from '../sheet/Move';
import { rewriteFormula } from '../sheet/Rewrite';
import { fromTsv, toTsv, type Block } from '../sheet/Tsv';
import { isError, type CellValue } from '../sheet/Values';
import { literalOf } from '../sheet/Workbook';
import type { SheetDocument } from './SheetDocument';
import { cornerOf, type SheetSelection } from './SheetContract';

/** A selection as the rectangle it covers, corners normalised. */
export interface Rect {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

export function rectOf(selection: SheetSelection): Rect {
  return {
    firstRow: Math.min(cornerOf(selection).row, selection.anchorRow),
    lastRow: Math.max(cornerOf(selection).row, selection.anchorRow),
    firstColumn: Math.min(cornerOf(selection).column, selection.anchorColumn),
    lastColumn: Math.max(cornerOf(selection).column, selection.anchorColumn)
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
      line.push(document.inputAt(row, column));
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

/**
 * One input written across a rectangle, each copy's references moved
 * by how far it is from `from` — what Ctrl+Enter does with what was
 * typed. One transaction, so one step of undo.
 */
export function writeRect(document: SheetDocument, rect: Rect, input: string, from: { row: number; column: number }): void {
  document.transact(() => {
    for (let row = rect.firstRow; row <= rect.lastRow; row++) {
      for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
        document.setCell(row, column, rewriteFormula(input, row - from.row, column - from.column));
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

/**
 * The block of data a cell is standing in.
 *
 * Excel calls this the current region, and it is what "sort my table"
 * has to mean when somebody has selected one cell — which is how
 * everybody sorts. It grows outwards from the cell while the next row
 * or column along has anything in it, and stops at the blank line
 * that separates one table from the next.
 *
 * Getting this wrong is not a small thing. Widening a single cell to
 * the *whole sheet* — which this did, once — sorts three unrelated
 * tables into one, and drags formulas across each other until some of
 * them point off the sheet and say `#REF!`. It was undone by one
 * press of ctrl-Z, and it should never have been offered.
 *
 * It is here and not on the screen because the screen cannot see
 * where the data stops: the render worker holds the rows it has
 * mounted, and the block may be larger or smaller than that.
 */
export function currentRegion(
  document: SheetDocument,
  row: number,
  column: number,
  rowCount: number,
  columnCount: number
): Rect {
  let rect: Rect = { firstRow: row, lastRow: row, firstColumn: column, lastColumn: column };
  for (;;) {
    const grown = grow(document, rect, rowCount, columnCount);
    if (
      grown.firstRow === rect.firstRow &&
      grown.lastRow === rect.lastRow &&
      grown.firstColumn === rect.firstColumn &&
      grown.lastColumn === rect.lastColumn
    ) {
      return rect;
    }
    rect = grown;
  }
}

function grow(document: SheetDocument, rect: Rect, rowCount: number, columnCount: number): Rect {
  const filledRow = (at: number): boolean => {
    if (at < 0 || at >= rowCount) {
      return false;
    }
    for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
      if (document.inputAt(at, column) !== '') {
        return true;
      }
    }
    return false;
  };
  const filledColumn = (at: number): boolean => {
    if (at < 0 || at >= columnCount) {
      return false;
    }
    for (let row = rect.firstRow; row <= rect.lastRow; row++) {
      if (document.inputAt(row, at) !== '') {
        return true;
      }
    }
    return false;
  };
  return {
    firstRow: filledRow(rect.firstRow - 1) ? rect.firstRow - 1 : rect.firstRow,
    lastRow: filledRow(rect.lastRow + 1) ? rect.lastRow + 1 : rect.lastRow,
    firstColumn: filledColumn(rect.firstColumn - 1) ? rect.firstColumn - 1 : rect.firstColumn,
    lastColumn: filledColumn(rect.lastColumn + 1) ? rect.lastColumn + 1 : rect.lastColumn
  };
}

/**
 * Whether a block's first row reads as a heading.
 *
 * The guess every spreadsheet makes, and it makes it because asking
 * is worse: a dialog in front of a sort is a dialog people dismiss
 * without reading. Text over anything that is not all text is a
 * heading; a column of names under the word `Name` is not, which is
 * why the row below has to disagree with it for this to say yes.
 */
export function looksLikeHeader(document: SheetDocument, rect: Rect): boolean {
  if (rect.lastRow <= rect.firstRow) {
    return false;
  }
  let sawText = false;
  for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
    const head = document.shown.value(rect.firstRow, column);
    const below = document.shown.value(rect.firstRow + 1, column);
    if (head === null) {
      continue;
    }
    if (typeof head !== 'string') {
      return false;
    }
    sawText = true;
    if (typeof below !== 'string' && below !== null) {
      return true;
    }
  }
  // Every column is text all the way down, or the block is one row of
  // text: no evidence either way, and leaving the first row in place
  // is the answer that cannot scramble a heading into the data.
  return sawText;
}

/**
 * A block as it was when it was copied: what was typed, what it came
 * to, and how it was formatted, cell by cell.
 *
 * Kept whole rather than re-read at paste time because the paste may
 * be onto the block itself, or after the source has changed — and
 * because a cut is pasted from here after the source is gone.
 */
export interface Copied {
  /** The text put on the clipboard, which is how a paste knows it is this copy. */
  readonly text: string;
  readonly sheet: number;
  readonly rect: Rect;
  readonly cut: boolean;
  readonly inputs: readonly (readonly string[])[];
  readonly values: readonly (readonly CellValue[])[];
  readonly formats: readonly (readonly CellFormat[])[];
}

export function copiedOf(document: SheetDocument, rect: Rect, cut: boolean): Copied {
  const inputs: string[][] = [];
  const values: CellValue[][] = [];
  const formats: CellFormat[][] = [];
  for (let row = rect.firstRow; row <= rect.lastRow; row++) {
    const typed: string[] = [];
    const came: CellValue[] = [];
    const looks: CellFormat[] = [];
    for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
      typed.push(document.inputAt(row, column));
      came.push(document.shown.value(row, column));
      looks.push(document.formatAt(row, column));
    }
    inputs.push(typed);
    values.push(came);
    formats.push(looks);
  }
  return { text: toTsv(inputs), sheet: document.active, rect, cut, inputs, values, formats };
}

/** What Paste special can paste. */
export type PasteMode = 'all' | 'values' | 'formats' | 'transposed';

const TEXT_FORMAT: CellFormat = { ...DEFAULT_FORMAT, number: { kind: 'text' } };

/**
 * A copied block written at a corner, in one of the four ways.
 *
 * - `all` is what was typed, with formulas moved by how far they went,
 *   and the formats with them;
 * - `values` is what the cells came to, as literals — a formula lands
 *   as its answer — into the formats already there, as in Excel. Text
 *   that would not read back as itself is kept as text;
 * - `formats` is the formats and nothing else;
 * - `transposed` is `all` with rows as columns, each formula moved by
 *   how far its own cell went.
 */
export function pasteCopied(document: SheetDocument, copied: Copied, at: { row: number; column: number }, mode: PasteMode): Rect {
  const rows = copied.inputs.length;
  const columns = copied.inputs[0]?.length ?? 0;
  const flipped = mode === 'transposed';
  const place = (row: number, column: number) =>
    flipped ? { row: at.row + column, column: at.column + row } : { row: at.row + row, column: at.column + column };
  document.transact(() => {
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < columns; column++) {
        const to = place(row, column);
        if (mode !== 'values') {
          document.setFormat(to.row, to.column, copied.formats[row][column]);
        }
        if (mode === 'formats') {
          continue;
        }
        if (mode === 'values') {
          const value = copied.values[row][column];
          const text = literalText(value);
          if (typeof value === 'string' && literalOf(value) !== value) {
            document.setFormat(to.row, to.column, TEXT_FORMAT);
          }
          document.setCell(to.row, to.column, text);
          continue;
        }
        const from = { row: copied.rect.firstRow + row, column: copied.rect.firstColumn + column };
        const input = copied.inputs[row][column];
        const rowDelta = to.row - from.row;
        const columnDelta = to.column - from.column;
        document.setCell(to.row, to.column, rowDelta === 0 && columnDelta === 0 ? input : rewriteFormula(input, rowDelta, columnDelta));
      }
    }
  });
  const height = flipped ? columns : rows;
  const width = flipped ? rows : columns;
  return { firstRow: at.row, lastRow: at.row + height - 1, firstColumn: at.column, lastColumn: at.column + width - 1 };
}

/** A value as the literal that types it back. */
function literalText(value: CellValue): string {
  if (value === null) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (isError(value)) {
    return `=${value.code}`;
  }
  return String(value);
}

/** Text from somewhere else, turned on its side. */
export function transposedText(text: string): string {
  const block = fromTsv(text);
  const width = Math.max(0, ...block.map(line => line.length));
  const out: string[][] = [];
  for (let column = 0; column < width; column++) {
    out.push(block.map(line => line[column] ?? ''));
  }
  return toTsv(out);
}

/**
 * A cut, pasted: the block moves, and every formula in the workbook
 * that pointed into it points at the same cells in their new place —
 * see `Move.ts`. One step of undo, whatever it touched.
 */
export function moveCopied(document: SheetDocument, copied: Copied, at: { row: number; column: number }, onSheet: number): Rect {
  const names = document.book.sheetNames();
  const from = names[copied.sheet];
  const to = names[onSheet];
  const rows = copied.inputs.length;
  const columns = copied.inputs[0]?.length ?? 0;
  const target: Rect = { firstRow: at.row, lastRow: at.row + rows - 1, firstColumn: at.column, lastColumn: at.column + columns - 1 };
  if (from === undefined || to === undefined) {
    return target;
  }
  const move = {
    sheet: from,
    ...copied.rect,
    rowDelta: at.row - copied.rect.firstRow,
    columnDelta: at.column - copied.rect.firstColumn,
    ...(onSheet === copied.sheet ? {} : { toSheet: to })
  };
  const inBlock = (sheet: number, row: number, column: number, rect: Rect): boolean =>
    row >= rect.firstRow && row <= rect.lastRow && column >= rect.firstColumn && column <= rect.lastColumn && sheet >= 0;

  // Everything that reads the block, found before anything moves.
  const rewrites: { sheet: number; row: number; column: number; input: string }[] = [];
  for (let sheet = 0; sheet < document.sheetCount; sheet++) {
    const page = document.pageAt(sheet);
    if (page === undefined) {
      continue;
    }
    for (const cell of page.sheet.entries()) {
      if (!cell.input.startsWith('=')) {
        continue;
      }
      if (sheet === copied.sheet && inBlock(sheet, cell.row, cell.column, copied.rect)) {
        continue;
      }
      if (sheet === onSheet && inBlock(sheet, cell.row, cell.column, target)) {
        continue;
      }
      const moved = moveFormula(cell.input, move, names[sheet] ?? '');
      if (moved !== cell.input) {
        rewrites.push({ sheet, row: cell.row, column: cell.column, input: moved });
      }
    }
  }

  const active = document.active;
  document.transact(() => {
    document.activate(copied.sheet);
    for (let row = copied.rect.firstRow; row <= copied.rect.lastRow; row++) {
      for (let column = copied.rect.firstColumn; column <= copied.rect.lastColumn; column++) {
        document.setCell(row, column, '');
        document.setFormat(row, column, DEFAULT_FORMAT);
      }
    }
    document.activate(onSheet);
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < columns; column++) {
        document.setFormat(at.row + row, at.column + column, copied.formats[row][column]);
        document.setCell(at.row + row, at.column + column, moveFormula(copied.inputs[row][column], move, from, true));
      }
    }
    for (const rewrite of rewrites) {
      document.activate(rewrite.sheet);
      document.setCell(rewrite.row, rewrite.column, rewrite.input);
    }
    document.activate(active);
  });
  return target;
}
