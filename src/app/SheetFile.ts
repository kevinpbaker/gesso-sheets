import { COLUMN_WIDTH } from './dimensions';
import type { SheetDocument } from './SheetDocument';

/**
 * A sheet as it is written down.
 *
 * **Inputs, not values.** `=SUM(A1:A9)` is stored, not the number it
 * showed: the value is derivable and the formula is not, so writing
 * values would be writing the one half that can always be rebuilt and
 * losing the half that cannot. Reloading recalculates, which is also
 * the cheapest possible check that the engine still agrees with
 * itself.
 *
 * `version` is here from the first write rather than added when it is
 * first needed, because the file that needs it is the one already on
 * somebody's disk.
 */
export interface SheetSnapshot {
  readonly version: 1;
  readonly cells: readonly StoredCell[];
  /**
   * Column widths, in order from A.
   *
   * On the document rather than on the screen, which is a change from
   * Phase 3: how wide a column is drawn is not the application's
   * business until it has to survive a reload, and then it is.
   */
  readonly columnWidths: readonly number[];
}

export interface StoredCell {
  readonly row: number;
  readonly column: number;
  readonly input: string;
}

export function snapshotOf(document: SheetDocument, columnWidths: readonly number[]): SheetSnapshot {
  return { version: 1, cells: [...document.sheet.entries()], columnWidths: [...columnWidths] };
}

/**
 * Puts a snapshot into a document, without it counting as an edit.
 *
 * Written through the model rather than through `SheetDocument.setCell`
 * so that loading a file does not land on the undo stack: a person's
 * first ctrl-Z after opening a sheet should do nothing, not empty it.
 */
export function applySnapshot(document: SheetDocument, snapshot: SheetSnapshot): void {
  for (const cell of snapshot.cells) {
    document.sheet.setCell(cell.row, cell.column, cell.input);
  }
  document.sheet.recalculate();
}

/**
 * Reads a snapshot back out of whatever was on disk.
 *
 * Every field is checked rather than trusted. The file is outside the
 * program — an older build wrote it, or a newer one, or something
 * truncated it — and a store that threw on a bad file would lose a
 * sheet that is mostly fine over one cell that is not.
 */
export function parseSnapshot(text: string, columnCount: number): SheetSnapshot | null {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return null;
  }
  if (typeof raw !== 'object' || raw === null) {
    return null;
  }
  const source = raw as Partial<SheetSnapshot>;
  if (source.version !== 1 || !Array.isArray(source.cells)) {
    return null;
  }
  const cells: StoredCell[] = [];
  for (const cell of source.cells) {
    if (
      typeof cell === 'object' &&
      cell !== null &&
      Number.isInteger((cell as StoredCell).row) &&
      Number.isInteger((cell as StoredCell).column) &&
      typeof (cell as StoredCell).input === 'string'
    ) {
      cells.push({ row: (cell as StoredCell).row, column: (cell as StoredCell).column, input: (cell as StoredCell).input });
    }
  }
  return { version: 1, cells, columnWidths: widthsFrom(source.columnWidths, columnCount) };
}

function widthsFrom(stored: unknown, columnCount: number): number[] {
  const widths = Array.from({ length: columnCount }, () => COLUMN_WIDTH);
  if (!Array.isArray(stored)) {
    return widths;
  }
  for (let column = 0; column < Math.min(stored.length, columnCount); column++) {
    const width = stored[column];
    if (typeof width === 'number' && Number.isFinite(width) && width > 0) {
      widths[column] = width;
    }
  }
  return widths;
}
