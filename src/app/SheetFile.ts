import { DEFAULT_FORMAT, GENERAL, PLAIN, type CellFormat, type CellPaint, type NumberFormat } from '../sheet/Format';
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
 * somebody's disk — and Phase 9 is the phase that needed it. A v1
 * file is a v2 file with no formats, so it loads, and that is the
 * whole job the field was put there to do.
 */
export interface SheetSnapshot {
  readonly version: 2;
  readonly cells: readonly StoredCell[];
  /**
   * The formats, as a palette and a list of cells pointing into it.
   *
   * The same shape the wire carries and for the same reason: a column
   * formatted as currency is one entry and an integer per cell, and a
   * format record per cell would make the file grow with the sheet
   * rather than with the number of distinct formats in it.
   *
   * Written through `Formats.compact`, so nothing no cell points at
   * reaches the disk.
   */
  readonly palette: readonly CellFormat[];
  readonly formats: readonly StoredFormat[];
  /**
   * Whole rows, columns and the sheet, each one number.
   *
   * The reason a sheet made bold is a file of a few hundred bytes
   * rather than thirty megabytes of per-cell entries.
   */
  readonly regions: StoredRegions;
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

export interface StoredRegions {
  readonly sheet: number;
  readonly rows: readonly (readonly [number, number])[];
  readonly columns: readonly (readonly [number, number])[];
}

export interface StoredFormat {
  readonly row: number;
  readonly column: number;
  /** An index into the snapshot's own palette. */
  readonly id: number;
}

/**
 * The sheet as it would be written down.
 *
 * `rowCount` is the sheet's own height, and cells at or below it are
 * left out. The store is addressed by an integer key and will hold a
 * cell anywhere in a million rows, but the *sheet* is `rowCount` tall
 * — nothing outside that can be scrolled to, selected or typed in, so
 * nothing outside it is somebody's data.
 *
 * What lives out there is the proof surface's chain of two hundred
 * thousand cells. Without this it was saved: pressing the button once
 * put a quarter of a million formulas into the file behind the sheet
 * and every load after that parsed and recalculated all of them, for
 * cells the person could never reach and had not asked for. The chain
 * is a measuring instrument, and an instrument is not a document.
 */
export function snapshotOf(
  document: SheetDocument,
  columnWidths: readonly number[],
  rowCount = Number.POSITIVE_INFINITY
): SheetSnapshot {
  const cells: StoredCell[] = [];
  for (const cell of document.sheet.entries()) {
    if (cell.row < rowCount) {
      cells.push(cell);
    }
  }
  const formats = document.formats.compact();
  return {
    version: 2,
    cells,
    palette: formats.palette,
    formats: formats.cells.filter(cell => cell.row < rowCount),
    regions: {
      sheet: formats.regions.sheet,
      rows: formats.regions.rows.filter(([row]) => row < rowCount),
      columns: formats.regions.columns
    },
    columnWidths: [...columnWidths]
  };
}

/**
 * Puts a snapshot into a document, without it counting as an edit.
 *
 * Written through the model rather than through `SheetDocument.setCell`
 * so that loading a file does not land on the undo stack: a person's
 * first ctrl-Z after opening a sheet should do nothing, not empty it.
 */
export function applySnapshot(document: SheetDocument, snapshot: SheetSnapshot): void {
  // Formats first, so that a Text-formatted cell is Text *before* its
  // input is read: written the other way round, `007` would be parsed
  // as the number seven and then formatted as text, and the leading
  // zeros somebody saved would be gone by the time the format said
  // to keep them.
  document.formats.restore(snapshot.palette, snapshot.formats, {
    sheet: snapshot.regions.sheet,
    rows: snapshot.regions.rows,
    columns: snapshot.regions.columns
  });
  for (const cell of snapshot.cells) {
    document.sheet.setCell(
      cell.row,
      cell.column,
      cell.input,
      document.formats.formatAt(cell.row, cell.column).number.kind === 'text'
    );
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
  const source = raw as Omit<Partial<SheetSnapshot>, 'version'> & { version?: number };
  // A v1 file is a v2 file with no formats in it, so it is read
  // rather than refused. The version exists to let an old file keep
  // working, and refusing one would be the version field costing
  // exactly what it was meant to save.
  if ((source.version !== 1 && source.version !== 2) || !Array.isArray(source.cells)) {
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
  const palette = paletteFrom(source.palette);
  return {
    version: 2,
    cells,
    palette,
    formats: formatsFrom(source.formats, palette.length),
    regions: regionsFrom(source.regions, palette.length),
    columnWidths: widthsFrom(source.columnWidths, columnCount)
  };
}

/**
 * The palette, with every entry rebuilt field by field.
 *
 * Not trusted and not spread: a file is outside the program, and an
 * entry with a missing `align` would otherwise reach the render
 * worker as `undefined` and be drawn as nothing at all. Each field
 * falls back to the default, so a format that is half-understood
 * loses the half nobody can read rather than the cell.
 */
function paletteFrom(stored: unknown): CellFormat[] {
  const palette: CellFormat[] = [DEFAULT_FORMAT];
  if (!Array.isArray(stored)) {
    return palette;
  }
  for (let id = 1; id < stored.length; id++) {
    const entry = stored[id] as Partial<CellFormat> | null;
    if (typeof entry !== 'object' || entry === null) {
      palette.push(DEFAULT_FORMAT);
      continue;
    }
    palette.push({ number: numberFrom(entry.number), paint: paintFrom(entry.paint) });
  }
  return palette;
}

function numberFrom(stored: unknown): NumberFormat {
  if (typeof stored !== 'object' || stored === null) {
    return GENERAL;
  }
  const format = stored as NumberFormat;
  const places = (value: unknown): number =>
    typeof value === 'number' && Number.isFinite(value) ? Math.min(Math.max(Math.trunc(value), 0), 15) : 2;
  switch (format.kind) {
    case 'number':
      return { kind: 'number', places: places(format.places), thousands: format.thousands === true };
    case 'currency':
      return {
        kind: 'currency',
        places: places(format.places),
        symbol: typeof format.symbol === 'string' ? format.symbol : '$'
      };
    case 'percent':
      return { kind: 'percent', places: places(format.places) };
    case 'scientific':
      return { kind: 'scientific', places: places(format.places) };
    case 'date':
      return { kind: 'date', pattern: format.pattern === 'dmy' || format.pattern === 'mdy' ? format.pattern : 'ymd' };
    case 'time':
      return { kind: 'time', pattern: format.pattern === 'hms' ? 'hms' : 'hm' };
    case 'text':
      return { kind: 'text' };
    default:
      return GENERAL;
  }
}

function paintFrom(stored: unknown): CellPaint {
  if (typeof stored !== 'object' || stored === null) {
    return PLAIN;
  }
  const paint = stored as Partial<CellPaint>;
  const align = paint.align;
  return {
    bold: paint.bold === true,
    italic: paint.italic === true,
    underline: paint.underline === true,
    fontSize: typeof paint.fontSize === 'number' && paint.fontSize >= 0 && paint.fontSize <= 96 ? paint.fontSize : 0,
    color: typeof paint.color === 'string' ? paint.color : '',
    fill: typeof paint.fill === 'string' ? paint.fill : '',
    align: align === 'start' || align === 'center' || align === 'end' ? align : 'auto',
    wrap: paint.wrap === true
  };
}

/** Rows, columns and the sheet, with anything pointing off the palette dropped. */
function regionsFrom(stored: unknown, paletteLength: number): StoredRegions {
  const empty: StoredRegions = { sheet: 0, rows: [], columns: [] };
  if (typeof stored !== 'object' || stored === null) {
    return empty;
  }
  const source = stored as Partial<StoredRegions>;
  const pairs = (value: unknown): (readonly [number, number])[] => {
    if (!Array.isArray(value)) {
      return [];
    }
    return value.filter(
      (pair): pair is [number, number] =>
        Array.isArray(pair) &&
        pair.length === 2 &&
        Number.isInteger(pair[0]) &&
        pair[0] >= 0 &&
        Number.isInteger(pair[1]) &&
        pair[1] > 0 &&
        pair[1] < paletteLength
    );
  };
  const sheet = source.sheet;
  return {
    sheet: Number.isInteger(sheet) && (sheet as number) > 0 && (sheet as number) < paletteLength ? (sheet as number) : 0,
    rows: pairs(source.rows),
    columns: pairs(source.columns)
  };
}

/** Cells pointing into the palette, with anything pointing off it dropped. */
function formatsFrom(stored: unknown, paletteLength: number): StoredFormat[] {
  const formats: StoredFormat[] = [];
  if (!Array.isArray(stored)) {
    return formats;
  }
  for (const entry of stored) {
    const cell = entry as Partial<StoredFormat>;
    if (
      typeof cell === 'object' &&
      cell !== null &&
      Number.isInteger(cell.row) &&
      Number.isInteger(cell.column) &&
      Number.isInteger(cell.id) &&
      (cell.id as number) > 0 &&
      (cell.id as number) < paletteLength
    ) {
      formats.push({ row: cell.row as number, column: cell.column as number, id: cell.id as number });
    }
  }
  return formats;
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
