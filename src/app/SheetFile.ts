import { relativeRef } from '../sheet/A1';
import { nameProblem } from '../sheet/Names';
import type { MergeRect } from '../sheet/Merges';
import {
  DEFAULT_FORMAT,
  GENERAL,
  NO_BORDERS,
  NO_EDGE,
  PLAIN,
  type CellBorders,
  type CellEdge,
  type CellFormat,
  type CellPaint,
  type NumberFormat
} from '../sheet/Format';
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
 * file is a v2 file with no formats, and a v2 file is a v3 file of
 * one sheet, so both load and that is the whole job the field was put
 * there to do.
 *
 * v3 is the one version bump that could not be an absent field. A v2
 * file has its cells at the top level and a v3 file has them inside a
 * sheet, so a reader that guessed would read a whole workbook as
 * nothing at all.
 */
export interface SheetSnapshot {
  readonly version: 3;
  /** Every sheet, in tab order. */
  readonly sheets: readonly StoredSheet[];
  /** Which one was showing when it was written. */
  readonly active: number;
  /**
   * The names, which belong to the workbook rather than to a sheet.
   *
   * Each says which sheet its range is on, because `Sales` on Sheet 2
   * and `Sales` on Sheet 1 are different cells and the name alone
   * cannot tell them apart. Absent in a v2 file, which had one sheet
   * — so it reads as the first one, which is where those cells were.
   */
  readonly names: readonly StoredName[];
}

/**
 * One sheet of the workbook, and everything drawn over its cells.
 *
 * All of it per sheet and none of it shared, which is the shape the
 * document already has: widths, hidden rows, freezes, merges and
 * formats are facts about *a* sheet. A v3 file is therefore a v2 file
 * with a name and a colour on it, repeated.
 */
export interface StoredSheet {
  readonly name: string;
  /** A tab colour somebody chose, or null for the plain one. */
  readonly colour: string | null;
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
   * reaches the disk. A palette per sheet rather than one for the
   * workbook, because sheets are added and removed and a shared
   * palette would need a reference count to know when an entry could
   * go — for a saving of a few hundred bytes on a file that already
   * holds every formula.
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
  /** The merged rectangles. */
  readonly merges: readonly MergeRect[];
  readonly frozenRows: number;
  readonly frozenColumns: number;
  readonly hiddenRows: readonly number[];
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
export function snapshotOf(document: SheetDocument, rowCount = Number.POSITIVE_INFINITY): SheetSnapshot {
  const sheets: StoredSheet[] = [];
  for (let index = 0; index < document.sheetCount; index++) {
    const page = document.pageAt(index);
    if (page === undefined) {
      continue;
    }
    const cells: StoredCell[] = [];
    for (const cell of page.sheet.entries()) {
      if (cell.row < rowCount) {
        cells.push(cell);
      }
    }
    const formats = page.formats.compact();
    sheets.push({
      name: page.sheet.name,
      colour: page.sheet.colour,
      cells,
      palette: formats.palette,
      formats: formats.cells.filter(cell => cell.row < rowCount),
      regions: {
        sheet: formats.regions.sheet,
        rows: formats.regions.rows.filter(([row]) => row < rowCount),
        columns: formats.regions.columns
      },
      merges: page.merges.all.filter(rect => rect.lastRow < rowCount),
      frozenRows: page.frozenRows,
      frozenColumns: page.frozenColumns,
      hiddenRows: [...page.hiddenRows].filter(row => row < rowCount).sort((a, b) => a - b),
      columnWidths: [...page.columnWidths]
    });
  }
  return {
    version: 3,
    sheets,
    active: document.active,
    names: document.book.names.all().map(entry => ({
      name: entry.name,
      // The name of the sheet rather than its index, so a name
      // survives the sheets being reordered between two saves — and
      // so a file somebody read can be understood without counting
      // tabs. Unknown on load means the first sheet, which is where a
      // v2 file's names were.
      sheet: entry.range.start.sheet ?? null,
      firstRow: Math.min(entry.range.start.row, entry.range.end.row),
      firstColumn: Math.min(entry.range.start.column, entry.range.end.column),
      lastRow: Math.max(entry.range.start.row, entry.range.end.row),
      lastColumn: Math.max(entry.range.start.column, entry.range.end.column)
    }))
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
  // The sheets before anything on them, because every write below
  // goes through the active page and there has to be one to go
  // through. Named first too: a formula reading `Data!A1` can only
  // find `Data` if `Data` exists by the time it is parsed.
  document.restoreSheets(snapshot.sheets.map(stored => stored.name));
  for (const [index, stored] of snapshot.sheets.entries()) {
    document.activate(index);
    document.setSheetColour(index, stored.colour);
    const page = document.pageAt(index);
    if (page === undefined) {
      continue;
    }
    // Formats first, so that a Text-formatted cell is Text *before*
    // its input is read: written the other way round, `007` would be
    // parsed as the number seven and then formatted as text, and the
    // leading zeros somebody saved would be gone by the time the
    // format said to keep them.
    page.merges.restore(stored.merges);
    page.frozenRows = stored.frozenRows;
    page.frozenColumns = stored.frozenColumns;
    page.hiddenRows.clear();
    for (const row of stored.hiddenRows) {
      page.hiddenRows.add(row);
    }
    page.columnWidths = [...stored.columnWidths];
    page.formats.restore(stored.palette, stored.formats, {
      sheet: stored.regions.sheet,
      rows: stored.regions.rows,
      columns: stored.regions.columns
    });
    for (const cell of stored.cells) {
      page.sheet.setCell(
        cell.row,
        cell.column,
        cell.input,
        page.formats.formatAt(cell.row, cell.column).number.kind === 'text'
      );
    }
  }
  document.book.names.restore(
    snapshot.names.map(stored => {
      const on = stored.sheet === null ? undefined : stored.sheet;
      return {
        name: stored.name,
        range: {
          start: { ...relativeRef(stored.firstRow, stored.firstColumn), sheet: on },
          end: { ...relativeRef(stored.lastRow, stored.lastColumn), sheet: on }
        }
      };
    })
  );
  // The names arrive after the cells, so the formulas that read them
  // were wired when the name meant nothing. This is the same wake-up
  // defining one by hand gives.
  document.book.namesChanged();
  document.activate(Math.min(Math.max(snapshot.active, 0), Math.max(snapshot.sheets.length - 1, 0)));
  document.book.recalculate();
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
  const source = raw as Record<string, unknown>;
  const version = source.version;

  /**
   * A v1 or v2 file is a workbook of one sheet, and is read as one.
   *
   * The version exists to let an old file keep working, and refusing
   * one would be the field costing exactly what it was meant to save.
   * A v1 file is a v2 file with no formats; a v2 file is this, with
   * its one sheet's fields at the top level where they used to be.
   */
  if (version === 1 || version === 2) {
    if (!Array.isArray(source.cells)) {
      return null;
    }
    return {
      version: 3,
      sheets: [sheetFrom(source, 'Sheet1', columnCount)],
      active: 0,
      names: namesFrom(source.names)
    };
  }

  if (version !== 3 || !Array.isArray(source.sheets)) {
    return null;
  }
  const sheets: StoredSheet[] = [];
  for (const [index, stored] of (source.sheets as unknown[]).entries()) {
    if (typeof stored !== 'object' || stored === null) {
      continue;
    }
    const held = stored as Record<string, unknown>;
    sheets.push(
      sheetFrom(held, typeof held.name === 'string' && held.name.trim() !== '' ? held.name : `Sheet${index + 1}`, columnCount)
    );
  }
  // A workbook of none is a state nothing else is written to survive,
  // and a file that says so is a file that has been truncated.
  if (sheets.length === 0) {
    return null;
  }
  const active = source.active;
  return {
    version: 3,
    sheets,
    active: Number.isInteger(active) && (active as number) >= 0 && (active as number) < sheets.length ? (active as number) : 0,
    names: namesFrom(source.names)
  };
}

/**
 * One sheet's worth of fields, wherever they were found.
 *
 * Shared by the v3 path and the v2 one, which is what makes "a v2
 * file is a workbook of one sheet" a fact about the reader rather
 * than a sentence in a comment: both go through the same checks and
 * neither can drift.
 */
function sheetFrom(source: Record<string, unknown>, name: string, columnCount: number): StoredSheet {
  const cells: StoredCell[] = [];
  if (Array.isArray(source.cells)) {
    for (const cell of source.cells) {
      if (
        typeof cell === 'object' &&
        cell !== null &&
        Number.isInteger((cell as StoredCell).row) &&
        Number.isInteger((cell as StoredCell).column) &&
        typeof (cell as StoredCell).input === 'string'
      ) {
        cells.push({
          row: (cell as StoredCell).row,
          column: (cell as StoredCell).column,
          input: (cell as StoredCell).input
        });
      }
    }
  }
  const palette = paletteFrom(source.palette);
  return {
    name,
    colour: typeof source.colour === 'string' ? source.colour : null,
    cells,
    palette,
    formats: formatsFrom(source.formats, palette.length),
    regions: regionsFrom(source.regions, palette.length),
    merges: mergesFrom(source.merges),
    frozenRows: countFrom(source.frozenRows),
    frozenColumns: countFrom(source.frozenColumns),
    hiddenRows: Array.isArray(source.hiddenRows)
      ? source.hiddenRows.filter((row): row is number => Number.isInteger(row) && row >= 0)
      : [],
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
    case 'datetime':
      return {
        kind: 'datetime',
        date: format.date === 'dmy' || format.date === 'mdy' ? format.date : 'ymd',
        time: format.time === 'hms' ? 'hms' : 'hm'
      };
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
    wrap: paint.wrap === true,
    borders: bordersFrom(paint.borders)
  };
}

function bordersFrom(stored: unknown): CellBorders {
  if (typeof stored !== 'object' || stored === null) {
    return NO_BORDERS;
  }
  const borders = stored as Partial<CellBorders>;
  return {
    top: edgeFrom(borders.top),
    right: edgeFrom(borders.right),
    bottom: edgeFrom(borders.bottom),
    left: edgeFrom(borders.left)
  };
}

function edgeFrom(stored: unknown): CellEdge {
  if (typeof stored !== 'object' || stored === null) {
    return NO_EDGE;
  }
  const edge = stored as Partial<CellEdge>;
  const width = edge.width;
  return {
    width: typeof width === 'number' && Number.isFinite(width) && width > 0 ? Math.min(width, 8) : 0,
    color: typeof edge.color === 'string' ? edge.color : ''
  };
}

/** Merged rectangles, with anything malformed dropped rather than trusted. */
function mergesFrom(stored: unknown): MergeRect[] {
  if (!Array.isArray(stored)) {
    return [];
  }
  return stored.filter((rect): rect is MergeRect => {
    const candidate = rect as Partial<MergeRect>;
    return (
      typeof candidate === 'object' &&
      candidate !== null &&
      Number.isInteger(candidate.firstRow) &&
      Number.isInteger(candidate.lastRow) &&
      Number.isInteger(candidate.firstColumn) &&
      Number.isInteger(candidate.lastColumn) &&
      (candidate.lastRow as number) >= (candidate.firstRow as number) &&
      (candidate.lastColumn as number) >= (candidate.firstColumn as number)
    );
  });
}

function countFrom(stored: unknown): number {
  return typeof stored === 'number' && Number.isInteger(stored) && stored >= 0 ? stored : 0;
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

/** A name as a file holds it: its text, its sheet, and the corners it names. */
export interface StoredName {
  readonly name: string;
  /**
   * The sheet its range is on, by name, or null for the first one.
   *
   * By name rather than by index so it survives the tabs being
   * reordered between two saves, and so a file can be read without
   * counting them. Null is what a v2 file means: it had one sheet.
   */
  readonly sheet: string | null;
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

/**
 * Names read back out of a file, with the bad ones dropped.
 *
 * Checked against the same rules a person's typing is checked
 * against, because a file is untrusted input in exactly the way a
 * keystroke is: a name saved by a future version, or edited by hand,
 * must not become a name this sheet cannot express.
 */
function namesFrom(stored: unknown): StoredName[] {
  if (!Array.isArray(stored)) {
    return [];
  }
  const found: StoredName[] = [];
  for (const entry of stored) {
    if (typeof entry !== 'object' || entry === null) {
      continue;
    }
    const held = entry as Partial<StoredName>;
    if (typeof held.name !== 'string' || nameProblem(held.name) !== null) {
      continue;
    }
    const corners = [held.firstRow, held.firstColumn, held.lastRow, held.lastColumn];
    if (!corners.every(value => Number.isInteger(value) && (value as number) >= 0)) {
      continue;
    }
    found.push({
      name: held.name,
      sheet: typeof held.sheet === 'string' ? held.sheet : null,
      firstRow: held.firstRow as number,
      firstColumn: held.firstColumn as number,
      lastRow: held.lastRow as number,
      lastColumn: held.lastColumn as number
    });
  }
  return found;
}
