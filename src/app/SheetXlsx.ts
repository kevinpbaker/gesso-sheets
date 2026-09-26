import { DEFAULT_FORMAT, keyOf, type CellFormat } from '../sheet/Format';
import { MIN_CHART_HEIGHT, MIN_CHART_WIDTH, type Chart } from '../sheet/Chart';
import type { XlsxBook, XlsxSheet } from '../sheet/Xlsx';
import type { XlsxChart } from '../sheet/XlsxCharts';
import { rangeOfRect } from '../sheet/XlsxRules';
import type { Inflate } from '../sheet/Zip';
import { COLUMN_WIDTH, MAX_ROW_HEIGHT, MIN_ROW_HEIGHT, ROW_HEIGHT } from './dimensions';
import type { SheetSnapshot, StoredFormat, StoredSheet } from './SheetFile';

/**
 * An Excel workbook as a snapshot: the shape a `.gsheet` has.
 *
 * Which is the whole design. A snapshot is what a saved workbook is,
 * and Phase 16 already knows how to open one as a document of its own
 * — a fresh workbook swapped in, a library entry, a route. An `.xlsx`
 * turned into one arrives by that path and nothing downstream can tell
 * it was ever anything else, which is also what makes Save afterwards
 * a `.gsheet`: a Save that silently wrote a different format from the
 * one opened would be worse than one that asks where to put a new
 * file. Writing an `.xlsx` is an export, File ▸ Download as Excel
 * workbook — see `SheetXlsxOut`.
 */

export function snapshotOfXlsx(book: XlsxBook, columnCount: number): SheetSnapshot {
  // A chart's id is the workbook's, not the sheet's.
  let chartId = 0;
  return {
    version: 3,
    sheets: book.sheets.map(sheet => storedSheet(sheet, columnCount, () => ++chartId)),
    active: 0,
    names: book.names.map(name => ({
      name: name.name,
      sheet: name.sheet,
      firstRow: name.firstRow,
      firstColumn: name.firstColumn,
      lastRow: name.lastRow,
      lastColumn: name.lastColumn
    })),
    ...(book.iteration === null ? {} : { iteration: book.iteration })
  };
}

function storedSheet(sheet: XlsxSheet, columnCount: number, nextChartId: () => number = () => 1): StoredSheet {
  // The palette the snapshot wants: the default at 0, then each format
  // a cell here actually uses, once.
  const palette: CellFormat[] = [DEFAULT_FORMAT];
  const ids = new Map<string, number>([[keyOf(DEFAULT_FORMAT), 0]]);
  const idOf = (format: CellFormat): number => {
    const key = keyOf(format);
    let id = ids.get(key);
    if (id === undefined) {
      id = palette.length;
      palette.push(format);
      ids.set(key, id);
    }
    return id;
  };

  const formats: StoredFormat[] = [];
  const place = (row: number, column: number, format: CellFormat): void => {
    const id = idOf(format);
    if (id !== 0) {
      formats.push({ row, column, id });
    }
  };
  for (const cell of sheet.cells) {
    const format = sheet.formats[cell.style] ?? DEFAULT_FORMAT;
    if (cell.asText) {
      // Text that would not read back as text is formatted Text, which
      // is the one format that changes what an input *means* — the
      // same path a person takes to keep `007`.
      place(cell.row, cell.column, { ...format, number: { kind: 'text' } });
    } else if (cell.input.startsWith('=') && format.number.kind === 'text') {
      // A formula under a Text format. In Excel it is still a formula —
      // the format only changes how the *next* thing typed is read — and
      // here a Text format turns an input into text, so the formula
      // would come across as its own source. It keeps its paint and
      // loses the one part of the format that would break it.
      place(cell.row, cell.column, { ...format, number: { kind: 'general' } });
    } else {
      place(cell.row, cell.column, format);
    }
  }
  for (const cell of sheet.styled) {
    place(cell.row, cell.column, sheet.formats[cell.style] ?? DEFAULT_FORMAT);
  }

  const columnWidths = Array.from({ length: columnCount }, (_, at) => sheet.columnWidths.get(at) ?? COLUMN_WIDTH);
  return {
    name: sheet.name,
    colour: null,
    cells: sheet.cells.map(cell => ({ row: cell.row, column: cell.column, input: cell.input })),
    palette,
    formats,
    regions: { sheet: 0, rows: [], columns: [] },
    merges: sheet.merges,
    conditional: [...sheet.conditional],
    validations: [...sheet.validations],
    charts: sheet.charts.map(chart => chartIn(chart, sheet, nextChartId())),
    notes: sheet.notes,
    frozenRows: sheet.frozenRows,
    frozenColumns: sheet.frozenColumns,
    hiddenRows: [...sheet.hiddenRows],
    rowHeights: [...sheet.rowHeights].map(([row, share]) => [
      row,
      Math.min(Math.max(Math.round(share * ROW_HEIGHT), MIN_ROW_HEIGHT), MAX_ROW_HEIGHT)
    ]),
    columnWidths
  };
}

/** What opening it did, as the status line says it. */
export function reportOfXlsx(fileName: string, book: XlsxBook): string {
  const sheets = book.sheets.length;
  const parts = [`Opened ${fileName}: ${sheets} ${sheets === 1 ? 'sheet' : 'sheets'}`];
  if (book.valuesKept > 0) {
    parts.push(
      `${book.valuesKept} ${book.valuesKept === 1 ? 'formula' : 'formulas'} this sheet cannot run kept as ${
        book.valuesKept === 1 ? 'its value' : 'their values'
      }`
    );
  }
  if (book.cut.rows > 0 || book.cut.columns > 0) {
    const lost: string[] = [];
    if (book.cut.rows > 0) {
      lost.push(`${book.cut.rows.toLocaleString('en-US')} ${book.cut.rows === 1 ? 'row' : 'rows'}`);
    }
    if (book.cut.columns > 0) {
      lost.push(`${book.cut.columns.toLocaleString('en-US')} ${book.cut.columns === 1 ? 'column' : 'columns'}`);
    }
    parts.push(`${lost.join(' and ')} past the sheet's edge left out`);
  }
  for (const [what, count] of Object.entries(book.leftOut)) {
    if (count > 0) {
      parts.push(`${count} ${count === 1 ? what.replace(/s$/, '') : what} this sheet cannot keep left out`);
    }
  }
  return `${parts.join('; ')}.`;
}

/** Inflate, as the application worker has it: the platform's own. */
export const platformInflate: Inflate = async bytes => {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

/**
 * A chart from a file, as a chart here: its range, and its anchor turned
 * into pixels from the top-left of A1 with the sheet's own widths and
 * heights — the anchor is in cells, and a chart here is a rectangle.
 */
function chartIn(chart: XlsxChart, sheet: XlsxSheet, id: number): Chart {
  const width = (column: number) => sheet.columnWidths.get(column) ?? COLUMN_WIDTH;
  const height = (row: number) =>
    sheet.hiddenRows.includes(row) ? 0 : Math.round((sheet.rowHeights.get(row) ?? 1) * ROW_HEIGHT);
  const left = (column: number) => {
    let x = 0;
    for (let at = 0; at < column; at++) {
      x += width(at);
    }
    return x;
  };
  const top = (row: number) => {
    let y = 0;
    for (let at = 0; at < row; at++) {
      y += height(at);
    }
    return y;
  };
  const anchor = chart.anchor;
  let place: Chart['place'];
  if (anchor.kind === 'pixels') {
    place = { x: anchor.x, y: anchor.y, width: anchor.width, height: anchor.height };
  } else {
    const x = left(anchor.fromColumn) + anchor.fromColumnOffset;
    const y = top(anchor.fromRow) + anchor.fromRowOffset;
    // A one-cell anchor carries its size rather than a far corner.
    const right = anchor.toColumn < 0 ? x + anchor.toColumnOffset : left(anchor.toColumn) + anchor.toColumnOffset;
    const bottom = anchor.toRow < 0 ? y + anchor.toRowOffset : top(anchor.toRow) + anchor.toRowOffset;
    place = { x, y, width: right - x, height: bottom - y };
  }
  return {
    id,
    kind: chart.kind,
    title: chart.title,
    legend: chart.legend,
    range: onSheet(rangeOfRect(chart.firstRow, chart.firstColumn, chart.lastRow, chart.lastColumn), chart.sheet),
    place: {
      x: Math.max(0, place.x),
      y: Math.max(0, place.y),
      width: Math.max(MIN_CHART_WIDTH, place.width),
      height: Math.max(MIN_CHART_HEIGHT, place.height)
    }
  };
}

/** A range read from another sheet, naming it, as a formula's would. */
function onSheet(range: Chart['range'], sheet: string | undefined): Chart['range'] {
  return sheet === undefined ? range : { ...range, start: { ...range.start, sheet } };
}
