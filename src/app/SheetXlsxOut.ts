import { DEFAULT_FORMAT, keyOf, type CellFormat } from '../sheet/Format';
import type { XlsxOut, XlsxOutCell, XlsxOutName, XlsxOutRow, XlsxOutSheet } from '../sheet/XlsxWrite';
import type { Deflate } from '../sheet/Zip';
import { ROW_HEIGHT } from './dimensions';
import type { SheetDocument } from './SheetDocument';

/**
 * A document as the records `XlsxWrite` writes: the other half of
 * `snapshotOfXlsx`, and its inverse wherever the two meet.
 *
 * Widths go out in Excel's characters by the rule the reader brings
 * them in by, and heights as a share of the default row, so a workbook
 * that came from Excel goes back with the widths and heights it came
 * with. Values are this sheet's answers; Excel recalculates on opening.
 *
 * What does not go: conditional formats, validations and charts, which
 * the reader does not bring in either. `leftOut` says which of them the
 * document had, for the line that reports the download.
 */
export function xlsxOfDocument(document: SheetDocument, rowCount: number): { book: XlsxOut; leftOut: string[] } {
  const formats: CellFormat[] = [DEFAULT_FORMAT];
  const ids = new Map<string, number>([[keyOf(DEFAULT_FORMAT), 0]]);
  const styleOf = (format: CellFormat): number => {
    const key = keyOf(format);
    let id = ids.get(key);
    if (id === undefined) {
      id = formats.length;
      formats.push(format);
      ids.set(key, id);
    }
    return id;
  };

  const sheets: XlsxOutSheet[] = [];
  const leftOut = new Set<string>();
  const names = document.book.sheetNames();
  for (let index = 0; index < document.sheetCount; index++) {
    const page = document.pageAt(index);
    if (page === undefined) {
      continue;
    }
    if (page.conditional.length > 0) {
      leftOut.add('conditional formats');
    }
    if (page.validations.length > 0) {
      leftOut.add('validations');
    }
    if (page.charts.length > 0) {
      leftOut.add('charts');
    }
    const cells = new Map<string, XlsxOutCell>();
    const put = (cell: XlsxOutCell): void => {
      cells.set(`${cell.row}:${cell.column}`, cell);
    };
    const format = (row: number, column: number): number => styleOf(page.formats.formatAt(row, column));

    for (const entry of page.sheet.entries()) {
      if (entry.row >= rowCount) {
        continue;
      }
      const spill = page.sheet.spillOf(entry.row, entry.column);
      put({
        row: entry.row,
        column: entry.column,
        input: entry.input,
        value: page.sheet.value(entry.row, entry.column),
        style: format(entry.row, entry.column),
        ...(spill === null ? {} : { spill })
      });
      // The cells an array fills are written with their values, as
      // Excel writes them: the formula is the anchor's alone.
      if (spill !== null) {
        for (let r = 0; r < spill.rows; r++) {
          for (let c = 0; c < spill.columns; c++) {
            if (r === 0 && c === 0) {
              continue;
            }
            const row = entry.row + r;
            const column = entry.column + c;
            put({ row, column, input: '', value: page.sheet.value(row, column), style: format(row, column) });
          }
        }
      }
    }
    // Cells with a format and nothing in them: the fill across a
    // header's blank cells, a border round an empty box.
    for (const cell of page.formats.compact().cells) {
      if (cell.row < rowCount && !cells.has(`${cell.row}:${cell.column}`)) {
        put({ row: cell.row, column: cell.column, input: '', value: null, style: format(cell.row, cell.column) });
      }
    }

    const columnWidths = new Map<number, number>();
    page.columnWidths.forEach((pixels, column) => {
      // The reader's `pixelsOf`, the other way: a width of w characters
      // is drawn 7w pixels wide, near enough that this round-trips.
      columnWidths.set(column, pixels <= 0 ? 0 : pixels / 7);
    });

    const rows = new Map<number, XlsxOutRow>();
    const shape = (row: number): { height?: number; custom: boolean; hidden: boolean } => {
      const current = rows.get(row);
      return current === undefined ? { custom: false, hidden: false } : { ...current };
    };
    // Points, as a share of the default row: 15 points is a row this
    // sheet draws `ROW_HEIGHT` tall, which is how the reader reads them.
    for (const [row, pixels] of page.fittedRows) {
      if (row < rowCount) {
        rows.set(row, { ...shape(row), height: (pixels / ROW_HEIGHT) * 15 });
      }
    }
    for (const [row, pixels] of page.rowHeights) {
      if (row < rowCount) {
        rows.set(row, { ...shape(row), height: (pixels / ROW_HEIGHT) * 15, custom: true });
      }
    }
    for (const row of [...page.hiddenRows, ...page.filteredRows]) {
      if (row < rowCount) {
        rows.set(row, { ...shape(row), hidden: true });
      }
    }

    sheets.push({
      name: names[index] ?? `Sheet${index + 1}`,
      cells: [...cells.values()],
      columnWidths,
      rows,
      merges: page.merges.all.filter(rect => rect.lastRow < rowCount),
      frozenRows: page.frozenRows,
      frozenColumns: page.frozenColumns,
      notes: page.notes.all().filter(note => note.row < rowCount)
    });
  }

  const definedNames: XlsxOutName[] = [];
  for (const entry of document.book.names.all()) {
    const sheet = entry.range.start.sheet ?? names[0];
    if (sheet === undefined) {
      continue;
    }
    definedNames.push({
      name: entry.name,
      sheet,
      firstRow: Math.min(entry.range.start.row, entry.range.end.row),
      firstColumn: Math.min(entry.range.start.column, entry.range.end.column),
      lastRow: Math.max(entry.range.start.row, entry.range.end.row),
      lastColumn: Math.max(entry.range.start.column, entry.range.end.column)
    });
  }

  return {
    book: { sheets, formats, names: definedNames, active: document.active, iteration: document.book.iteration },
    leftOut: [...leftOut]
  };
}

/** The platform's deflate, for the zip: the other half of `platformInflate`. */
export const platformDeflate: Deflate = async bytes => {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};
