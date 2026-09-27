import { DEFAULT_FORMAT, keyOf, type CellFormat } from '../sheet/Format';
import { columnName, quoteSheetName } from '../sheet/A1';
import type { Chart } from '../sheet/Chart';
import { layoutOf } from '../sheet/Series';
import type { CellValue } from '../sheet/Values';
import type { XlsxOutChart, XlsxOutSeries } from '../sheet/XlsxCharts';
import { writeValidations } from '../sheet/XlsxRules';
import type { XlsxOut, XlsxOutCell, XlsxOutFormulaName, XlsxOutName, XlsxOutRow, XlsxOutSheet } from '../sheet/XlsxWrite';
import { isNamedRange } from '../sheet/Names';
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

    // Written since Phase 23; only a rule the format cannot say — a
    // list longer than Excel allows, or one with a comma in a value —
    // is left out, and then it is said.
    if (writeValidations(page.validations).unwritten > 0) {
      leftOut.add('some validations');
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
      notes: page.notes.all().filter(note => note.row < rowCount),
      validations: [...page.validations],
      conditional: [...page.conditional],
      charts: page.charts.map(chart => {
        // The sheet the chart reads, which is its own unless it names another.
        const named = chart.range.start.sheet;
        const at = named === undefined ? index : document.book.sheetFor(named);
        const source = at === null ? null : document.book.sheet(at);
        return chartOut(chart, named ?? names[index] ?? `Sheet${index + 1}`, (row, column) => source?.value(row, column) ?? null);
      })
    });
  }

  const definedNames: (XlsxOutName | XlsxOutFormulaName)[] = [];
  for (const entry of document.book.names.all()) {
    if (!isNamedRange(entry)) {
      // A formula goes out as written; the writer spells its functions
      // and parameters as Excel's files do.
      definedNames.push({ name: entry.name, formula: entry.formula });
      continue;
    }
    // A name written without a sheet means the sheet of the formula
    // that reads it, where an Excel name belongs to one sheet. So it
    // goes out on the sheet whose formulas use it — the first, when
    // none does.
    const sheet = entry.range.start.sheet ?? names[sheetUsing(document, entry.name)] ?? names[0];
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

/** The first sheet with a formula that says a name, by index; 0 when none does. */
function sheetUsing(document: SheetDocument, name: string): number {
  const said = new RegExp(`(^|[^A-Za-z0-9_.!$])${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^A-Za-z0-9_.(])`, 'i');
  for (let index = 0; index < document.sheetCount; index++) {
    const page = document.pageAt(index);
    if (page === undefined) {
      continue;
    }
    for (const cell of page.sheet.entries()) {
      if (cell.input.startsWith('=') && said.test(cell.input.replace(/"[^"]*"/g, '""'))) {
        return index;
      }
    }
  }
  return 0;
}

/** The platform's deflate, for the zip: the other half of `platformInflate`. */
export const platformDeflate: Deflate = async bytes => {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

/**
 * A chart as the file writes it: its series spelled out as references,
 * read the way the chart on screen reads its range — `layoutOf`, which
 * both share, so the file charts the numbers somebody was looking at.
 */
function chartOut(chart: Chart, sheet: string, value: (row: number, column: number) => CellValue): XlsxOutChart {
  const firstRow = Math.min(chart.range.start.row, chart.range.end.row);
  const lastRow = Math.max(chart.range.start.row, chart.range.end.row);
  const firstColumn = Math.min(chart.range.start.column, chart.range.end.column);
  const lastColumn = Math.max(chart.range.start.column, chart.range.end.column);
  const grid: CellValue[][] = [];
  for (let row = firstRow; row <= lastRow; row++) {
    const line: CellValue[] = [];
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(value(row, column));
    }
    grid.push(line);
  }
  const { byColumn, headers, labels } = layoutOf(grid);
  const prefix = `${quoteSheetName(sheet)}!`;
  const cell = (row: number, column: number) => `$${columnName(column)}$${row + 1}`;
  const span = (r1: number, c1: number, r2: number, c2: number) => `${prefix}${cell(r1, c1)}:${cell(r2, c2)}`;
  const series: XlsxOutSeries[] = [];
  if (byColumn) {
    const dataFrom = firstRow + (headers ? 1 : 0);
    for (let column = firstColumn + (labels ? 1 : 0); column <= lastColumn; column++) {
      series.push({
        ...(headers ? { name: `${prefix}${cell(firstRow, column)}` } : {}),
        ...(labels ? { categories: span(dataFrom, firstColumn, lastRow, firstColumn) } : {}),
        values: span(dataFrom, column, lastRow, column)
      });
    }
  } else {
    const dataFrom = firstColumn + (headers ? 1 : 0);
    for (let row = firstRow + (labels ? 1 : 0); row <= lastRow; row++) {
      series.push({
        ...(headers ? { name: `${prefix}${cell(row, firstColumn)}` } : {}),
        ...(labels ? { categories: span(firstRow, dataFrom, firstRow, lastColumn) } : {}),
        values: span(row, dataFrom, row, lastColumn)
      });
    }
  }
  return { kind: chart.kind, title: chart.title, legend: chart.legend, series, place: chart.place };
}
