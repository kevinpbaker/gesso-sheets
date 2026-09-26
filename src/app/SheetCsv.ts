import { MAX_SHEETS } from '../sheet/A1';
import { looksLikeFormula, parseCsv, toCsv } from '../sheet/Csv';
import type { Block } from '../sheet/Tsv';
import type { SheetDocument } from './SheetDocument';

/**
 * A CSV into a sheet of its own, and a sheet back out as one.
 *
 * Beside `SheetRanges.ts` and for the same reason: these are edits
 * over many cells that the service only has to call, and a spec can
 * drive them against a bare document without a channel in sight.
 */

export interface CsvImport {
  /** The sheet it landed on, or -1 when it could not be added. */
  readonly sheet: number;
  /** A sentence about what happened, for the person who dropped it. */
  readonly report: string;
}

/**
 * Writes a CSV onto a new sheet named after its file, and shows it.
 *
 * **A new sheet, never the one in view.** A dropped file overwriting
 * whatever the cursor was on would be the most destructive thing a
 * drag could do, and a sheet of its own is also what makes the import
 * one gesture to take back: delete the tab.
 *
 * **A cell starting `=` is written as text**, formatted Text so that
 * it stays text through an edit and a reload — the same path a person
 * takes to keep `007` from becoming seven. A CSV is data, and a
 * spreadsheet that ran the formulas in one is a spreadsheet anybody
 * can put a program into by sending a file. Everything else goes
 * through the ordinary literal path, so numbers are numbers and a
 * date is a date, as they would be if the same text were typed.
 *
 * A file bigger than the sheet is cut at the sheet's edges and the
 * report says by how much, which is better than refusing a file
 * because its thousand-and-first column is empty.
 */
export function importCsv(
  document: SheetDocument,
  fileName: string,
  text: string,
  size: { rows: number; columns: number }
): CsvImport {
  if (document.sheetCount >= MAX_SHEETS) {
    return { sheet: -1, report: `${fileName} was not opened: a workbook holds at most ${MAX_SHEETS} sheets.` };
  }
  const { rows } = parseCsv(text);
  const width = rows[0]?.length ?? 0;
  const keptRows = Math.min(rows.length, size.rows);
  const keptColumns = Math.min(width, size.columns);

  const sheet = document.addSheet(sheetNameOf(fileName));
  document.transact(() => {
    for (let row = 0; row < keptRows; row++) {
      const line = rows[row];
      for (let column = 0; column < keptColumns; column++) {
        const cell = line[column];
        if (cell === '') {
          continue;
        }
        if (cell.startsWith('=')) {
          document.setFormat(row, column, { ...document.formatAt(row, column), number: { kind: 'text' } });
        }
        document.setCell(row, column, cell);
      }
    }
  });
  // Opening a file is not an edit to take back one cell at a time;
  // taking it back is closing the tab, as it is for adding a sheet.
  document.forgetHistory();

  const cut: string[] = [];
  if (rows.length > keptRows) {
    cut.push(`${count(rows.length - keptRows)} ${plural(rows.length - keptRows, 'row')}`);
  }
  if (width > keptColumns) {
    cut.push(`${count(width - keptColumns)} ${plural(width - keptColumns, 'column')}`);
  }
  const read = `${count(keptRows)} ${plural(keptRows, 'row')} from ${fileName}`;
  return {
    sheet,
    report:
      cut.length === 0
        ? `Opened ${read}.`
        : `Opened ${read}; ${cut.join(' and ')} did not fit on the sheet and were left out.`
  };
}

/** `Q3 sales.csv` is a sheet called `Q3 sales`. */
function sheetNameOf(fileName: string): string {
  const base = fileName.replace(/\.[^.]*$/, '').trim();
  return base === '' ? 'Imported' : base;
}

function count(value: number): string {
  return value.toLocaleString('en-US');
}

function plural(value: number, word: string): string {
  return value === 1 ? word : `${word}s`;
}

/**
 * The sheet in view, as a CSV of what it shows.
 *
 * **What it shows, not what it holds**: `$1,234.50` and `22 Sep 2026`
 * rather than `1234.5` and `46287`, because the person exporting is
 * looking at the first and would not recognise the second — which is
 * also what Excel and Sheets both write. A formula's result goes out
 * and the formula does not; a CSV has nowhere to put one.
 *
 * From A1 to the last row and column anything is in, so a sheet that
 * starts at C5 keeps its blank margin and lands where it was when it
 * is read back.
 *
 * A *text* value that another program would run is written with an
 * apostrophe in front of it; see `looksLikeFormula`. A number that
 * happens to be negative is not text and is left alone.
 */
export function exportCsv(document: SheetDocument): string {
  let lastRow = -1;
  let lastColumn = -1;
  for (const { row, column } of document.sheet.entries()) {
    lastRow = Math.max(lastRow, row);
    lastColumn = Math.max(lastColumn, column);
  }
  const rows: string[][] = [];
  for (let row = 0; row <= lastRow; row++) {
    const line: string[] = [];
    for (let column = 0; column <= lastColumn; column++) {
      const shown = document.display(row, column);
      const value = document.sheet.value(row, column);
      line.push(typeof value === 'string' && looksLikeFormula(shown) ? `'${shown}` : shown);
    }
    rows.push(line);
  }
  return toCsv(rows satisfies Block);
}
