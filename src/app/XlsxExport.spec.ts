import { describe, expect, it } from 'vitest';

import { GENERAL, PLAIN, type CellFormat } from '../sheet/Format';
import { openXlsx } from '../sheet/Xlsx';
import { excelFormula, xlsxParts, writeXlsx } from '../sheet/XlsxWrite';
import { formatValue } from '../sheet/Values';
import { zipEntries } from '../sheet/Zip';
import { applySnapshot } from './SheetFile';
import { SheetDocument } from './SheetDocument';
import { seed } from './SheetSeed';
import { platformInflate, snapshotOfXlsx } from './SheetXlsx';
import { platformDeflate, xlsxOfDocument } from './SheetXlsxOut';

/**
 * A workbook out as an `.xlsx` and back in through the reader.
 *
 * The reader was tested against files LibreOffice and Excel wrote, so
 * what it reads back is a fair account of what another program will
 * read — and anything the round trip changes is either something the
 * format cannot say or a bug here.
 */

const ROWS = 200;

async function roundTrip(document: SheetDocument, deflate = true): Promise<SheetDocument> {
  const { book } = xlsxOfDocument(document, ROWS);
  const bytes = await writeXlsx(book, deflate ? platformDeflate : undefined);
  const read = await openXlsx(bytes, platformInflate, { rows: ROWS, columns: 20 });
  const back = new SheetDocument();
  applySnapshot(back, snapshotOfXlsx(read, 20));
  back.sheet.recalculate();
  return back;
}

function everything(document: SheetDocument): { row: number; column: number; input: string; shown: string }[][] {
  const sheets = [];
  for (let index = 0; index < document.sheetCount; index++) {
    document.activate(index);
    // A formula as written, and every cell as it shows: a number typed
    // `48.00` is the number 48 in a file, and shows as it did because
    // the format is what says two places.
    const cells = [...document.sheet.entries()]
      .map(cell => ({
        row: cell.row,
        column: cell.column,
        input: cell.input.startsWith('=') ? cell.input : '',
        shown: document.display(cell.row, cell.column)
      }))
      .sort((a, b) => a.row - b.row || a.column - b.column);
    sheets.push(cells);
  }
  document.activate(0);
  return sheets;
}

describe('a workbook written as an .xlsx', () => {
  it('reads back as the workbook it was: cells, formulas and what they show', async () => {
    const document = new SheetDocument();
    seed(document);
    document.sheet.recalculate();
    const back = await roundTrip(document);
    expect(everything(back)).toEqual(everything(document));
    expect(back.book.sheetNames()).toEqual(document.book.sheetNames());
  });

  it('keeps formats, widths, heights, merges, panes and names', async () => {
    const document = new SheetDocument();
    seed(document);
    document.sheet.recalculate();
    document.rowHeights.set(3, 48);
    document.fittedRows.set(5, 36);
    document.hiddenRows.add(20);
    const back = await roundTrip(document);
    for (let index = 0; index < document.sheetCount; index++) {
      const before = document.pageAt(index)!;
      const after = back.pageAt(index)!;
      for (const cell of before.sheet.entries()) {
        expect(after.formats.formatAt(cell.row, cell.column)).toEqual(before.formats.formatAt(cell.row, cell.column));
      }
      expect(after.columnWidths.slice(0, before.columnWidths.length)).toEqual(before.columnWidths);
      expect(after.merges.all).toEqual(before.merges.all);
      expect([after.frozenRows, after.frozenColumns]).toEqual([before.frozenRows, before.frozenColumns]);
    }
    // A height set by hand is one Excel keeps; a fitted one this sheet
    // fits again, as it does for a file from Excel.
    expect(back.pageAt(0)!.rowHeights.get(3)).toBe(48);
    expect([...back.pageAt(0)!.hiddenRows]).toContain(20);
    expect(back.book.names.all().map(name => name.name)).toEqual(document.book.names.all().map(name => name.name));
  });

  it('is a zip of the parts Excel expects, deflated or stored', async () => {
    const document = new SheetDocument();
    document.setCell(0, 0, 'hello');
    const { book } = xlsxOfDocument(document, ROWS);
    for (const deflate of [platformDeflate, undefined]) {
      const names = zipEntries(await writeXlsx(book, deflate)).map(entry => entry.name);
      expect(names).toEqual([
        '[Content_Types].xml',
        '_rels/.rels',
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/styles.xml',
        'xl/worksheets/sheet1.xml'
      ]);
    }
  });

  it('says what it cannot carry', () => {
    const document = new SheetDocument();
    seed(document);
    expect(xlsxOfDocument(document, ROWS).leftOut.sort()).toEqual(['conditional formats', 'validations']);
  });

  it('writes an array that spills over the cells it fills, and back', async () => {
    const document = new SheetDocument();
    for (const [row, value] of ['1', '2', '3'].entries()) {
      document.setCell(row, 0, value);
    }
    document.setCell(0, 2, '=A1:A3*2');
    document.setCell(0, 3, '=SUM(A1:A3*A1:A3)');
    document.setCell(1, 3, '=@A1:A3*10');
    document.sheet.recalculate();
    const sheet = xlsxParts(xlsxOfDocument(document, ROWS).book).find(part => part.name === 'xl/worksheets/sheet1.xml')!.text;
    expect(sheet).toContain('<f t="array" ref="C1:C3">A1:A3*2</f>');
    // Arithmetic over ranges that does not spill is an array formula
    // too, or an older Excel would take one value from each range.
    expect(sheet).toContain('<f t="array" ref="D1">SUM(A1:A3*A1:A3)</f>');
    // And one value is Excel's default in a plain formula.
    expect(sheet).toContain('<f>A1:A3*10</f>');

    const back = await roundTrip(document);
    expect([0, 1, 2].map(row => back.sheet.value(row, 2))).toEqual([2, 4, 6]);
    expect(back.sheet.value(0, 3)).toBe(14);
    expect(back.sheet.value(1, 3)).toBe(20);
  });

  it('keeps text that looks like something else as text', async () => {
    const document = new SheetDocument();
    const text: CellFormat = { number: { kind: 'text' }, paint: PLAIN };
    document.setFormat(0, 0, text);
    document.setCell(0, 0, '007');
    document.setCell(1, 0, 'TRUE');
    document.setCell(2, 0, 'a < b & "c"');
    const back = await roundTrip(document);
    expect(formatValue(back.sheet.value(0, 0))).toBe('007');
    expect(back.sheet.value(2, 0)).toBe('a < b & "c"');
    expect(back.formatAt(0, 0).number).toEqual({ kind: 'text' });
    expect(back.formatAt(5, 5).number).toEqual(GENERAL);
  });
});

describe('a formula as Excel writes it', () => {
  it('prefixes the functions newer than the format', () => {
    expect(excelFormula('=XLOOKUP(A1, B:B, C:C)', false)).toBe('_xlfn.XLOOKUP(A1, B:B, C:C)');
    expect(excelFormula('=FILTER(A1:A9, A1:A9>1)', true)).toBe('_xlfn._xlws.FILTER(A1:A9, A1:A9>1)');
    expect(excelFormula('=SUM(A1:A9)', false)).toBe('SUM(A1:A9)');
    expect(excelFormula('="XLOOKUP(" & A1', false)).toBe('"XLOOKUP(" & A1');
  });

  it('writes A1# as the call Excel’s files spell it with', () => {
    expect(excelFormula('=SUM(C1#)', false)).toBe('SUM(_xlfn.ANCHORARRAY(C1))');
    expect(excelFormula("=SUM('My Data'!B2#)", false)).toBe("SUM(_xlfn.ANCHORARRAY('My Data'!B2))");
  });

  it('writes @ as SINGLE in an array formula, and leaves it out of a plain one', () => {
    expect(excelFormula('=@A1:A9*B1:B9', true)).toBe('_xlfn.SINGLE(A1:A9)*B1:B9');
    expect(excelFormula('=@(A1:A9)*2', true)).toBe('_xlfn.SINGLE((A1:A9))*2');
    expect(excelFormula('=@A1:A9*2', false)).toBe('A1:A9*2');
  });
});
