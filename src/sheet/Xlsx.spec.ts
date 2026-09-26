import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { numberFormatOf, openXlsx, readXlsx, XlsxError, type XlsxBook, type XlsxSheet } from './Xlsx';
import type { Inflate } from './Zip';

/**
 * Reading an Excel workbook.
 *
 * Two kinds of evidence. `fixtures/orders.xlsx` is a real file, written
 * by LibreOffice's own Excel filter from `scripts/xlsx-fixtures.py`, so
 * what is asserted against it is what a person exporting from another
 * spreadsheet would hand this one. The rest hand the reader XML parts
 * directly, for the shapes that file does not happen to contain —
 * Excel's shared formulas, a frozen pane, a function this sheet lacks.
 */

const HERE = dirname(fileURLToPath(import.meta.url));
const LIMITS = { rows: 10_000, columns: 100 };

const inflate: Inflate = async bytes => {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

function cellAt(sheet: XlsxSheet, address: string) {
  const column = address.charCodeAt(0) - 65;
  const row = Number(address.slice(1)) - 1;
  return sheet.cells.find(cell => cell.row === row && cell.column === column);
}

function formatAt(sheet: XlsxSheet, address: string) {
  const cell = cellAt(sheet, address);
  return cell === undefined ? undefined : sheet.formats[cell.style];
}

describe('a workbook LibreOffice wrote', () => {
  let book: XlsxBook;
  const orders = () => book.sheets[0];

  it('opens, with its sheets in order', async () => {
    book = await openXlsx(new Uint8Array(readFileSync(join(HERE, 'fixtures', 'orders.xlsx'))), inflate, LIMITS);
    expect(book.sheets.map(sheet => sheet.name)).toEqual(['Orders', 'Rates']);
  });

  it('keeps text, numbers and formulas as what would be typed', () => {
    expect(cellAt(orders(), 'A1')?.input).toBe('Quarter orders');
    expect(cellAt(orders(), 'B4')?.input).toBe('3');
    expect(cellAt(orders(), 'D4')?.input).toBe('=B4*C4');
    expect(cellAt(orders(), 'D8')?.input).toBe('=SUM(D4:D7)');
    expect(cellAt(orders(), 'D9')?.input).toBe('=D8*TaxRate');
    expect(cellAt(orders(), 'D11')?.input).toBe('=D8*Rates!B1');
    expect(cellAt(orders(), 'G4')?.input).toBe('=TRUE()');
  });

  /**
   * The fixture's due dates are `=DATEVALUE("2026-09-24")`, and this
   * sheet has no DATEVALUE. So the four of them keep the serial
   * LibreOffice calculated — which, under their date format, still
   * shows the date — and the count says four formulas became values.
   */
  it('keeps the value of each formula calling a function this sheet lacks', () => {
    expect(cellAt(orders(), 'E4')).toMatchObject({ input: '46289', asText: false });
    expect(book.valuesKept).toBe(4);
  });

  /** A leading zero is the commonest thing a spreadsheet import destroys. */
  it('keeps text that looks like a number as text', () => {
    expect(cellAt(orders(), 'F4')).toMatchObject({ input: '007', asText: true });
    expect(cellAt(orders(), 'A4')).toMatchObject({ input: 'Widget', asText: false });
  });

  it('reads the number formats it used', () => {
    expect(formatAt(orders(), 'C4')?.number).toEqual({ kind: 'currency', places: 2, symbol: '$' });
    expect(formatAt(orders(), 'E4')?.number).toEqual({ kind: 'date', pattern: 'ymd' });
    expect(formatAt(orders(), 'D10')?.number).toEqual({ kind: 'percent', places: 1 });
    expect(formatAt(orders(), 'F4')?.number).toEqual({ kind: 'text' });
  });

  it('reads bold, a font colour and a fill', () => {
    expect(formatAt(orders(), 'A1')?.paint).toMatchObject({ bold: true, color: '#ffffff', fill: '#1f3864' });
    expect(formatAt(orders(), 'A3')?.paint).toMatchObject({ bold: true, fill: '#d9d9d9' });
    expect(formatAt(orders(), 'A4')?.paint).toMatchObject({ bold: false, fill: '', fontSize: 0 });
  });

  /** The title's fill runs across the merge, over cells that hold nothing. */
  it('keeps the style of a cell with nothing in it', () => {
    expect(orders().styled).toContainEqual({ row: 0, column: 4, style: cellAt(orders(), 'A1')!.style });
  });

  it('keeps the merge, the hidden row, the widths and the name', () => {
    expect(orders().merges).toEqual([{ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 }]);
    expect(orders().hiddenRows).toEqual([11]);
    expect(orders().columnWidths.get(0)).toBeGreaterThan(orders().columnWidths.get(4)!);
    expect(book.names).toEqual([{ name: 'TaxRate', sheet: 'Rates', firstRow: 0, firstColumn: 1, lastRow: 0, lastColumn: 1 }]);
  });
});

/** A workbook of one sheet, from the sheet's XML and whatever else a spec needs. */
function book(sheet: string, extra: Record<string, string> = {}, limits = LIMITS): XlsxBook {
  const parts: Record<string, string> = {
    'xl/workbook.xml': '<workbook><sheets><sheet name="One" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': `<worksheet>${sheet}</worksheet>`,
    ...extra
  };
  return readXlsx(path => parts[path] ?? null, limits);
}

const inputs = (read: XlsxBook) => read.sheets[0].cells.map(cell => [cell.row, cell.column, cell.input]);

describe('the shapes Excel writes', () => {
  it('moves a shared formula to each cell that shares it', () => {
    const read = book(`<sheetData>
      <row r="4"><c r="D4"><f t="shared" ref="D4:D6" si="0">B4*C4</f><v>1</v></c></row>
      <row r="5"><c r="D5"><f t="shared" si="0"/><v>2</v></c></row>
      <row r="6"><c r="D6"><f t="shared" si="0"/><v>3</v></c></row></sheetData>`);
    // Moved by the same rewrite a fill uses, and printed the way it
    // prints: the brackets are its, and they change nothing.
    expect(inputs(read)).toEqual([
      [3, 3, '=B4*C4'],
      [4, 3, '=(B5*C5)'],
      [5, 3, '=(B6*C6)']
    ]);
  });

  it('drops the prefix Excel puts on newer functions', () => {
    const read = book('<sheetData><row r="1"><c r="A1"><f>_xlfn.IFS(B1&gt;1,"big",TRUE,"small")</f><v>0</v></c></row></sheetData>');
    expect(inputs(read)).toEqual([[0, 0, '=IFS(B1>1,"big",TRUE,"small")']]);
  });

  /** A formula that silently became a number is a sheet that stops updating, so it is counted. */
  it('keeps the value of a formula it cannot run, and counts it', () => {
    const read = book(`<sheetData><row r="1">
      <c r="A1"><f>BESSELJ(1.9,2)</f><v>0.3299</v></c>
      <c r="B1" t="str"><f>[1]Other!A1</f><v>from elsewhere</v></c></row></sheetData>`);
    expect(inputs(read)).toEqual([
      [0, 0, '0.3299'],
      [0, 1, 'from elsewhere']
    ]);
    expect(read.valuesKept).toBe(2);
  });

  it('reads a frozen pane, and ignores a split that is not frozen', () => {
    const frozen = book('<sheetViews><sheetView><pane xSplit="1" ySplit="3" state="frozen"/></sheetView></sheetViews><sheetData/>');
    expect([frozen.sheets[0].frozenRows, frozen.sheets[0].frozenColumns]).toEqual([3, 1]);
    const split = book('<sheetViews><sheetView><pane ySplit="2000" state="split"/></sheetView></sheetViews><sheetData/>');
    expect(split.sheets[0].frozenRows).toBe(0);
  });

  it('reads inline strings, booleans and errors', () => {
    const read = book(`<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t>typed here</t></is></c>
      <c r="B1" t="b"><v>0</v></c>
      <c r="C1" t="e"><v>#N/A</v></c></row></sheetData>`);
    expect(inputs(read)).toEqual([
      [0, 0, 'typed here'],
      [0, 1, 'FALSE'],
      [0, 2, '#N/A']
    ]);
  });

  it('joins a rich string’s runs and leaves out its phonetic guide', () => {
    const read = book('<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>', {
      'xl/sharedStrings.xml':
        '<sst><si><r><rPr><b/></rPr><t>Tan</t></r><r><t xml:space="preserve">aka Taro</t></r><rPh><t>タナカ</t></rPh></si></sst>'
    });
    expect(inputs(read)).toEqual([[0, 0, 'Tanaka Taro']]);
  });

  it('cuts what is past the sheet’s edge, and says how much', () => {
    const read = book(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="E1"><v>2</v></c></row><row r="7"><c r="A7"><v>3</v></c></row></sheetData>',
      {},
      { rows: 5, columns: 3 }
    );
    expect(inputs(read)).toEqual([[0, 0, '1']]);
    expect(read.cut).toEqual({ rows: 2, columns: 2 });
  });

  it('takes a hidden column as one of no width', () => {
    const read = book('<cols><col min="2" max="3" hidden="1"/><col min="4" max="4" width="20" customWidth="1"/></cols><sheetData/>');
    const widths = read.sheets[0].columnWidths;
    expect([widths.get(1), widths.get(2), widths.get(3)]).toEqual([0, 0, 140]);
  });

  it('refuses a file that is not a workbook, with a sentence', async () => {
    await expect(openXlsx(new TextEncoder().encode('Region,Units\n'), inflate, LIMITS)).rejects.toThrow(XlsxError);
    expect(() => readXlsx(() => null, LIMITS)).toThrow('there is no xl/workbook.xml in it');
  });
});

describe('number format codes', () => {
  it.each([
    ['General', { kind: 'general' }],
    ['0', { kind: 'number', places: 0, thousands: false }],
    ['#,##0.00', { kind: 'number', places: 2, thousands: true }],
    ['0.0%', { kind: 'percent', places: 1 }],
    ['0.00E+00', { kind: 'scientific', places: 2 }],
    ['"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', { kind: 'number', places: 2, thousands: true }],
    ['$#,##0.00', { kind: 'currency', places: 2, symbol: '$' }],
    ['[$€-407]#,##0.00', { kind: 'currency', places: 2, symbol: '€' }],
    ['yyyy\\-mm\\-dd', { kind: 'date', pattern: 'ymd' }],
    ['dd/mm/yyyy', { kind: 'date', pattern: 'dmy' }],
    ['m/d/yyyy', { kind: 'date', pattern: 'mdy' }],
    ['h:mm AM/PM', { kind: 'time', pattern: 'hm' }],
    ['[h]:mm:ss', { kind: 'time', pattern: 'hms' }],
    ['m/d/yyyy h:mm', { kind: 'datetime', date: 'mdy', time: 'hm' }],
    ['@', { kind: 'text' }]
  ])('reads %s', (code, expected) => {
    expect(numberFormatOf(code)).toEqual(expected);
  });
});
