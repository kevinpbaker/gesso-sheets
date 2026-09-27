import { describe, expect, it } from 'vitest';

import { parseAddress } from '../sheet/A1';

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

  /**
   * A name written without a sheet means the sheet of the formula that
   * reads it; an Excel name belongs to one sheet. It used to go out on
   * the first sheet, which was right only while the name was used there.
   */
  it('puts a name on the sheet whose formulas use it', async () => {
    const document = new SheetDocument();
    document.renameSheet(0, 'Cover');
    document.activate(document.addSheet('Data'));
    document.setCell(0, 0, '2');
    document.setCell(1, 0, '3');
    document.defineName('Figures', parseAddress('A1:A2')!);
    document.setCell(2, 0, '=SUM(Figures)');
    document.sheet.recalculate();
    const workbook = xlsxParts(xlsxOfDocument(document, ROWS).book).find(part => part.name === 'xl/workbook.xml')!.text;
    expect(workbook).toContain('<definedName name="Figures">Data!$A$1:$A$2</definedName>');
    const back = await roundTrip(document);
    back.activate(1);
    expect(back.sheet.value(2, 0)).toBe(5);
  });

  /**
   * Phase 28. Excel writes a named LAMBDA as `_xlfn.LAMBDA(_xlpm.x, …)`
   * — the function marked as one from after the format, the parameter
   * as a parameter — and writes the same prefixes inside a cell's LET.
   */
  it('carries a named LAMBDA and a LET, spelt as Excel spells them', async () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '21');
    expect(document.defineFormulaName('Double', '=LAMBDA(x, x*2)')).toBeNull();
    document.setCell(0, 1, '=Double(A1)');
    document.setCell(0, 2, '=LET(total, A1+B1, total/3)');
    document.setCell(0, 3, '=SUM(MAP(A1:B1, LAMBDA(v, v+1)))');
    document.sheet.recalculate();
    const parts = xlsxParts(xlsxOfDocument(document, ROWS).book);
    const workbook = parts.find(part => part.name === 'xl/workbook.xml')!.text;
    expect(workbook).toContain('<definedName name="Double">_xlfn.LAMBDA(_xlpm.x, _xlpm.x*2)</definedName>');
    const sheet = parts.find(part => part.name === 'xl/worksheets/sheet1.xml')!.text;
    expect(sheet).toContain('_xlfn.LET(_xlpm.total, A1+B1, _xlpm.total/3)');
    expect(sheet).toContain('_xlfn.MAP(A1:B1, _xlfn.LAMBDA(_xlpm.v, _xlpm.v+1))');

    const back = await roundTrip(document);
    expect(back.sheet.names.all()).toEqual([expect.objectContaining({ name: 'Double', formula: '=LAMBDA(x, x*2)' })]);
    expect([0, 1, 2, 3].map(column => back.sheet.value(0, column))).toEqual([21, 42, 21, 65]);
    // Read back as formulas that run, not as the values Excel last saw.
    expect(back.sheet.input(0, 2)).toBe('=LET(total, A1+B1, total/3)');
  });

  it('carries an optional parameter in its brackets', async () => {
    const document = new SheetDocument();
    document.defineFormulaName('Taxed', '=LAMBDA(amount, [rate], amount*(1+IF(ISOMITTED(rate), 0.2, rate)))');
    document.setCell(0, 0, '=Taxed(100)');
    document.sheet.recalculate();
    const workbook = xlsxParts(xlsxOfDocument(document, ROWS).book).find(part => part.name === 'xl/workbook.xml')!.text;
    expect(workbook).toContain('_xlfn.LAMBDA(_xlpm.amount, [_xlpm.rate], _xlpm.amount*(1+IF(_xlfn.ISOMITTED(_xlpm.rate), 0.2, _xlpm.rate)))');
    const back = await roundTrip(document);
    expect(back.sheet.value(0, 0)).toBe(120);
  });

  it('says what it cannot carry', () => {
    const document = new SheetDocument();
    seed(document);
    // Validations and conditional formats are written since Phase 23,
    // and the seed's all fit: nothing it has is left out any more.
    expect(xlsxOfDocument(document, ROWS).leftOut).toEqual([]);
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

/**
 * Notes out as Excel's comments — after Phase 22, which read them in.
 *
 * Out and back through the reader, which was tested against files
 * Excel wrote; and the parts Excel needs before it will show one, which
 * the reader does not need and so could not catch the lack of.
 */
describe('notes in an exported workbook', () => {
  function noted(): SheetDocument {
    const document = new SheetDocument();
    document.setCell(0, 0, 'Region');
    document.setNote(0, 0, 'Where the order shipped from');
    document.setNote(4, 2, 'Two lines\nof note, & an ampersand <tag>');
    document.addSheet('Plain');
    return document;
  }

  it('come back as the notes they were', async () => {
    const back = await roundTrip(noted());
    back.activate(0);
    expect(back.noteAt(0, 0)).toBe('Where the order shipped from');
    expect(back.noteAt(4, 2)).toBe('Two lines\nof note, & an ampersand <tag>');
    back.activate(1);
    expect(back.notes.size).toBe(0);
  });

  it('carry the drawing Excel shows a comment through, and say so in the content types', () => {
    const parts = xlsxParts(xlsxOfDocument(noted(), ROWS).book);
    const named = (name: string) => parts.find(part => part.name === name)?.text ?? '';

    expect(named('xl/comments1.xml')).toContain('<comment ref="C5" authorId="0">');
    expect(named('xl/drawings/vmlDrawing1.vml').match(/<v:shape /g)?.length).toBe(2);
    expect(named('xl/worksheets/_rels/sheet1.xml.rels')).toContain('Target="../drawings/vmlDrawing1.vml"');
    expect(named('xl/worksheets/sheet1.xml')).toContain('<legacyDrawing r:id="rId2"/>');
    expect(named('[Content_Types].xml')).toContain('/xl/comments1.xml');
    expect(named('[Content_Types].xml')).toContain('Extension="vml"');
    // The sheet with no notes has none of it.
    expect(parts.some(part => part.name === 'xl/comments2.xml')).toBe(false);
    expect(named('xl/worksheets/sheet2.xml')).not.toContain('legacyDrawing');
  });
});

/**
 * Validations out and back — Phase 23.
 *
 * Every rule kind this sheet has, through the writer and the reader;
 * the one kind of rule Excel cannot say, left out and said; and a list
 * that names a range of cells rather than listing them, read from a
 * file the way Excel writes one.
 */
describe('validations in an exported workbook', () => {
  function ruled(): SheetDocument {
    const document = new SheetDocument();
    const at = (row: number, column: number, lastRow = row) => ({
      start: { row, column, rowAbsolute: false, columnAbsolute: false },
      end: { row: lastRow, column, rowAbsolute: false, columnAbsolute: false }
    });
    document.addValidation({ range: at(1, 0, 20), rule: { kind: 'list', values: ['North', 'South', 'Say "hi"'] }, strict: true, message: 'Pick a region' });
    document.addValidation({ range: at(1, 1, 20), rule: { kind: 'number', min: 0, max: 100, integer: true }, strict: true });
    document.addValidation({ range: at(1, 2, 20), rule: { kind: 'number', min: 0.5 } });
    document.addValidation({ range: at(1, 3, 20), rule: { kind: 'text', maxLength: 12 }, strict: true });
    document.addValidation({ range: at(1, 4, 20), rule: { kind: 'date', from: 46000, to: 46400 }, strict: false });
    return document;
  }

  it('come back as the rules they were', async () => {
    const back = await roundTrip(ruled());
    expect(back.validations.map(validation => ({ rule: validation.rule, strict: validation.strict === true, message: validation.message }))).toEqual([
      { rule: { kind: 'list', values: ['North', 'South', 'Say "hi"'] }, strict: true, message: 'Pick a region' },
      { rule: { kind: 'number', min: 0, max: 100, integer: true }, strict: true, message: undefined },
      { rule: { kind: 'number', min: 0.5 }, strict: false, message: undefined },
      { rule: { kind: 'text', maxLength: 12 }, strict: true, message: undefined },
      { rule: { kind: 'date', from: 46000, to: 46400 }, strict: false, message: undefined }
    ]);
    expect(back.validationAt(5, 1)?.rule.kind).toBe('number');
    expect(back.validationAt(21, 1)).toBeNull();
  });

  it('go on refusing what they refused', async () => {
    const back = await roundTrip(ruled());
    expect(back.setCell(3, 1, '101')).not.toBeNull();
    expect(back.setCell(3, 1, '42')).toBeNull();
    expect(back.setCell(3, 0, 'West')).not.toBeNull();
  });

  it('leave out a list Excel cannot hold, and say so', () => {
    const document = new SheetDocument();
    document.addValidation({
      range: { start: { row: 0, column: 0, rowAbsolute: false, columnAbsolute: false }, end: { row: 0, column: 0, rowAbsolute: false, columnAbsolute: false } },
      rule: { kind: 'list', values: ['Smith, J', 'Jones, K'] }
    });
    const out = xlsxOfDocument(document, ROWS);
    expect(out.leftOut).toEqual(['some validations']);
    expect(xlsxParts(out.book).find(part => part.name === 'xl/worksheets/sheet1.xml')!.text).not.toContain('dataValidation');
  });
});

/**
 * Conditional formats out and back — Phase 23.
 */
describe('conditional formats in an exported workbook', () => {
  const column = (at: number) => ({
    start: { row: 1, column: at, rowAbsolute: false, columnAbsolute: false },
    end: { row: 30, column: at, rowAbsolute: false, columnAbsolute: false }
  });

  function ruled(): SheetDocument {
    const document = new SheetDocument();
    const red = { fill: '#fce8e6', color: '#c5221f' };
    document.addConditional({ range: column(0), test: { kind: 'greaterThan', value: 100 }, paint: red });
    document.addConditional({ range: column(1), test: { kind: 'between', low: 1, high: 5 }, paint: { bold: true } });
    document.addConditional({ range: column(2), test: { kind: 'equalTo', value: 'Held' }, paint: { fill: '#e6f4ea' } });
    document.addConditional({ range: column(3), test: { kind: 'textContains', text: 'Below' }, paint: { color: '#c5221f', italic: true } });
    document.addConditional({ range: column(4), test: { kind: 'isEmpty' }, paint: red });
    document.addConditional({ range: column(5), test: { kind: 'formula', input: '=F2>E2*2' }, paint: red });
    document.addConditional({ range: column(6), test: null, scale: { from: '#fde2e2', middle: '#fff4cc', to: '#d9efdc' } });
    document.addConditional({ range: column(7), test: null, scale: { from: '#ffffff', to: '#1967d2' } });
    return document;
  }

  it('come back as the rules they were, in the same order', async () => {
    const back = await roundTrip(ruled());
    expect(back.conditional.map(rule => ({ test: rule.test, paint: rule.paint, scale: rule.scale }))).toEqual(
      ruled().conditional.map(rule => ({ test: rule.test, paint: rule.paint, scale: rule.scale }))
    );
    expect(back.conditional[5].range.start).toMatchObject({ row: 1, column: 5 });
  });

  it('write one differential format for a paint two rules share', () => {
    const styles = xlsxParts(xlsxOfDocument(ruled(), ROWS).book).find(part => part.name === 'xl/styles.xml')!.text;
    // Red three times, bold, green, italic red: four paints, not six.
    expect(styles.match(/<dxf>/g)?.length).toBe(4);
    expect(styles).toContain('<bgColor rgb="FFFCE8E6"/>');
  });
});

/**
 * Charts out and back — Phase 23.
 */
describe('charts in an exported workbook', () => {
  const KINDS = ['column', 'bar', 'stacked', 'line', 'area', 'pie', 'scatter'] as const;

  function charted(): SheetDocument {
    const document = new SheetDocument();
    document.setCell(0, 0, 'Month');
    document.setCell(0, 1, 'North');
    document.setCell(0, 2, 'South');
    ['Jan', 'Feb', 'Mar'].forEach((month, at) => {
      document.setCell(at + 1, 0, month);
      document.setCell(at + 1, 1, String((at + 1) * 10));
      document.setCell(at + 1, 2, String((at + 1) * 7));
    });
    KINDS.forEach((kind, at) => {
      document.addChart({
        kind,
        title: at === 0 ? 'Sales by month' : '',
        range: {
          start: { row: 0, column: 0, rowAbsolute: false, columnAbsolute: false },
          end: { row: 3, column: kind === 'pie' ? 1 : 2, rowAbsolute: false, columnAbsolute: false }
        },
        place: { x: 400, y: at * 320, width: 480, height: 300 },
        legend: at % 2 === 0
      });
    });
    document.sheet.recalculate();
    return document;
  }

  it('come back as the charts they were: kind, title, legend, range and place', async () => {
    const back = await roundTrip(charted());
    expect(back.charts.map(chart => ({ kind: chart.kind, title: chart.title, legend: chart.legend, place: chart.place }))).toEqual(
      charted().charts.map(chart => ({ kind: chart.kind, title: chart.title, legend: chart.legend, place: chart.place }))
    );
    expect(back.charts[0].range).toMatchObject({ start: { row: 0, column: 0 }, end: { row: 3, column: 2 } });
    expect(back.charts[5].range).toMatchObject({ start: { row: 0, column: 0 }, end: { row: 3, column: 1 } });
  });

  it('spell out each series as the chart on screen reads it', () => {
    const parts = xlsxParts(xlsxOfDocument(charted(), ROWS).book);
    const first = parts.find(part => part.name === 'xl/charts/chart1.xml')!.text;
    expect(first).toContain('<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>');
    expect(first).toContain('<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f>');
    expect(first).toContain('<c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>');
    expect(parts.find(part => part.name === 'xl/worksheets/sheet1.xml')!.text).toContain('<drawing r:id="rId3"/>');
    expect(parts.filter(part => part.name.startsWith('xl/charts/')).length).toBe(7);
    expect(parts.find(part => part.name === '[Content_Types].xml')!.text).toContain('/xl/charts/chart7.xml');
  });

  it('say nothing is left out', () => {
    expect(xlsxOfDocument(charted(), ROWS).leftOut).toEqual([]);
  });
});

/** A summary sheet of charts over a data sheet — after Part Four. */
describe('a chart of another sheet in an exported workbook', () => {
  function summary(): SheetDocument {
    const document = new SheetDocument();
    document.renameSheet(0, 'Summary');
    document.addSheet('Q3 data');
    document.activate(1);
    document.setCell(0, 0, 'Month');
    document.setCell(0, 1, 'Units');
    ['Jul', 'Aug', 'Sep'].forEach((month, row) => {
      document.setCell(row + 1, 0, month);
      document.setCell(row + 1, 1, String((row + 1) * 5));
    });
    document.activate(0);
    document.addChart({
      kind: 'line',
      title: 'Q3',
      legend: true,
      range: {
        start: { row: 0, column: 0, rowAbsolute: false, columnAbsolute: false, sheet: 'Q3 data' },
        end: { row: 3, column: 1, rowAbsolute: false, columnAbsolute: false }
      },
      place: { x: 20, y: 20, width: 400, height: 260 }
    });
    return document;
  }

  it('writes the series as references to the other sheet, quoted', () => {
    const chart = xlsxParts(xlsxOfDocument(summary(), ROWS).book).find(part => part.name === 'xl/charts/chart1.xml')!.text;
    expect(chart).toContain("<c:val><c:numRef><c:f>'Q3 data'!$B$2:$B$4</c:f>");
    expect(chart).toContain("<c:tx><c:strRef><c:f>'Q3 data'!$B$1</c:f>");
  });

  it('comes back on the summary sheet, reading the data sheet', async () => {
    const back = await roundTrip(summary());
    back.activate(0);
    expect(back.charts).toHaveLength(1);
    expect(back.charts[0].range).toMatchObject({ start: { row: 0, column: 0, sheet: 'Q3 data' }, end: { row: 3, column: 1 } });
  });
});
