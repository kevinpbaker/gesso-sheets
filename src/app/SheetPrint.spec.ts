import { describe, expect, it } from 'vitest';

import { relativeRef } from '../sheet/A1';
import { PLAIN } from '../sheet/Format';
import { SheetDocument } from './SheetDocument';
import { cellStyle, printTable, printedExtent, type PrintJob } from './SheetPrint';
import { SheetService } from './SheetService';

/**
 * Print, as the table a browser prints — the exit for printing in
 * Phase 39. What is asserted is the markup: the cells there are, in
 * the order and spans they are drawn, and each one's paint as a style.
 */

function filled(): SheetDocument {
  const document = new SheetDocument();
  document.setCell(0, 0, 'Region');
  document.setCell(0, 1, 'Revenue');
  document.setCell(1, 0, 'North <b>');
  document.setCell(1, 1, '1200');
  document.setCell(2, 0, 'South');
  document.setCell(2, 1, '900');
  document.sheet.recalculate();
  return document;
}

describe('the printed table', () => {
  it('runs from A1 to the last cell anything is in', () => {
    const document = filled();
    expect(printedExtent(document)).toEqual({ rows: 3, columns: 2 });
    const html = printTable(document);
    expect(html.match(/<tr/g)).toHaveLength(3);
    expect(html.match(/<td/g)).toHaveLength(6);
  });

  it('escapes what a cell holds', () => {
    expect(printTable(filled())).toContain('<td>North &lt;b&gt;</td>');
  });

  it('sets numbers to the right, as the grid does', () => {
    expect(printTable(filled())).toContain('<td style="text-align:right">1200</td>');
  });

  it('leaves out hidden and filtered rows', () => {
    const document = filled();
    document.hiddenRows.add(1);
    const html = printTable(document);
    expect(html).not.toContain('North');
    expect(html.match(/<tr/g)).toHaveLength(2);
  });

  it('spans a merge, and draws nothing under it', () => {
    const document = filled();
    document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1 });
    const html = printTable(document);
    expect(html).toContain('<td colspan="2">Region</td>');
    expect(html).not.toContain('Revenue');
  });

  it('keeps the column widths', () => {
    const document = filled();
    document.columnWidths = [80, 140];
    expect(printTable(document)).toContain('<colgroup><col style="width:80px"><col style="width:140px"></colgroup>');
  });

  it('says so for an empty sheet', () => {
    expect(printTable(new SheetDocument())).toBe('<p class="empty">This sheet is empty.</p>');
  });
});

describe('a cell’s paint as a style', () => {
  it('writes what is set and nothing else', () => {
    expect(cellStyle(PLAIN, 'left')).toBe('');
    expect(
      cellStyle(
        {
          ...PLAIN,
          bold: true,
          italic: true,
          underline: true,
          fontSize: 16,
          color: '#b91c1c',
          fill: '#fef3c7',
          align: 'center',
          wrap: true,
          borders: { ...PLAIN.borders, bottom: { width: 2, color: '' } }
        },
        'right'
      )
    ).toBe(
      'font-weight:bold;font-style:italic;text-decoration:underline;font-size:16pt;color:#b91c1c;background:#fef3c7;text-align:center;white-space:normal;border-bottom:2px solid #000'
    );
  });

  it('drops a colour that is not a plain colour', () => {
    expect(cellStyle({ ...PLAIN, fill: 'red;background:url(x)' }, 'left')).toBe('');
  });
});

describe('printing from the service', () => {
  it('hands the print window the sheet in view, titled, with a rule’s paint', async () => {
    const document = filled();
    document.addConditional({
      range: { start: relativeRef(1, 1), end: relativeRef(2, 1) },
      test: { kind: 'greaterThan', value: 1000 },
      paint: { fill: '#dcfce7' }
    });
    const jobs: PrintJob[] = [];
    const service = new SheetService(document, { rowCount: 50, columnCount: 10, printer: job => jobs.push(job) });
    await service.restore();
    service.print(true);
    expect(jobs).toHaveLength(1);
    expect(jobs[0]).toMatchObject({ id: 1, title: 'Untitled', pdf: true });
    expect(jobs[0].html).toContain('<td style="background:#dcfce7;text-align:right">1200</td>');
    expect(jobs[0].html).toContain('<td style="text-align:right">900</td>');
  });

  it('says it cannot, with nowhere to print to', () => {
    const service = new SheetService(filled(), { rowCount: 50, columnCount: 10 });
    let report = '';
    service.transfer.subscribe(transfer => (report = transfer.report)).unsubscribe();
    service.print(false);
    service.transfer.subscribe(transfer => (report = transfer.report)).unsubscribe();
    expect(report).toBe('This browser has no way to print from here.');
  });
});
