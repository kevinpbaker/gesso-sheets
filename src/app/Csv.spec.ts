import { describe, expect, it } from 'vitest';

import type { SheetTransfer } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { InMemorySheetRepository } from './SheetRepository';
import { SheetService, type Schedule } from './SheetService';

/**
 * A CSV in and a CSV out, through the service.
 *
 * The format itself is specced in `src/sheet/Csv.spec.ts`; this is
 * what a file turns into when it reaches a sheet, and what a sheet
 * turns into on its way out — which is where the decisions are.
 */

function harness(repository = new InMemorySheetRepository()) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, repository, rowCount: 100, columnCount: 10 });
  const drain = () => {
    while (queue.length > 0) {
      queue.shift()!();
    }
  };
  return { document, service, repository, drain };
}

function transfer(service: SheetService): SheetTransfer {
  let seen!: SheetTransfer;
  service.transfer.subscribe(value => (seen = value)).unsubscribe();
  return seen;
}

describe('opening a CSV', () => {
  it('lands on a new sheet named after the file, and shows it', () => {
    const { document, service } = harness();
    service.setCell(0, 0, 'mine');
    service.importCsv('Q3 sales.csv', 'Region,Units\nNorth,120\n');

    expect(document.sheets().map(sheet => sheet.name)).toEqual(['Sheet1', 'Q3 sales']);
    expect(document.active).toBe(1);
    expect(document.sheet.input(1, 1)).toBe('120');
    // The sheet that was in view is untouched.
    expect(document.pageAt(0)!.sheet.input(0, 0)).toBe('mine');
  });

  it('reads numbers and dates as it would if they were typed', () => {
    const { document, service, drain } = harness();
    service.importCsv('a.csv', 'n,d\n1.5,2026-09-22\n');
    drain();
    expect(document.sheet.value(1, 0)).toBe(1.5);
    expect(typeof document.sheet.value(1, 1)).toBe('number');
  });

  /** A CSV is data, and a spreadsheet that ran one would run anybody's. */
  it('keeps a formula as the text it is, through a recalculation', () => {
    const { document, service, drain } = harness();
    service.importCsv('a.csv', 'x,=1+2,=HYPERLINK("http://example.com")\n');
    drain();
    expect(document.sheet.value(0, 1)).toBe('=1+2');
    expect(document.formatAt(0, 1).number.kind).toBe('text');
    expect(document.sheet.value(0, 2)).toBe('=HYPERLINK("http://example.com")');
  });

  it('and still as text after a reload', async () => {
    const repository = new InMemorySheetRepository();
    const first = harness(repository);
    await first.service.restore();
    first.service.importCsv('a.csv', '=1+2\n');
    await first.service.flush();

    const second = harness(repository);
    await second.service.restore();
    second.drain();
    second.service.activateSheet(1);
    expect(second.document.sheet.value(0, 0)).toBe('=1+2');
  });

  it('is not undone a cell at a time', () => {
    const { document, service } = harness();
    service.importCsv('a.csv', 'a,b\n');
    expect(document.canUndo).toBe(false);
  });

  it('says what it opened', () => {
    const { service } = harness();
    service.importCsv('a.csv', 'a\nb\nc\n');
    expect(transfer(service).report).toBe('Opened 3 rows from a.csv.');
  });

  it('cuts a file at the sheet edges, and says by how much', () => {
    const { document, service } = harness();
    const wide = Array.from({ length: 12 }, (_, at) => `c${at}`).join(',');
    const text = Array.from({ length: 101 }, () => wide).join('\n');
    service.importCsv('big.csv', text);

    expect(document.sheet.input(99, 9)).toBe('c9');
    expect(transfer(service).report).toBe(
      'Opened 100 rows from big.csv; 1 row and 2 columns did not fit on the sheet and were left out.'
    );
  });

  it('takes a second file of the same name as a second sheet', () => {
    const { document, service } = harness();
    service.importCsv('a.csv', 'x\n');
    service.importCsv('a.csv', 'y\n');
    expect(document.sheets().map(sheet => sheet.name)).toEqual(['Sheet1', 'a', 'a 2']);
  });
});

describe('saving a sheet as CSV', () => {
  it('writes what the sheet shows, from A1 to the last cell used', () => {
    const { service, drain } = harness();
    service.setCell(0, 0, 'Region');
    service.setCell(0, 1, 'Price');
    service.setCell(1, 0, 'North, upper');
    service.setCell(1, 1, '4');
    service.setCell(1, 2, '=B2*2');
    service.setCell(3, 0, 'last');
    drain();
    service.exportCsv();

    const download = transfer(service).download!;
    expect(download.name).toBe('Sheet1.csv');
    expect(download.mediaType).toBe('text/csv');
    expect(download.text).toBe('Region,Price,\r\n"North, upper",4,8\r\n,,\r\nlast,,\r\n');
  });

  it('writes a formatted number as it is shown', () => {
    const { service, drain } = harness();
    service.setCell(0, 0, '1234.5');
    service.format({ number: { kind: 'currency', places: 2, symbol: '$' } });
    drain();
    service.exportCsv();
    expect(transfer(service).download!.text).toBe('"$1,234.50"\r\n');
  });

  /** The other half of not running a CSV: not handing one to a program that would. */
  it('defuses text another program would run, and leaves negative numbers alone', () => {
    const { service, drain } = harness();
    service.importCsv('a.csv', '=HYPERLINK("x"),@SUM(A1),-4\n');
    drain();
    service.exportCsv();
    expect(transfer(service).download!.text).toBe(`"'=HYPERLINK(""x"")",'@SUM(A1),-4\r\n`);
  });

  it('is a new download each time it is asked for', () => {
    const { service } = harness();
    service.exportCsv();
    const first = transfer(service).download!.serial;
    service.exportCsv();
    expect(transfer(service).download!.serial).toBe(first + 1);
  });
});
