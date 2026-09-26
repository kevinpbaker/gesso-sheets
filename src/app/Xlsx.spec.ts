import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { base64OfBytes } from './base64';
import type { SheetDocumentView, SheetTransfer } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { InMemorySheetLibrary } from './SheetLibrary';
import { SheetService, type Schedule } from './SheetService';

/**
 * An Excel workbook opened as a document, through the service.
 *
 * The reader is specced in `src/sheet/Xlsx.spec.ts`; this is what a
 * workbook becomes on the way in — a document of its own, calculating
 * — read from the same LibreOffice-written file.
 */

const HERE = dirname(fileURLToPath(import.meta.url));
const ORDERS = base64OfBytes(new Uint8Array(readFileSync(join(HERE, '..', 'sheet', 'fixtures', 'orders.xlsx'))));

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

async function opened(base64 = ORDERS, name = 'orders.xlsx') {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const service = new SheetService(new SheetDocument(), {
    schedule,
    library: new InMemorySheetLibrary(),
    rowCount: 200,
    columnCount: 20
  });
  service.openDocument('');
  await service.settled;
  service.importXlsx(name, base64);
  await service.settled;
  while (queue.length > 0) {
    queue.shift()!();
  }
  /** What the sheet shows at a cell, read through the window it publishes. */
  const shown = (sheet: number, row: number, column: number): string | undefined => {
    service.activateSheet(sheet);
    service.setViewport(sheet, 0, 30, 0, 10);
    while (queue.length > 0) {
      queue.shift()!();
    }
    return latest(service.window).cells[row]?.[column];
  };
  return { service, shown };
}

describe('opening an Excel workbook', () => {
  it('opens it as a document of its own, with no file to save back to', async () => {
    const { service } = await opened();
    const view = latest<SheetDocumentView>(service.documentView);
    expect(view).toMatchObject({ name: 'orders', file: null, edited: false });
    expect(latest(service.sheets).entries.map(sheet => sheet.name)).toEqual(['Orders', 'Rates']);
  });

  it('says what came across, and what could not', async () => {
    const { service } = await opened();
    expect(latest<SheetTransfer>(service.transfer).report).toBe(
      'Opened orders.xlsx: 2 sheets.'
    );
  });

  it('calculates its formulas, across sheets and through names', async () => {
    const { shown } = await opened();
    expect(shown(0, 3, 3)).toBe('$13.50');
    expect(shown(0, 7, 3)).toBe('$141.23');
    expect(shown(0, 8, 3)).toBe('$11.30');
    expect(shown(0, 10, 3)).toBe('$11.30');
    expect(shown(0, 9, 3)).toBe('9.6%');
  });

  it('shows its dates as dates, and its codes with their zeros', async () => {
    const { shown } = await opened();
    expect(shown(0, 3, 4)).toBe('2026-09-24');
    expect(shown(0, 3, 5)).toBe('007');
  });

  it('keeps its merge, its hidden row and its bold', async () => {
    const { service } = await opened();
    const orders = service.snapshot().sheets[0];
    expect(orders.merges).toEqual([{ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 }]);
    expect(orders.hiddenRows).toEqual([11]);
    const title = orders.formats.find(cell => cell.row === 0 && cell.column === 0)!;
    expect(orders.palette[title.id].paint).toMatchObject({ bold: true, fill: '#1f3864', color: '#ffffff' });
  });

  it('saves as a .gsheet, asking where', async () => {
    const { service } = await opened();
    service.saveDocument(false);
    expect(latest<SheetTransfer>(service.transfer).download).toMatchObject({
      kind: 'workbook',
      name: 'orders.gsheet',
      handle: null
    });
  });

  it('turns down a file that is not a workbook, and stays where it was', async () => {
    const { service } = await opened(base64OfBytes(new TextEncoder().encode('Region,Units\n')), 'fake.xlsx');
    expect(latest<SheetTransfer>(service.transfer).report).toBe(
      'fake.xlsx was not opened: it is not an Excel workbook; an .xlsx is a zip file, and this is not one.'
    );
    expect(latest(service.sheets).entries.map(sheet => sheet.name)).toEqual(['Sheet1']);
  });
});
