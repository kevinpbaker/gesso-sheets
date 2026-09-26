import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiDroppedFile } from 'gesso-core';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * A file dropped on the window, end to end.
 *
 * The shell half is the engine's and is specced there; this starts
 * where it hands over — a `fileDrop` message reaching the runtime —
 * and ends at the sheet the application worker made of it, with the
 * status line in between saying so.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
}

let h: Harness | undefined;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
  h = undefined;
});

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 600 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served, document };
}

function csv(name: string, text: string): UiDroppedFile {
  const bytes = new TextEncoder().encode(text).buffer;
  return { name, mediaType: 'text/csv', size: bytes.byteLength, lastModified: 0, bytes };
}

async function drop(files: readonly UiDroppedFile[]): Promise<void> {
  const runtime = h!.ui.runtime;
  // What a browser shows during the drag: types, no names, no bytes.
  const during = files.map(file => ({ name: '', mediaType: file.mediaType, size: 0, lastModified: 0 }));
  runtime.applyFileDrop({ type: 'fileDrop', phase: 'enter', x: 600, y: 300, files: during });
  runtime.applyFileDrop({ type: 'fileDrop', phase: 'over', x: 610, y: 310, files: during });
  runtime.applyFileDrop({ type: 'fileDrop', phase: 'drop', x: 610, y: 310, files });
  await h!.ui.settle();
  await h!.served.settle();
  await h!.ui.settle();
}

describe('a file dropped on the window', () => {
  it('opens a CSV on a sheet of its own, and says so', async () => {
    h = await mount();
    await drop([csv('Q3 sales.csv', '﻿Region,Units\r\nNorth,120\r\n')]);

    expect(h.document.sheets().map(sheet => sheet.name)).toEqual(['Sheet1', 'Q3 sales']);
    expect(h.document.sheet.input(0, 0)).toBe('Region');
    expect(h.document.sheet.value(1, 1)).toBe(120);
    expect(h.ui.getByRole('tab', { name: 'Q3 sales' })).toHaveSemantics({ states: ['selected'] });
    expect(h.ui.getByText('Opened 2 rows from Q3 sales.csv.')).toBeDefined();
  });

  it('opens each of several files', async () => {
    h = await mount();
    await drop([csv('a.csv', 'x\n'), csv('b.csv', 'y\n')]);
    expect(h.document.sheets().map(sheet => sheet.name)).toEqual(['Sheet1', 'a', 'b']);
  });

  it('turns down a file it cannot read, and leaves the workbook alone', async () => {
    h = await mount();
    await drop([{ name: 'photo.png', mediaType: 'image/png', size: 4, lastModified: 0, bytes: new ArrayBuffer(4) }]);
    expect(h.document.sheetCount).toBe(1);
    expect(h.ui.getByText('photo.png was not opened: only CSV files can be opened so far.')).toBeDefined();
  });
});
