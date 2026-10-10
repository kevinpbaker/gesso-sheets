import { afterEach, describe, expect, it } from 'vitest';

import { createComponent, type ShellFileRequest, type ShellFileResult, type ShellRequest } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { InMemorySheetLibrary } from './SheetLibrary';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';
import type { PrintJob } from './SheetPrint';

/**
 * The File menu, end to end, with the shell played by the spec.
 *
 * The pickers are the engine's and are specced there. What is asserted
 * here is the conversation around them: what this application asks the
 * shell for, and what it does with each answer — which is where a Save
 * that saved nothing, or an Open that opened the wrong thing, would be.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  /** Every file request the shell has been sent, answered by `answer`. */
  asked: { id: number; request: ShellFileRequest }[];
  /** Every popup asked for, and the print jobs the service handed out. */
  popups: { id: number; url: string; name: string }[];
  printed: PrintJob[];
}

let h: Harness | undefined;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
  h = undefined;
});

async function settle(): Promise<void> {
  for (let round = 0; round < 3; round++) {
    await h!.ui.settle();
    await h!.served.settle();
    await h!.service.settled;
  }
}

async function mount(): Promise<Harness> {
  const library = new InMemorySheetLibrary();
  const printed: PrintJob[] = [];
  const service = new SheetService(new SheetDocument(), {
    library,
    rowCount: 200,
    columnCount: 20,
    printer: job => printed.push(job)
  });
  service.openDocument('');
  await service.settled;
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 600 });
  const asked: Harness['asked'] = [];
  const popups: Harness['popups'] = [];
  ui.runtime.onShellRequest((request: ShellRequest) => {
    if (request.type === 'file') {
      asked.push({ id: request.id, request: request.request });
    }
    if (request.type === 'popup') {
      popups.push({ id: request.id, url: request.url, name: request.name });
    }
  });
  h = { ui, served, service, asked, popups, printed };
  await settle();
  return h;
}

const EMPTY = { files: [], saved: null, recent: [], error: null } as const;

async function answer(result: ShellFileResult): Promise<void> {
  const last = h!.asked[h!.asked.length - 1];
  h!.ui.runtime.settleFile(last.id, result);
  await settle();
}

async function menu(mnemonic: string, item: string): Promise<void> {
  const grid = h!.ui.getByRole('grid');
  let stops = 0;
  while (h!.ui.runtime.input.focus.focusedNode !== grid) {
    if (stops++ > 12) {
      throw new Error('Tab never reached the grid');
    }
    h!.ui.fireEvent.tab();
    await h!.ui.settle();
  }
  h!.ui.fireEvent.press('F10');
  await h!.ui.settle();
  h!.ui.fireEvent.press(mnemonic);
  await h!.ui.settle();
  h!.ui.fireEvent.click(h!.ui.getByRole('menuitem', { name: item }));
  await settle();
}

describe('the File menu', () => {
  it('is first, and offers what a file can have done to it', async () => {
    await mount();
    expect(h!.ui.queryByRole('menuitem')).toBeNull();
    h!.ui.fireEvent.click(h!.ui.getByText('File'));
    await h!.ui.settle();
    for (const name of ['New', 'Open…', 'Open recent…', 'Save', 'Save as…', 'Download sheet as CSV']) {
      expect(h!.ui.getByRole('menuitem', { name })).toBeDefined();
    }
  });

  it('saves the workbook where the shell says, and remembers the file', async () => {
    await mount();
    h!.service.setCell(0, 0, '=6*7');
    await menu('f', 'Save');

    const { request } = h!.asked[0];
    expect(request).toMatchObject({ op: 'save', name: 'Untitled.gsheet', mediaType: 'application/json' });
    expect(request.op === 'save' && request.handle).toBeFalsy();
    expect(request.op === 'save' && JSON.parse(request.text).sheets[0].cells).toContainEqual({ row: 0, column: 0, input: '=6*7' });

    await answer({ ...EMPTY, outcome: 'ok', saved: { name: 'Answer.gsheet', handle: 7, via: 'file' } });
    expect(h!.ui.getByText('Saved in this browser and to Answer.gsheet')).toBeDefined();
    expect(h!.ui.getByText('Saved Answer.gsheet.')).toBeDefined();

    // And the next Save goes back to that file, with no picker.
    h!.service.setCell(0, 1, 'more');
    await settle();
    expect(h!.ui.getByText('Saved in this browser · Answer.gsheet is behind')).toBeDefined();
    await menu('f', 'Save');
    expect(h!.asked[1].request).toMatchObject({ op: 'save', name: 'Answer.gsheet', handle: 7 });
  });

  it('says nothing when the picker is closed, and says why when the browser refuses', async () => {
    await mount();
    await menu('f', 'Save');
    await answer({ ...EMPTY, outcome: 'cancelled' });
    expect(h!.ui.queryByText(/not saved/)).toBeNull();

    await menu('f', 'Save as…');
    await answer({ ...EMPTY, outcome: 'denied', error: 'Must be handling a user gesture.' });
    expect(h!.ui.getByText('Untitled.gsheet was not saved: the browser did not allow it.')).toBeDefined();
  });

  it('downloads the sheet in view as a CSV', async () => {
    await mount();
    h!.service.setCell(0, 0, 'Region');
    await menu('f', 'Download sheet as CSV');
    expect(h!.asked[0].request).toMatchObject({ op: 'save', name: 'Sheet1.csv', mediaType: 'text/csv', text: 'Region\r\n' });
    await answer({ ...EMPTY, outcome: 'ok', saved: { name: 'Sheet1.csv', handle: null, via: 'download' } });
    expect(h!.ui.getByText('Downloaded Sheet1.csv.')).toBeDefined();
  });

  it('opens a workbook as the document here', async () => {
    await mount();
    await menu('f', 'Open…');
    expect(h!.asked[0].request).toMatchObject({ op: 'open', multiple: true });

    const saved = new SheetService(new SheetDocument());
    saved.setCell(2, 1, 'from disk');
    const bytes = new TextEncoder().encode(JSON.stringify(saved.snapshot())).buffer;
    await answer({
      ...EMPTY,
      outcome: 'ok',
      files: [{ name: 'Budget.gsheet', mediaType: 'application/json', lastModified: 0, bytes, handle: 3 }]
    });

    expect(h!.ui.getByText('Opened Budget.gsheet.')).toBeDefined();
    expect(h!.ui.getByText('Saved in this browser and to Budget.gsheet')).toBeDefined();
    expect(h!.service.snapshot().sheets[0].cells).toContainEqual({ row: 2, column: 1, input: 'from disk' });
  });

  it('lists the files the shell remembers, and reopens one', async () => {
    await mount();
    await menu('f', 'Open recent…');
    expect(h!.asked[0].request).toEqual({ op: 'recent' });
    await answer({ ...EMPTY, outcome: 'ok', recent: [{ handle: 3, name: 'Budget.gsheet', used: 1 }] });

    h!.ui.fireEvent.click(h!.ui.getByRole('button', { name: 'Budget.gsheet' }));
    await settle();
    expect(h!.asked[1].request).toEqual({ op: 'reopen', handle: 3 });
    expect(h!.ui.queryByRole('form', { name: 'Recent files' })).toBeNull();
  });

  it('imports a CSV or an Excel workbook from a picker', async () => {
    await mount();
    await menu('f', 'Import CSV or Excel…');
    const { request } = h!.asked[0];
    expect(request.op === 'open' && request.accept?.flatMap(type => type.extensions)).toEqual(['.csv', '.tsv', '.txt', '.xlsx']);
    await answer({
      ...EMPTY,
      outcome: 'ok',
      files: [{ name: 'regions.csv', mediaType: 'text/csv', lastModified: 0, bytes: new TextEncoder().encode('North,10\n').buffer, handle: null }]
    });
    expect(h!.service.snapshot().sheets.map(sheet => sheet.name)).toEqual(['Sheet1', 'regions']);
  });

  it('prints into a window of its own, opened while the click is fresh', async () => {
    await mount();
    h!.service.setCell(0, 0, 'Region');
    await menu('f', 'Print…');
    expect(h!.popups).toEqual([{ id: expect.any(Number), url: '/print.html', name: 'gessosheet-print' }]);
    expect(h!.printed).toHaveLength(1);
    expect(h!.printed[0]).toMatchObject({ title: 'Untitled', pdf: false });
    expect(h!.printed[0].html).toContain('<td>Region</td>');
  });

  it('says what to do when the print window is blocked', async () => {
    await mount();
    await menu('f', 'Export as PDF…');
    expect(h!.printed[0].pdf).toBe(true);
    h!.ui.runtime.settlePopup(h!.popups[0].id, false);
    await settle();
    expect(h!.ui.getByText('The print window was blocked. Allow pop-ups for this site, then print again.')).toBeDefined();
  });
});
