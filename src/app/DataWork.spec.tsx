import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import type { UiKeyModifiers } from 'gesso-core';

import { FUNCTION_ENTRIES, functionsMatching } from './FunctionsDialog';
import type { SheetTransfer } from './SheetContract';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Phase 41: the filter on the toolbar and in the header, AutoSum, and
 * the function reference — the service directly, then the screen from
 * the keyboard.
 */

function orders(): SheetDocument {
  const document = new SheetDocument();
  const rows = [
    ['Region', 'Units', 'Price'],
    ['North', '10', '2.5'],
    ['South', '20', '3'],
    ['North', '30', '4']
  ];
  rows.forEach((line, row) => line.forEach((input, column) => document.setCell(row, column, input)));
  document.sheet.recalculate();
  return document;
}

function service(document = orders()): SheetService {
  return new SheetService(document, { rowCount: 100, columnCount: 10 });
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

describe('AutoSum', () => {
  it('totals the run of numbers above one cell', () => {
    const document = orders();
    const sheet = service(document);
    sheet.setSelection(4, 1, 4, 1);
    sheet.autoSum('SUM');
    expect(document.inputAt(4, 1)).toBe('=SUM(B2:B4)');
  });

  it('totals the numbers to the left when there are none above', () => {
    const document = orders();
    document.setCell(1, 4, '=B2*C2');
    document.sheet.recalculate();
    const sheet = service(document);
    sheet.setSelection(1, 5, 1, 5);
    sheet.autoSum('MAX');
    expect(document.inputAt(1, 5)).toBe('=MAX(E2)');
  });

  it('writes a total under every column of a range', () => {
    const document = orders();
    const sheet = service(document);
    sheet.setSelection(1, 1, 3, 2);
    sheet.autoSum('AVERAGE');
    expect([document.inputAt(4, 1), document.inputAt(4, 2)]).toEqual(['=AVERAGE(B2:B4)', '=AVERAGE(C2:C4)']);
  });

  it('writes one to the right of a range one row high', () => {
    const document = orders();
    const sheet = service(document);
    sheet.setSelection(1, 1, 1, 2);
    sheet.autoSum('COUNT');
    expect(document.inputAt(1, 3)).toBe('=COUNT(B2:C2)');
  });

  it('is one step of undo, however many it wrote', () => {
    const document = orders();
    const sheet = service(document);
    sheet.setSelection(1, 1, 3, 2);
    sheet.autoSum('SUM');
    expect(document.undoLabel).toBe('AutoSum');
    sheet.undo();
    expect([document.inputAt(4, 1), document.inputAt(4, 2)]).toEqual(['', '']);
  });

  it('says so when there is nothing to total', () => {
    const document = orders();
    const sheet = service(document);
    sheet.setSelection(8, 6, 8, 6);
    sheet.autoSum('SUM');
    expect(document.inputAt(8, 6)).toBe('');
    expect(latest<SheetTransfer>(sheet.transfer).report).toBe('There are no numbers above G9 or to its left to total.');
  });
});

describe('the function reference', () => {
  it('lists every function the formula editor knows, alphabetically', () => {
    expect(FUNCTION_ENTRIES.length).toBeGreaterThan(150);
    expect(FUNCTION_ENTRIES.find(entry => entry.name === 'SUMIF')).toEqual({
      name: 'SUMIF',
      call: expect.stringMatching(/^SUMIF\(range, criteria/),
      summary: expect.any(String)
    });
    const names = FUNCTION_ENTRIES.map(entry => entry.name);
    expect(names).toEqual([...names].sort((a, b) => a.localeCompare(b)));
  });

  it('finds names that start with the search first, then names and descriptions that hold it', () => {
    const found = functionsMatching('sum').map(entry => entry.name);
    expect(found[0]).toBe('SUM');
    expect(found.indexOf('SUMIF')).toBeLessThan(found.indexOf('DSUM') === -1 ? Infinity : found.indexOf('DSUM'));
    expect(functionsMatching('largest').map(entry => entry.name)).toContain('MAX');
  });
});

describe('on the screen', () => {
  let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument } | undefined;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
    h = undefined;
  });

  async function mount(): Promise<void> {
    const document = orders();
    const sheet = service(document);
    const served = serveForTest([sheetChannel(sheet)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 600 });
    h = { ui, served, service: sheet, document };
    await settle();
  }

  async function settle(): Promise<void> {
    await h!.ui.settle();
    await h!.served.settle();
    await h!.ui.settle();
  }

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
    h!.ui.fireEvent.press(key, modifiers);
    await settle();
  }

  async function reachTheGrid(): Promise<void> {
    const grid = h!.ui.getByRole('grid');
    let stops = 0;
    while (h!.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the grid');
      }
      h!.ui.fireEvent.tab();
      await h!.ui.settle();
    }
  }

  /** Onto the toolbar and along it to the button called `name`, then Enter. */
  async function toolbarButton(name: string): Promise<void> {
    const toolbar = h!.ui.getByRole('toolbar');
    while (h!.ui.runtime.input.focus.focusedNode !== toolbar) {
      h!.ui.fireEvent.tab();
      await h!.ui.settle();
    }
    const buttons = h!.ui.allNodes().filter(node => node.parent === toolbar);
    const index = buttons.findIndex(node => node.properties.get('label') === name);
    await press('Home');
    for (let step = 0; step < index; step++) {
      await press('ArrowRight');
    }
    await press('Enter');
  }

  async function menu(mnemonic: string, item: string): Promise<void> {
    await reachTheGrid();
    await press('F10');
    await press(mnemonic);
    h!.ui.fireEvent.click(h!.ui.getByRole('menuitem', { name: item }));
    await settle();
  }

  it('filters from the funnel on the toolbar, marks the column, and clears it again', async () => {
    await mount();
    h!.service.setSelection(1, 0, 1, 0);
    await settle();
    await toolbarButton('Filter');
    expect([...h!.document.filteredRows]).toEqual([2]);
    expect(h!.ui.getByRole('button', { name: 'Filter' })).toHaveSemantics({ states: ['pressed'] });
    expect(h!.ui.getByRole('columnheader', { name: 'A, filtered' })).toBeDefined();

    await toolbarButton('Filter');
    expect([...h!.document.filteredRows]).toEqual([]);
    expect(h!.ui.queryByRole('columnheader', { name: 'A, filtered' })).toBeNull();
  });

  it('totals with Σ ▾, choosing the function', async () => {
    await mount();
    h!.service.setSelection(4, 2, 4, 2);
    await settle();
    await toolbarButton('AutoSum');
    expect(h!.ui.getByRole('menu', { name: 'AutoSum' })).toBeDefined();
    h!.ui.fireEvent.click(h!.ui.getByRole('menuitem', { name: 'Average' }));
    await settle();
    expect(h!.document.inputAt(4, 2)).toBe('=AVERAGE(C2:C4)');
  });

  it('totals with Alt+= from the sheet', async () => {
    await mount();
    h!.service.setSelection(4, 1, 4, 1);
    await settle();
    await reachTheGrid();
    await press('=', { alt: true });
    expect(h!.document.inputAt(4, 1)).toBe('=SUM(B2:B4)');
  });

  it('lists the functions from Help, and narrows them as the search is typed', async () => {
    await mount();
    await menu('h', 'Functions…');
    expect(h!.ui.getByRole('dialog', { name: 'Functions' })).toBeDefined();
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(h!.ui.getByRole('searchbox', { name: 'Search functions' }));
    h!.ui.fireEvent.type('xlook');
    await settle();
    const items = h!.ui.getAllByRole('listitem').map(node => String(node.properties.get('label')));
    expect(items).toHaveLength(1);
    expect(items[0]).toMatch(/^XLOOKUP\(/);
    expect(h!.ui.queryByRole('button', { name: 'Insert XLOOKUP' })).toBeNull();
  });

  it('puts a function in the cell from Insert ▸ Function', async () => {
    await mount();
    h!.service.setSelection(5, 1, 5, 1);
    await settle();
    await menu('i', 'Function…');
    expect(h!.ui.getByRole('dialog', { name: 'Insert a function' })).toBeDefined();
    h!.ui.fireEvent.type('sumif');
    await settle();
    h!.ui.fireEvent.click(h!.ui.getByRole('button', { name: 'Insert SUMIF' }));
    await settle();
    expect(h!.ui.queryByRole('dialog', { name: 'Insert a function' })).toBeNull();
    expect(h!.ui.getByRole('textbox', { name: 'Formula' }).properties.get('value')).toBe('=SUMIF(');
  });
});
