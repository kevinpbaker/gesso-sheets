import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, textProperty, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import type { UiKeyModifiers } from 'gesso-core';

import { openXlsx } from '../sheet/Xlsx';
import { writeXlsx } from '../sheet/XlsxWrite';
import { fontStack } from './fonts';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { applySnapshot, snapshotOf } from './SheetFile';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';
import { platformInflate, snapshotOfXlsx } from './SheetXlsx';
import { platformDeflate, xlsxOfDocument } from './SheetXlsxOut';

/**
 * Phase 40: the toolbar's lists — number format, font, size, borders,
 * merge — reached from the keyboard, one tab stop and the arrows, and
 * the font kept in the file and drawn.
 */

let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument } | undefined;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
  h = undefined;
});

async function mount(): Promise<void> {
  const document = new SheetDocument();
  document.setCell(1, 1, '0.25');
  document.setCell(1, 2, 'Hello');
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 600 });
  h = { ui, served, service, document };
  await settle();
  h.service.setSelection(1, 1, 1, 1);
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

/** Onto the toolbar with Tab, then along it to the button called `name`, then Enter. */
async function openList(name: string): Promise<void> {
  const toolbar = h!.ui.getByRole('toolbar');
  let stops = 0;
  while (h!.ui.runtime.input.focus.focusedNode !== toolbar) {
    if (stops++ > 6) {
      throw new Error('Tab never reached the toolbar');
    }
    h!.ui.fireEvent.tab();
    await h!.ui.settle();
  }
  const buttons = h!.ui.allNodes().filter(node => node.parent?.properties.get('role') === 'toolbar' || isButtonIn(node, toolbar));
  const index = buttons.findIndex(node => node.properties.get('label') === name);
  if (index < 0) {
    throw new Error(`no toolbar button called ${name}`);
  }
  await press('Home');
  for (let step = 0; step < index; step++) {
    await press('ArrowRight');
  }
  await press('Enter');
}

function isButtonIn(node: { parent: unknown; properties: Map<string, unknown> }, toolbar: unknown): boolean {
  return node.parent === toolbar;
}

/** A menu item by the end of its label: the label carries a tick or the space for one. */
async function choose(ending: string): Promise<void> {
  const item = h!.ui.getAllByRole('menuitem').find(node => String(node.properties.get('label') ?? '').trimEnd().endsWith(ending));
  if (item === undefined) {
    throw new Error(`no menu item ending ${ending}`);
  }
  h!.ui.fireEvent.click(item);
  await settle();
}

const buttonText = (name: string): string | undefined => {
  const button = h!.ui.getByRole('button', { name });
  return h!.ui
    .allNodes()
    .filter(node => node.parent === button)
    .map(node => textProperty(node))
    .find(text => text !== undefined);
};

describe('the toolbar’s lists', () => {
  it('sets a number format from 123 ▾, and ticks the one the cell has', async () => {
    await mount();
    await openList('Number format');
    expect(h!.ui.getByRole('menu', { name: 'Number format' })).toBeDefined();
    expect(h!.ui.getAllByRole('menuitem').map(node => node.properties.get('label'))).toContain('✓ General');
    await choose('Percent   10%');
    expect(h!.document.formatAt(1, 1).number).toEqual({ kind: 'percent', places: 0 });
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(h!.ui.getByRole('grid'));
  });

  it('sets a font, and the button says which', async () => {
    await mount();
    expect(buttonText('Font')).toBe('Default ▾');
    await openList('Font');
    await choose('Georgia');
    expect(h!.document.formatAt(1, 1).paint.fontFamily).toBe('Georgia');
    expect(buttonText('Font')).toBe('Georgia ▾');
  });

  it('takes a font back to the sheet’s own', async () => {
    await mount();
    await openList('Font');
    await choose('Georgia');
    await openList('Font');
    await choose('Default');
    expect(h!.document.formatAt(1, 1).paint.fontFamily).toBeUndefined();
  });

  it('sets a size', async () => {
    await mount();
    expect(buttonText('Font size')).toBe('12 ▾');
    await openList('Font size');
    await choose('18');
    expect(h!.document.formatAt(1, 1).paint.fontSize).toBe(18);
    expect(buttonText('Font size')).toBe('18 ▾');
  });

  it('draws borders from the borders list', async () => {
    await mount();
    await openList('Borders');
    await choose('All borders');
    expect(h!.document.formatAt(1, 1).paint.borders.top.width).toBe(1);
  });

  it('merges from the merge list', async () => {
    await mount();
    h!.service.setSelection(1, 1, 1, 2);
    await settle();
    await openList('Merge cells');
    await choose('Merge cells');
    expect(h!.document.merges.all).toEqual([{ firstRow: 1, lastRow: 1, firstColumn: 1, lastColumn: 2 }]);
  });

  it('opens the font list from Format ▸ Font…', async () => {
    await mount();
    h!.ui.fireEvent.focus(h!.ui.getByRole('grid'));
    await press('F10');
    await press('o');
    h!.ui.fireEvent.click(h!.ui.getByRole('menuitem', { name: 'Font…' }));
    await settle();
    expect(h!.ui.getByRole('menu', { name: 'Font' })).toBeDefined();
  });
});

describe('a font', () => {
  it('is drawn, as a stack the machine can fall back through', async () => {
    await mount();
    h!.service.format({ fontFamily: 'Georgia' });
    await settle();
    const cell = h!.ui.allNodes().find(node => node.properties.get('text') === '0.25');
    expect(cell?.properties.get('fontFamily')).toBe(fontStack('Georgia'));
    expect(fontStack('Georgia')).toBe('Georgia, "DejaVu Serif", serif');
    expect(fontStack('Optima')).toBe('"Optima", system-ui, sans-serif');
  });

  it('is kept in the file, and an old file without one reads as the sheet’s own', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, 'x');
    document.setFormat(0, 0, { ...document.formatAt(0, 0), paint: { ...document.formatAt(0, 0).paint, fontFamily: 'Courier New' } });
    const back = new SheetDocument();
    applySnapshot(back, JSON.parse(JSON.stringify(snapshotOf(document))));
    expect(back.formatAt(0, 0).paint.fontFamily).toBe('Courier New');
    expect(back.formatAt(1, 1).paint.fontFamily).toBeUndefined();
  });

  it('goes out to an .xlsx by name and comes back', async () => {
    const document = new SheetDocument();
    document.setCell(0, 0, 'x');
    document.setFormat(0, 0, { ...document.formatAt(0, 0), paint: { ...document.formatAt(0, 0).paint, fontFamily: 'Georgia' } });
    const { book } = xlsxOfDocument(document, 50);
    const read = await openXlsx(await writeXlsx(book, platformDeflate), platformInflate, { rows: 50, columns: 10 });
    const back = new SheetDocument();
    applySnapshot(back, snapshotOfXlsx(read, 10));
    expect(back.formatAt(0, 0).paint.fontFamily).toBe('Georgia');
  });
});
