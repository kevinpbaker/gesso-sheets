import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { PLAIN, type CellFormat } from '../sheet/Format';
import { COLUMN_WIDTH, ROW_HEIGHT } from './dimensions';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Phase 25: a block picked up by its border and put down somewhere
 * else — Phase 20's cut and paste as one gesture, or a copy with Ctrl.
 *
 * Through the hit tester by coordinate, because the claim is about
 * where a press lands: on the selection's edge, or inside it.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

const bold: CellFormat = { number: { kind: 'general' }, paint: { ...PLAIN, bold: true } };

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  document.setCell(1, 1, '1');
  document.setFormat(1, 1, bold);
  document.setCell(1, 2, '2');
  document.setCell(2, 1, '3');
  document.setCell(2, 2, '=B2+C2');
  document.setCell(0, 7, '=SUM(B2:C3)');
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 400 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  service.setSelection(1, 1, 2, 2);
  await served.settle();
  await ui.settle();
  return { ui, served, service, document };
}

let h: Harness;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
});

async function settle(): Promise<void> {
  await h.ui.settle();
  await h.served.settle();
  await h.ui.settle();
}

/** Where the selection B2:C3's top edge is, half way along. */
function topEdge(): { x: number; y: number } {
  const b = h.ui.getVisibleBox(h.ui.getByRole('columnheader', { name: 'B' }));
  const two = h.ui.getVisibleBox(h.ui.getByRole('rowheader', { name: '2' }));
  return { x: b.x + b.width, y: two.y - 1 };
}

async function drag(from: { x: number; y: number }, by: { x: number; y: number }, ctrl = false): Promise<void> {
  h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
  h.ui.fireEvent.pointerMove(from.x + 6, from.y + 6, { buttons: 1 });
  h.ui.fireEvent.pointerMove(from.x + by.x, from.y + by.y, { buttons: 1 });
  h.ui.fireEvent.pointerUp(from.x + by.x, from.y + by.y, { modifiers: { ctrl } });
  await settle();
}

const three = { x: 3 * COLUMN_WIDTH, y: 3 * ROW_HEIGHT };

describe('a block dragged by its border', () => {
  it('moves to where it is let go, and what read it reads it there', async () => {
    h = await mount();
    await drag(topEdge(), three);
    const sheet = h.document.sheet;
    // B2:C3 is now E5:F6.
    expect(sheet.input(4, 4)).toBe('1');
    expect(sheet.input(5, 5)).toBe('=E5+F5');
    expect(h.document.formatAt(4, 4)).toEqual(bold);
    expect(sheet.input(1, 1)).toBe('');
    expect(sheet.input(2, 2)).toBe('');
    expect(sheet.input(0, 7)).toBe('=SUM(E5:F6)');
    expect(sheet.value(0, 7)).toBe(9);
    expect(h.document.selection).toEqual({ row: 4, column: 4, anchorRow: 5, anchorColumn: 5 });

    h.service.undo();
    await settle();
    expect(sheet.input(1, 1)).toBe('1');
    expect(sheet.input(4, 4)).toBe('');
    expect(sheet.input(0, 7)).toBe('=SUM(B2:C3)');
  });

  it('copies with Ctrl held at the drop, and leaves the original', async () => {
    h = await mount();
    await drag(topEdge(), three, true);
    const sheet = h.document.sheet;
    expect(sheet.input(1, 1)).toBe('1');
    expect(sheet.input(4, 4)).toBe('1');
    // A copy's formula moves with it, as a paste's does.
    expect(sheet.input(5, 5)).toBe('=E5+F5');
    expect(sheet.input(0, 7)).toBe('=SUM(B2:C3)');

    h.service.undo();
    await settle();
    expect(sheet.input(4, 4)).toBe('');
  });

  it('still sweeps when the press is inside the selection', async () => {
    h = await mount();
    const c = h.ui.getVisibleBox(h.ui.getByRole('columnheader', { name: 'C' }));
    const three = h.ui.getVisibleBox(h.ui.getByRole('rowheader', { name: '3' }));
    await drag({ x: c.x + c.width / 2, y: three.y + three.height / 2 }, { x: 2 * COLUMN_WIDTH, y: 2 * ROW_HEIGHT });
    expect(h.document.sheet.input(1, 1)).toBe('1');
    expect(h.document.selection).toMatchObject({ anchorRow: 2, anchorColumn: 2, row: 2, column: 2, cornerRow: 4, cornerColumn: 4 });
  });

  it('leaves the clipboard alone', async () => {
    h = await mount();
    let marked: unknown = 'unset';
    await drag(topEdge(), three);
    h.service.clipboard.subscribe(clipboard => (marked = clipboard.marked)).unsubscribe();
    expect(marked).toBeNull();
  });

  it('shows the move pointer over the border, and not inside', async () => {
    h = await mount();
    const grid = h.ui.getByRole('grid');
    const edge = topEdge();
    h.ui.fireEvent.pointerMove(edge.x, edge.y);
    await h.ui.settle();
    expect(grid.properties.get('cursor')).toBe('move');
    h.ui.fireEvent.pointerMove(edge.x, edge.y + ROW_HEIGHT / 2);
    await h.ui.settle();
    expect(grid.properties.get('cursor')).toBeUndefined();
  });
});
