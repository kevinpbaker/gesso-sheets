import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Phase 26: a finger.
 *
 * Every press here is made with `pointer: 'touch'`, through the hit
 * tester, which is the path a tablet's contacts take — the engine's own
 * scroller, long press and slop answer to the device the press names.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  for (let row = 0; row < 200; row++) {
    document.setCell(row, 0, `r${row + 1}`);
  }
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 500, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 400 });
  await ui.settle();
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

/** The middle of a cell, from its column letter and row number. */
function centre(column: string, row: string): { x: number; y: number } {
  const letter = h.ui.getVisibleBox(h.ui.getByRole('columnheader', { name: column }));
  const number = h.ui.getVisibleBox(h.ui.getByRole('rowheader', { name: row }));
  return { x: letter.x + letter.width / 2, y: number.y + number.height / 2 };
}

async function tap(at: { x: number; y: number }): Promise<void> {
  h.ui.fireEvent.pointerDown(at.x, at.y, { pointer: 'touch' });
  h.ui.fireEvent.pointerUp(at.x, at.y, { pointer: 'touch' });
  await settle();
}

async function drag(from: { x: number; y: number }, to: { x: number; y: number }): Promise<void> {
  h.ui.fireEvent.pointerDown(from.x, from.y, { pointer: 'touch' });
  // Past the finger's slop, then the rest of the way.
  h.ui.fireEvent.pointerMove(from.x + (to.x - from.x) / 4, from.y + (to.y - from.y) / 4, { buttons: 1, pointer: 'touch' });
  h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1, pointer: 'touch' });
  h.ui.fireEvent.pointerUp(to.x, to.y, { pointer: 'touch' });
  await settle();
}

const firstRowShown = () => h.ui.getAllByRole('rowheader')[0].properties.get('label');
const items = (): string[] => {
  try {
    return h.ui.getAllByRole('menuitem').map(item => String(item.properties.get('label')));
  } catch {
    return [];
  }
};

describe('a finger on the sheet', () => {
  it('selects a cell with a tap', async () => {
    h = await mount();
    await tap(centre('C', '4'));
    expect(h.document.selection).toEqual({ row: 3, column: 2, anchorRow: 3, anchorColumn: 2 });
  });

  it('scrolls with a drag, rather than sweeping a selection', async () => {
    h = await mount();
    await tap(centre('B', '2'));
    const before = firstRowShown();
    await drag(centre('C', '8'), centre('C', '2'));
    expect(firstRowShown()).not.toBe(before);
    expect(h.document.selection).toEqual({ row: 1, column: 1, anchorRow: 1, anchorColumn: 1 });
  });

  it('opens the menu with a long press, over the cell under the finger', async () => {
    h = await mount();
    const at = centre('D', '5');
    h.ui.fireEvent.pointerDown(at.x, at.y, { pointer: 'touch' });
    await new Promise(resolve => setTimeout(resolve, 650));
    await settle();
    expect(items()).toContain('Copy');
    expect(h.document.selection).toMatchObject({ row: 4, column: 3 });
    h.ui.fireEvent.pointerUp(at.x, at.y, { pointer: 'touch' });
    await settle();
  });

  it('shows two handles on the selection, and extends it by dragging one', async () => {
    h = await mount();
    await tap(centre('B', '2'));
    const end = h.ui.getVisibleBox(h.ui.getByRole('button', { name: 'Selection end' }));
    expect(h.ui.getByRole('button', { name: 'Selection start' })).toBeDefined();
    expect(h.ui.queryByRole('button', { name: 'Fill' })).toBeNull();
    await drag({ x: end.x + end.width / 2, y: end.y + end.height / 2 }, centre('D', '4'));
    // Shift's rule for a hand with no Shift: the corner moved, the
    // active cell stayed.
    expect(h.document.selection).toEqual({ row: 1, column: 1, anchorRow: 1, anchorColumn: 1, cornerRow: 3, cornerColumn: 3 });
  });

  it('pulls the first corner out with the other handle', async () => {
    h = await mount();
    h.service.setSelection(3, 3, 3, 3);
    await settle();
    await tap(centre('D', '4'));
    const start = h.ui.getVisibleBox(h.ui.getByRole('button', { name: 'Selection start' }));
    await drag({ x: start.x + start.width / 2, y: start.y + start.height / 2 }, centre('B', '2'));
    const at = h.document.selection;
    const corner = { row: at.cornerRow ?? at.row, column: at.cornerColumn ?? at.column };
    expect([Math.min(corner.row, at.anchorRow), Math.min(corner.column, at.anchorColumn)]).toEqual([1, 1]);
    expect([Math.max(corner.row, at.anchorRow), Math.max(corner.column, at.anchorColumn)]).toEqual([3, 3]);
    expect({ row: at.row, column: at.column }).toEqual({ row: 3, column: 3 });
  });

  it('widens the grips for a finger, and gives a mouse back its fill handle', async () => {
    h = await mount();
    type Node = ReturnType<Rendered['getByRole']>;
    const grip = (): Node => {
      for (let child = h.ui.getAllByRole('columnheader')[0].firstChild; child !== null; child = child.nextSibling) {
        if (child.properties.get('cursor') === 'col-resize') {
          return child;
        }
      }
      throw new Error('column A has no grip');
    };
    expect(h.ui.getLayout(grip()).width).toBe(8);
    await tap(centre('B', '2'));
    expect(h.ui.getLayout(grip()).width).toBe(24);

    const b3 = centre('B', '3');
    h.ui.fireEvent.pointerDown(b3.x, b3.y);
    h.ui.fireEvent.pointerUp(b3.x, b3.y);
    await settle();
    expect(h.ui.getLayout(grip()).width).toBe(8);
    expect(h.ui.getByRole('button', { name: 'Fill' })).toBeDefined();
    expect(h.ui.queryByRole('button', { name: 'Selection end' })).toBeNull();
  });
});
