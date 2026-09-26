import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';

import { COLUMN_WIDTH, GUTTER_WIDTH, HEADER_HEIGHT, ROW_HEIGHT } from './dimensions';
import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Zoom — Phase 22's exit.
 *
 * A zoomed sheet whose click, drag, fill and resize land on the cell
 * under the pointer at 50%, 100% and 200%. The failure being guarded
 * against is the one a transform over the paint would have: a sheet
 * that looks zoomed and answers the pointer as if it were not.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

let h: Harness;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) {
    await h.ui.settle();
    await h.served.settle();
  }
  await h.ui.settle();
}

async function mount(zoom: number, fill?: (document: SheetDocument) => void): Promise<void> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 900 });
  h = { ui, served, service, document };
  await settle();
  service.setZoom(zoom);
  await settle();
}

/** The middle of a cell on screen, from the grid's box and the zoomed sizes. */
function pointAt(zoom: number, row: number, column: number): { x: number; y: number } {
  const grid = h.ui.getLayout(h.ui.getByRole('grid'));
  return {
    x: grid.x + Math.round(GUTTER_WIDTH * zoom) + column * Math.round(COLUMN_WIDTH * zoom) + Math.round((COLUMN_WIDTH * zoom) / 2),
    y: grid.y + Math.round(HEADER_HEIGHT * zoom) + row * Math.round(ROW_HEIGHT * zoom) + Math.round((ROW_HEIGHT * zoom) / 2)
  };
}

describe.each([0.5, 1, 2])('at %s', zoom => {
  it('draws the columns and rows at that size', async () => {
    await mount(zoom);
    expect(h.ui.getByRole('columnheader', { name: 'B' })).toHaveBox({ width: Math.round(COLUMN_WIDTH * zoom) });
    expect(h.ui.getByRole('rowheader', { name: '2' })).toHaveBox({ height: Math.round(ROW_HEIGHT * zoom) });
  });

  it('selects the cell a click lands on', async () => {
    await mount(zoom);
    const at = pointAt(zoom, 3, 2);
    h.ui.fireEvent.pointerDown(at.x, at.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(at.x, at.y);
    await settle();
    expect(h.document.selection).toMatchObject({ row: 3, column: 2 });
  });

  it('sweeps to the cell a drag ends on', async () => {
    await mount(zoom);
    const from = pointAt(zoom, 1, 1);
    const to = pointAt(zoom, 4, 3);
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 10, from.y + 10, { buttons: 1 });
    h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(to.x, to.y);
    await settle();
    expect(h.document.selection).toEqual({ row: 1, column: 1, anchorRow: 1, anchorColumn: 1, cornerRow: 4, cornerColumn: 3 });
  });

  it('fills to the cell the handle is dragged to', async () => {
    await mount(zoom, d => d.setCell(0, 0, '=ROW()'));
    const handle = h.ui.getLayout(h.ui.getByRole('button', { name: 'Fill' }));
    const from = { x: handle.x + handle.width / 2, y: handle.y + handle.height / 2 };
    const to = pointAt(zoom, 3, 0);
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x, from.y + 10, { buttons: 1 });
    h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(to.x, to.y);
    await settle();
    expect(h.document.sheet.input(3, 0)).toBe('=ROW()');
    expect(h.document.sheet.input(4, 0)).toBe('');
  });

  it('keeps a column resized on screen at its size in the document', async () => {
    await mount(zoom);
    const box = h.ui.getLayout(h.ui.getByRole('columnheader', { name: 'A' }));
    const from = box.x + box.width - 2;
    const y = box.y + box.height / 2;
    h.ui.fireEvent.pointerDown(from, y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from + 6, y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from + 40, y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(from + 40, y);
    await settle();
    // Forty pixels on screen is forty at 100%, twenty at 200%.
    expect(h.document.columnWidths[0]).toBe(Math.round((Math.round(COLUMN_WIDTH * zoom) + 40) / zoom));
  });
});

describe('the zoom', () => {
  it('is the sheet’s own, and saved with it', async () => {
    const document = new SheetDocument();
    document.zoom = 1.5;
    document.addSheet('Second');
    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(snapshotOf(document)), 10)!);
    reopened.activate(0);
    expect(reopened.zoom).toBe(1.5);
    reopened.activate(1);
    expect(reopened.zoom).toBe(1);
  });

  it('is not an edit', async () => {
    await mount(1, d => d.setCell(0, 0, 'x'));
    const before = h.document.undoLabel;
    h.service.setZoom(2);
    expect(h.document.undoLabel).toBe(before);
  });

  it('is held between 50% and 200%', async () => {
    await mount(1);
    h.service.setZoom(9);
    expect(h.document.zoom).toBe(2);
    h.service.setZoom(0.1);
    expect(h.document.zoom).toBe(0.5);
  });

  it('steps from the status bar, and goes back to 100% from its figure', async () => {
    await mount(1);
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Zoom in' }));
    await settle();
    expect(h.document.zoom).toBe(1.1);
    expect(h.ui.getByRole('button', { name: 'Zoom to 100%' })).toHaveText('110%');

    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Zoom out' }));
    await settle();
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Zoom out' }));
    await settle();
    expect(h.document.zoom).toBe(0.9);

    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Zoom to 100%' }));
    await settle();
    expect(h.document.zoom).toBe(1);
  });

  it('doubles the columns while formulas are shown, and not the document’s widths', async () => {
    await mount(1.5);
    h.service.showFormulas(true);
    await settle();
    expect(h.ui.getByRole('columnheader', { name: 'B' })).toHaveBox({ width: Math.round(COLUMN_WIDTH * 1.5 * 2) });
    expect(h.ui.getByRole('rowheader', { name: '2' })).toHaveBox({ height: Math.round(ROW_HEIGHT * 1.5) });
    expect(h.document.columnWidths[1] ?? COLUMN_WIDTH).toBe(COLUMN_WIDTH);

    h.service.showFormulas(false);
    await settle();
    expect(h.ui.getByRole('columnheader', { name: 'B' })).toHaveBox({ width: Math.round(COLUMN_WIDTH * 1.5) });
  });

  it('keeps the selection in view when the grid is built again', async () => {
    await mount(1);
    h.service.setSelection(150, 0, 150, 0);
    await settle();
    h.service.setZoom(2);
    await settle();
    expect(h.ui.getByRole('rowheader', { name: '151' })).toBeDefined();
  });
});
