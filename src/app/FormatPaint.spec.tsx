import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { GENERAL, PLAIN, type CellFormat } from '../sheet/Format';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * That the grid actually paints what the palette says.
 *
 * This file exists because of what it did not exist for. Phase 9's
 * other specs assert the *document* — `formatAt(0, 0).paint.bold` —
 * and every one of them passed while a browser showed a sheet with no
 * bold in it anywhere: the palette was never published on the load
 * path, so every cell pointed at an entry the render worker had never
 * been sent and fell back to plain. The number formats looked right
 * throughout, because those are applied on the application thread and
 * cross as finished strings.
 *
 * So the assertions here read the properties the renderer will draw
 * with, on a sheet whose formats arrived the way a real one's do:
 * already in the document when the channel opens.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
}

const bold: CellFormat = { number: GENERAL, paint: { ...PLAIN, bold: true, fill: '#eef2f7' } };
const italic: CellFormat = { number: GENERAL, paint: { ...PLAIN, italic: true, color: '#b00020' } };
const centred: CellFormat = { number: GENERAL, paint: { ...PLAIN, align: 'center' } };

async function mount(fill: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 300 });
  // The load path, which is the one that was broken.
  await service.restore();
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served };
}

describe('what the grid draws a formatted cell with', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  /**
   * The cell node holding a value.
   *
   * By role and not by text: the formula bar shows the active cell's
   * input too, so `getByText` finds two nodes and the one it is not
   * looking for is the one that never had a format.
   */
  const propertyOf = (text: string, name: string): unknown =>
    h.ui.getByRole('cell', { name: text }).properties.get(name);

  it('draws a bold cell bold', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'Region');
      d.setFormat(0, 0, bold);
      d.setCell(1, 0, 'North');
    });

    expect(propertyOf('Region', 'fontWeight')).toBe('bold');
    expect(propertyOf('North', 'fontWeight')).toBe('normal');
  });

  it('fills a cell that has a fill', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'Region');
      d.setFormat(0, 0, bold);
    });
    expect(propertyOf('Region', 'backgroundColor')).toBe('#eef2f7');
  });

  it('colours and italicises', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'Loss');
      d.setFormat(0, 0, italic);
    });
    expect(propertyOf('Loss', 'fontStyle')).toBe('italic');
    expect(propertyOf('Loss', 'color')).toBe('#b00020');
  });

  it('aligns where it is told, and by the value where it is not', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'Middle');
      d.setFormat(0, 0, centred);
      d.setCell(1, 0, '42');
      d.setCell(2, 0, 'text');
    });
    expect(propertyOf('Middle', 'textAlign')).toBe('center');
    // Untouched: numbers right, text left, which is the spreadsheet
    // rule and a real alignment rather than an absent one.
    expect(propertyOf('42', 'textAlign')).toBe('right');
    expect(propertyOf('text', 'textAlign')).toBe('start');
  });

  it('shows a number under its number format', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '1234.5');
      d.setFormat(0, 0, { number: { kind: 'currency', places: 2, symbol: '$' }, paint: PLAIN });
    });
    expect(h.ui.getByRole('cell', { name: '$1,234.50' })).toBeDefined();
  });

  /**
   * A format applied while the sheet is open has to reach the screen
   * too, not only one that was there when the channel opened.
   */
  it('repaints when a format is applied', async () => {
    h = await mount(d => d.setCell(0, 0, 'Region'));
    expect(propertyOf('Region', 'fontWeight')).toBe('normal');

    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid && stops++ < 10) {
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('b', { ctrl: true });
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    expect(propertyOf('Region', 'fontWeight')).toBe('bold');
  });
});
