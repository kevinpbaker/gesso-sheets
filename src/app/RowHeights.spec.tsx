import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { GENERAL, PLAIN, type CellFormat } from '../sheet/Format';
import { MIN_ROW_HEIGHT, ROW_HEIGHT } from './dimensions';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Rows of different heights: wrapped text, and a row dragged.
 *
 * Asserted by the boxes the renderer lays out, on the rule Phase 3
 * learned: a height asserted as a property passes while the text it
 * was meant to hold is cut off. The application worker says which
 * rows could be taller, the render worker measures them, and the
 * answer crosses back — so every spec here settles both sides twice.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

const wrapped: CellFormat = { number: GENERAL, paint: { ...PLAIN, wrap: true } };
const LONG = 'a sentence long enough that a column of the default width has to break it over several lines';

async function mount(fill: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 400 });
  await service.restore();
  await settle({ ui, served, service, document });
  return { ui, served, service, document };
}

/** Both threads, twice: a question one way and its answer the other. */
async function settle(h: Harness): Promise<void> {
  for (let round = 0; round < 3; round++) {
    await h.ui.settle();
    await h.served.settle();
  }
  await h.ui.settle();
}

const boxOf = (h: Harness, text: string) => h.ui.getLayout(h.ui.getByRole('cell', { name: text }));

describe('a row as tall as what it holds', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  it('grows to hold wrapped text, and moves the rows below it down', async () => {
    h = await mount(d => {
      d.setFormat(0, 0, wrapped);
      d.setCell(0, 0, LONG);
      d.setCell(1, 0, 'below');
    });

    expect(boxOf(h, LONG).height).toBeGreaterThan(ROW_HEIGHT * 2);
    expect(boxOf(h, 'below').y).toBe(boxOf(h, LONG).y + boxOf(h, LONG).height);
    // And the number down the side is as tall as its row.
    expect(h.ui.getAllByRole('rowheader')[0]).toHaveBox({ height: boxOf(h, LONG).height });
  });

  it('leaves a row alone when the same text is not wrapped', async () => {
    h = await mount(d => d.setCell(0, 0, LONG));
    expect(boxOf(h, LONG).height).toBe(ROW_HEIGHT);
  });

  it('goes back down when the wrapped text is taken out', async () => {
    h = await mount(d => {
      d.setFormat(0, 0, wrapped);
      d.setCell(0, 0, LONG);
      d.setCell(1, 0, 'below');
    });
    h.service.setCell(0, 0, 'short');
    await settle(h);

    expect(boxOf(h, 'short').height).toBe(ROW_HEIGHT);
    expect(h.document.fittedRows.size).toBe(0);
  });

  it('grows when wrap is switched on, which is a format and not an edit', async () => {
    h = await mount(d => d.setCell(0, 0, LONG));
    h.service.format({ wrap: true });
    await settle(h);

    expect(boxOf(h, LONG).height).toBeGreaterThan(ROW_HEIGHT);
  });

  it('shrinks when its column is widened, since the text breaks less', async () => {
    h = await mount(d => {
      d.setFormat(0, 0, wrapped);
      d.setCell(0, 0, LONG);
    });
    const narrow = boxOf(h, LONG).height;
    h.service.setColumnWidth(0, 400);
    await settle(h);

    expect(boxOf(h, LONG).height).toBeLessThan(narrow);
  });

  it('keeps the fitted height in the file, so it opens at that height', async () => {
    h = await mount(d => {
      d.setFormat(0, 0, wrapped);
      d.setCell(0, 0, LONG);
    });
    const tall = boxOf(h, LONG).height;
    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(snapshotOf(h.document, 200)), 20)!);

    expect(reopened.fittedRows.get(0)).toBe(tall);
  });
});

describe('a row dragged to a height', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  /** Drags row 1's grip by `by` pixels. */
  function dragGrip(by: number): void {
    const box = h.ui.getLayout(h.ui.getAllByRole('rowheader')[0]);
    // Two inside the foot; the foot itself belongs to the next row.
    const x = box.x + box.width / 2;
    const from = box.y + box.height - 2;
    h.ui.fireEvent.pointerDown(x, from, { buttons: 1 });
    h.ui.fireEvent.pointerMove(x, from + 6, { buttons: 1 });
    h.ui.fireEvent.pointerMove(x, from + by, { buttons: 1 });
    h.ui.fireEvent.pointerUp(x, from + by);
  }

  it('takes the height it was dragged to, and tells the document', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'A1');
      d.setCell(1, 0, 'A2');
    });
    dragGrip(30);
    await settle(h);

    expect(boxOf(h, 'A1').height).toBe(ROW_HEIGHT + 30);
    expect(boxOf(h, 'A2').y).toBe(boxOf(h, 'A1').y + ROW_HEIGHT + 30);
    expect(h.document.rowHeights.get(0)).toBe(ROW_HEIGHT + 30);
  });

  it('resizes the row without selecting it', async () => {
    h = await mount(d => d.setCell(0, 0, 'A1'));
    h.service.setSelection(5, 3, 5, 3);
    await settle(h);
    dragGrip(30);
    await settle(h);

    expect(h.document.selection).toMatchObject({ row: 5, column: 3, anchorRow: 5, anchorColumn: 3 });
  });

  it('will not be dragged away to nothing', async () => {
    h = await mount(d => d.setCell(0, 0, 'A1'));
    dragGrip(-200);
    await settle(h);

    expect(boxOf(h, 'A1').height).toBe(MIN_ROW_HEIGHT);
  });

  /** Excel's rule: a row told its height keeps it, whatever goes in it. */
  it('keeps a height set by hand over one fitted to wrapped text', async () => {
    h = await mount(d => {
      d.setFormat(0, 0, wrapped);
      d.setCell(0, 0, LONG);
    });
    h.service.setRowHeight(0, 30);
    h.service.setCell(0, 0, `${LONG} ${LONG}`);
    await settle(h);

    expect(boxOf(h, `${LONG} ${LONG}`).height).toBe(30);

    // Until it is given back.
    h.service.fitRowsToContents(0, 0);
    await settle(h);
    expect(boxOf(h, `${LONG} ${LONG}`).height).toBeGreaterThan(ROW_HEIGHT * 2);
  });
});

describe('row heights and the rows around them', () => {
  const service = (): { service: SheetService; document: SheetDocument } => {
    const document = new SheetDocument();
    return { service: new SheetService(document, { rowCount: 50, columnCount: 5 }), document };
  };

  it('move down with an insert above them, and back with its undo', () => {
    const { service: s, document } = service();
    s.setRowHeight(4, 60);
    s.hideRows(6, 6);
    s.insertRows(2, 3);
    expect([...document.rowHeights]).toEqual([[7, 60]]);
    expect([...document.hiddenRows]).toEqual([9]);

    s.undo();
    expect([...document.rowHeights]).toEqual([[4, 60]]);
    // A hidden row came back to where it was too; it used to stay moved.
    expect([...document.hiddenRows]).toEqual([6]);

    s.redo();
    expect([...document.rowHeights]).toEqual([[7, 60]]);
  });

  it('go with a deleted row', () => {
    const { service: s, document } = service();
    s.setRowHeight(4, 60);
    s.deleteRows(4, 1);
    expect(document.rowHeights.size).toBe(0);
  });

  it('refuses a height from a file that no row can have', () => {
    const document = new SheetDocument();
    const snapshot = JSON.parse(JSON.stringify(snapshotOf(document, 50))) as { sheets: Record<string, unknown>[] };
    snapshot.sheets[0].rowHeights = [[1, 40], [2, 'tall'], [-1, 30], [3, 1e9], 'nonsense'];
    applySnapshot(document, parseSnapshot(JSON.stringify(snapshot), 5)!);
    expect([...document.rowHeights]).toEqual([
      [1, 40],
      [3, 545]
    ]);
  });

  it('drops an answer to a question that has since been replaced', () => {
    const { service: s, document } = service();
    document.setFormat(0, 0, wrapped);
    s.setCell(0, 0, LONG);
    let serial = 0;
    s.rowFit.subscribe(fit => (serial = fit.serial));
    const stale = serial;
    s.setCell(0, 0, `${LONG}!`);
    s.fitRows(stale, [[0, 90]]);
    expect(document.fittedRows.size).toBe(0);
    s.fitRows(serial, [[0, 90]]);
    expect(document.fittedRows.get(0)).toBe(90);
  });
});
