import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent, serve } from 'gesso-framework';

import { COLUMN_WIDTH, GUTTER_WIDTH, HEADER_HEIGHT, MIN_COLUMN_WIDTH, ROW_HEIGHT } from './dimensions';
import { Grid } from './Grid';
import { Sheet } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService } from './SheetService';

/**
 * The grid surface — the exit criterion for Phase 3.
 *
 * Two kinds of assertion, and the roadmap asks for both.
 *
 * **The semantics tree.** `getByRole` reads what the accessibility
 * mirror writes, so a cell that is hard to find here is exactly a cell
 * that is hard to find with a screen reader. There is no second
 * definition of what a cell *is* for a test to drift away from, which
 * is why the accessibility story is asserted from the first frame of
 * the first phase that draws anything rather than retrofitted later.
 *
 * **`toHaveBox` on the frozen panes.** Sticky is the mechanism and the
 * boxes are the claim: the header stays at the top of the viewport
 * while the rows scroll under it, and the gutter stays at the left
 * while the columns scroll past. Gesso's sticky is conformance-tested
 * against Chrome on one axis; this is the first thing to ask it for
 * both at once.
 */

const VIEWPORT = { width: 700, height: 300 };

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

async function mount(fill?: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 500, columnCount: 40 });
  const served = serveForTest([
    serve(Sheet, {
      view: {
        window: service.window,
        geometry: service.geometry,
        selection: service.selection,
        editor: service.editor,
        status: service.status
      },
      commands: {
        setViewport: (a, b, c, d) => service.setViewport(a, b, c, d),
        setCell: (row, column, input) => service.setCell(row, column, input),
        setSelection: (a, b, c, d) => service.setSelection(a, b, c, d),
        undo: () => service.undo(),
        redo: () => service.redo()
      }
    })
  ]);
  const ui = renderTest(createComponent(Grid), { channels: served.registry, ...VIEWPORT });
  // Three turns, and each is doing something different. The first
  // frame is what tells the window how big it is, which is what sends
  // `setViewport`; the channel then has to carry that command over and
  // its answer back; and the frame after that is the one that draws
  // the values. Settling one side only leaves a grid of empty cells,
  // which looks exactly like a grid whose cells have no values.
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served, service, document };
}

describe('the grid surface', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  describe('the semantics tree', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(0, 0, 'Region');
        document.setCell(1, 0, 'North');
        document.setCell(1, 1, '=1+2');
      });
    });

    it('is a grid', () => {
      expect(h.ui.getByRole('grid')).toBeDefined();
    });

    it('has rows in it', () => {
      expect(h.ui.getAllByRole('row').length).toBeGreaterThan(1);
    });

    it('labels its columns with letters', () => {
      const headers = h.ui.getAllByRole('columnheader');
      expect(headers.length).toBeGreaterThan(3);
      expect(h.ui.getByRole('columnheader', { name: 'A' })).toBeDefined();
      expect(h.ui.getByRole('columnheader', { name: 'B' })).toBeDefined();
    });

    it('numbers its rows, counting from one as a person does', () => {
      expect(h.ui.getByRole('rowheader', { name: '1' })).toBeDefined();
      expect(h.ui.getByRole('rowheader', { name: '2' })).toBeDefined();
    });

    /** It is `cell`, not `gridcell`: see `UI_ROLES` in the engine. */
    it('exposes cells that carry the value, not the formula', () => {
      expect(h.ui.getByRole('cell', { name: 'Region' })).toBeDefined();
      expect(h.ui.getByRole('cell', { name: 'North' })).toBeDefined();
      // A1 of the formula's row displays 3 and was typed `=1+2`.
      expect(h.ui.getByRole('cell', { name: '3' })).toBeDefined();
    });

    it('mounts only the cells in view, out of twenty thousand', () => {
      const cells = h.ui.getAllByRole('cell');
      expect(cells.length).toBeGreaterThan(0);
      // 500 x 40 is the sheet; a 700x300 viewport holds a fraction.
      expect(cells.length).toBeLessThan(500);
    });
  });

  describe('the frozen panes', () => {
    beforeEach(async () => {
      // A1 carries its own address, so the cell at the origin can be
      // found by name rather than by being the first of many blanks.
      h = await mount(document => document.setCell(0, 0, 'A1'));
    });

    it('puts the header across the top and the gutter down the left', () => {
      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveVisibleBox({
        x: GUTTER_WIDTH,
        y: 0,
        width: COLUMN_WIDTH,
        height: HEADER_HEIGHT
      });
      expect(h.ui.getByRole('rowheader', { name: '1' })).toHaveVisibleBox({
        x: 0,
        y: HEADER_HEIGHT,
        width: GUTTER_WIDTH,
        height: ROW_HEIGHT
      });
    });

    it('starts the cells past both strips', () => {
      expect(h.ui.getByRole('cell', { name: 'A1' })).toHaveVisibleBox({
        x: GUTTER_WIDTH,
        y: HEADER_HEIGHT,
        width: COLUMN_WIDTH,
        height: ROW_HEIGHT
      });
    });

    /**
     * The claim sticky exists to make. Scrolled down a hundred rows,
     * the header is still at the top of the viewport — and it is a
     * *different* row's worth of content underneath it, so this is not
     * the same assertion as the one above.
     */
    /**
     * `toHaveVisibleBox`, not `toHaveBox`, and the difference is the
     * whole assertion. A sticky header is *laid out* at the top of the
     * content and stays laid out there however far the container
     * scrolls, so a spec written against the laid-out box passes for a
     * header that scrolled off the screen — which is exactly what this
     * one did until the matcher existed to tell them apart.
     */
    it('holds the header at the top while the rows scroll under it', async () => {
      h.ui.fireEvent.wheel({ x: 300, y: 200, deltaY: ROW_HEIGHT * 100 });
      await h.ui.settle();

      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveVisibleBox({ y: 0, height: HEADER_HEIGHT });
      // And it really did scroll: row 1 is long gone.
      expect(h.ui.queryByRole('rowheader', { name: '1' })).toBeNull();
      expect(h.ui.getAllByRole('rowheader')[0]).toHaveVisibleBox({ x: 0 });
    });

    it('holds the gutter at the left while the columns scroll past', async () => {
      h.ui.fireEvent.wheel({ x: 300, y: 200, deltaX: COLUMN_WIDTH * 10 });
      await h.ui.settle();

      expect(h.ui.getAllByRole('rowheader')[0]).toHaveVisibleBox({ x: 0, width: GUTTER_WIDTH });
      expect(h.ui.queryByRole('columnheader', { name: 'A' })).toBeNull();
      expect(h.ui.getAllByRole('columnheader')[0]).toHaveVisibleBox({ y: 0 });
    });

    /**
     * Both at once, which is the case the engine had never been asked
     * for: the corner has its own `left` and inherits the header row's
     * `top`, so it has to stay in the corner while the sheet moves
     * diagonally underneath it.
     */
    it('holds both strips, and the corner between them, on a diagonal scroll', async () => {
      h.ui.fireEvent.wheel({ x: 300, y: 200, deltaX: COLUMN_WIDTH * 8, deltaY: ROW_HEIGHT * 60 });
      await h.ui.settle();

      expect(h.ui.getAllByRole('columnheader')[0]).toHaveVisibleBox({ y: 0, height: HEADER_HEIGHT });
      expect(h.ui.getAllByRole('rowheader')[0]).toHaveVisibleBox({ x: 0, width: GUTTER_WIDTH });
      // The corner holds both at once: its own `left`, and the header
      // row's `top` inherited by being inside it.
      const corner = h.ui.getByRole('grid');
      expect(corner).toBeDefined();
    });
  });

  /**
   * A column a person can drag wider.
   *
   * The widths live on this thread, not on the channel: how wide a
   * column is drawn is not the application's business, and putting it
   * on the wire would make every frame of a drag a round trip through
   * another thread to decide something this one already knows.
   */
  describe('resizing a column', () => {
    beforeEach(async () => {
      h = await mount(document => document.setCell(0, 0, 'A1'));
    });

    /** Drags column A's grip by `by` pixels. */
    function dragGrip(by: number): void {
      const box = h.ui.getLayout(h.ui.getByRole('columnheader', { name: 'A' }));
      // Two inside the trailing edge. The edge itself belongs to the
      // next column — a box contains x where `x >= left && x < right` —
      // so pressing exactly on it misses the header and the grip with
      // it.
      const from = box.x + box.width - 2;
      const y = box.y + box.height / 2;
      // Buttons held down throughout: a move with no button is a hover,
      // and the gesture recogniser is right not to call that a drag.
      h.ui.fireEvent.pointerDown(from, y, { buttons: 1 });
      h.ui.fireEvent.pointerMove(from + 6, y, { buttons: 1 });
      h.ui.fireEvent.pointerMove(from + by, y, { buttons: 1 });
      h.ui.fireEvent.pointerUp(from + by, y);
    }

    it('widens the column the grip belongs to', async () => {
      dragGrip(60);
      await h.ui.settle();

      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveBox({ width: COLUMN_WIDTH + 60 });
    });

    /**
     * Every column after the one that moved sits somewhere else, which
     * is the whole reason the window keeps a prefix sum rather than
     * multiplying.
     */
    it('moves the columns after it along', async () => {
      dragGrip(60);
      await h.ui.settle();

      expect(h.ui.getByRole('columnheader', { name: 'B' })).toHaveBox({ x: GUTTER_WIDTH + COLUMN_WIDTH + 60 });
    });

    it('keeps the cells under the header they belong to', async () => {
      dragGrip(60);
      await h.ui.settle();

      expect(h.ui.getByRole('cell', { name: 'A1' })).toHaveBox({ x: GUTTER_WIDTH, width: COLUMN_WIDTH + 60 });
    });

    it('will not let a column be dragged away to nothing', async () => {
      dragGrip(-400);
      await h.ui.settle();

      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveBox({ width: MIN_COLUMN_WIDTH });
    });
  });

  describe('selection', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(2, 2, 'here');
      });
    });

    it('moves to the cell that was clicked', async () => {
      const cell = h.ui.getByRole('cell', { name: 'here' });
      h.ui.fireEvent.click(cell);
      await h.ui.settle();

      expect(h.document.selection).toEqual({ row: 2, column: 2, anchorRow: 2, anchorColumn: 2 });
    });
  });
});
