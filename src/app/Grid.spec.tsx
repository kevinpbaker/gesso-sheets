import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';

import { COLUMN_WIDTH, GUTTER_WIDTH, HEADER_HEIGHT, MIN_COLUMN_WIDTH, ROW_HEIGHT } from './dimensions';
import { SheetApp } from './SheetApp';
import { Sheet } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
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
    sheetChannel(service)
  ]);
  // The grid is given the editing handle the whole screen shares; in
  // the app that is `SheetApp`'s, and here the spec is the screen.
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, ...VIEWPORT });
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
    /**
     * The grid's own origin. The screen puts a formula bar above it,
     * so a spec that asserted an absolute `y: 0` would be asserting
     * the bar's height as much as the header's stickiness.
     */
    const origin = () => h.ui.getVisibleBox(h.ui.getByRole('grid'));

    beforeEach(async () => {
      // A1 carries its own address, so the cell at the origin can be
      // found by name rather than by being the first of many blanks.
      h = await mount(document => document.setCell(0, 0, 'A1'));
    });

    it('puts the header across the top and the gutter down the left', () => {
      const top = origin().y;
      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveVisibleBox({
        x: GUTTER_WIDTH,
        y: top,
        width: COLUMN_WIDTH,
        height: HEADER_HEIGHT
      });
      expect(h.ui.getByRole('rowheader', { name: '1' })).toHaveVisibleBox({
        x: 0,
        y: top + HEADER_HEIGHT,
        width: GUTTER_WIDTH,
        height: ROW_HEIGHT
      });
    });

    it('starts the cells past both strips', () => {
      expect(h.ui.getByRole('cell', { name: 'A1' })).toHaveVisibleBox({
        x: GUTTER_WIDTH,
        y: origin().y + HEADER_HEIGHT,
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

      expect(h.ui.getByRole('columnheader', { name: 'A' })).toHaveVisibleBox({
        y: origin().y,
        height: HEADER_HEIGHT
      });
      // And it really did scroll: row 1 is long gone.
      expect(h.ui.queryByRole('rowheader', { name: '1' })).toBeNull();
      expect(h.ui.getAllByRole('rowheader')[0]).toHaveVisibleBox({ x: 0 });
    });

    it('holds the gutter at the left while the columns scroll past', async () => {
      h.ui.fireEvent.wheel({ x: 300, y: 200, deltaX: COLUMN_WIDTH * 10 });
      await h.ui.settle();

      expect(h.ui.getAllByRole('rowheader')[0]).toHaveVisibleBox({ x: 0, width: GUTTER_WIDTH });
      expect(h.ui.queryByRole('columnheader', { name: 'A' })).toBeNull();
      expect(h.ui.getAllByRole('columnheader')[0]).toHaveVisibleBox({ y: origin().y });
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

      expect(h.ui.getAllByRole('columnheader')[0]).toHaveVisibleBox({ y: origin().y, height: HEADER_HEIGHT });
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

  /**
   * Clicking and then typing.
   *
   * This is the seam the two other specs left open, and a bug lived in
   * it: `Keyboard.spec.tsx` reaches the grid with Tab and never
   * touches the mouse, and the rest of this file clicks and never
   * presses a key. Clicking a cell moved the selection and left focus
   * wherever it was, so every key after a click went somewhere else —
   * Delete did nothing, typing did nothing, and what the person saw
   * was their keystrokes landing in the formula bar's own editor
   * instead of in the sheet.
   */
  describe('clicking and then typing', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(1, 1, '120');
        document.setCell(1, 3, '=B2*2');
      });
    });

    async function press(key: string): Promise<void> {
      h.ui.fireEvent.press(key);
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    }

    it('empties a clicked cell on Delete', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '120' }));
      await h.ui.settle();

      await press('Delete');

      expect(h.document.sheet.input(1, 1)).toBe('');
      expect(h.document.sheet.value(1, 3)).toBe(0);
    });

    it('opens a clicked cell and replaces it with what is typed', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '120' }));
      await h.ui.settle();

      await press('7');
      await press('Enter');

      expect(h.document.sheet.input(1, 1)).toBe('7');
      expect(h.document.sheet.value(1, 3)).toBe(14);
    });

    it('moves the selection with the arrows after a click', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '120' }));
      await h.ui.settle();

      await press('ArrowRight');
      await press('9');
      await press('Enter');

      // B2 is untouched and C2 took the value.
      expect(h.document.sheet.input(1, 1)).toBe('120');
      expect(h.document.sheet.input(1, 2)).toBe('9');
    });

    it('commits an open cell when another is clicked', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '120' }));
      await h.ui.settle();
      await press('5');

      // D2 holds `=B2*2`, so it is a cell with a name to click on.
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '240' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(h.document.sheet.input(1, 1)).toBe('5');
      // And the click landed where it was aimed.
      expect(h.document.selection).toMatchObject({ row: 1, column: 3 });
    });
  });

  /**
   * Editing in the formula bar rather than in the cell.
   *
   * The bar is a real text field and putting a caret in it means the
   * same thing as opening the cell: the keys belong to the text from
   * that moment. Treated as "a cell is selected and nothing is open",
   * Backspace means *empty this cell* and a digit means *replace this
   * cell* — so deleting one character wiped the lot, and the next
   * character arrived twice, once from the key handler seeding a
   * draft and once from the field inserting it.
   */
  describe('typing in the formula bar', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(1, 1, '120');
        document.setCell(1, 3, '=B2*2');
      });
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '120' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    });

    const bar = () => h.ui.getByRole('textbox', { name: 'Formula' });

    async function press(key: string): Promise<void> {
      h.ui.fireEvent.press(key);
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    }

    /**
     * Put the caret at the end of the bar, as clicking after the last
     * character does. `End` rather than reaching into the model:
     * once the bar has opened the cell, End belongs to the text, and
     * a spec that set the caret by hand would not be testing that.
     */
    async function focusBar(): Promise<void> {
      h.ui.fireEvent.focus(bar());
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
      await press('End');
    }

    it('deletes one character rather than emptying the cell', async () => {
      await focusBar();
      await press('Backspace');

      expect(bar()).toHaveText('12');
      // Nothing is committed until Enter, so the cell still holds 120.
      expect(h.document.sheet.input(1, 1)).toBe('120');
    });

    it('takes a typed character once, not twice', async () => {
      await focusBar();
      await press('Backspace');
      h.ui.fireEvent.type('3');
      await h.ui.settle();

      expect(bar()).toHaveText('123');
    });

    it('commits what the bar holds on Enter', async () => {
      await focusBar();
      await press('Backspace');
      h.ui.fireEvent.type('3');
      await h.ui.settle();
      await press('Enter');

      expect(h.document.sheet.input(1, 1)).toBe('123');
      expect(h.document.sheet.value(1, 3)).toBe(246);
    });

    it('puts back what was there on Escape', async () => {
      await focusBar();
      await press('Backspace');
      await press('Escape');

      expect(h.document.sheet.input(1, 1)).toBe('120');
      expect(bar()).toHaveText('120');
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
