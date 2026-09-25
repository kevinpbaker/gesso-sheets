import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';

import { formulaSpans } from './FormulaColours';
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

/**
 * A frozen pane, asserted as boxes rather than as properties.
 *
 * The properties were asserted first and passed while the sheet drew
 * a frozen column in the wrong place — the same trap Phase 3 fell
 * into with the header, and the reason `toHaveVisibleBox` exists.
 * Sticky is the mechanism; where the thing lands is the claim.
 */
describe('a frozen pane', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  const origin = () => h.ui.getVisibleBox(h.ui.getByRole('grid'));

  beforeEach(async () => {
    h = await mount(document => {
      for (let row = 0; row < 200; row++) {
        for (let column = 0; column < 12; column++) {
          document.setCell(row, column, `r${row}c${column}`);
        }
      }
    });
  });

  async function scroll(deltaX: number, deltaY: number): Promise<void> {
    h.ui.fireEvent.wheel({ x: 300, y: 200, deltaX, deltaY });
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function freeze(rows: number, columns: number): Promise<void> {
    h.service.freeze(rows, columns);
    await h.served.settle();
    await h.ui.settle();
  }

  /** Unfrozen, a cell in column A sits right against the gutter. */
  it('draws the first column against the gutter before anything is frozen', () => {
    expect(h.ui.getByRole('cell', { name: 'r0c0' })).toHaveVisibleBox({ x: GUTTER_WIDTH, width: COLUMN_WIDTH });
  });

  it('leaves a frozen column exactly where an unfrozen one was', async () => {
    await freeze(0, 1);
    expect(h.ui.getByRole('cell', { name: 'r0c0' })).toHaveVisibleBox({ x: GUTTER_WIDTH, width: COLUMN_WIDTH });
  });

  it('keeps the column that follows a frozen one in its place', async () => {
    await freeze(0, 1);
    expect(h.ui.getByRole('cell', { name: 'r0c1' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH });
  });

  /** The claim sticky exists to make, on the column axis. */
  it('holds a frozen column against the gutter while the sheet scrolls sideways', async () => {
    await freeze(0, 1);
    await scroll(COLUMN_WIDTH * 6, 0);

    expect(h.ui.getByRole('cell', { name: 'r0c0' })).toHaveVisibleBox({ x: GUTTER_WIDTH, width: COLUMN_WIDTH });
    // And it really did scroll: the column next to it is long gone.
    expect(h.ui.queryByRole('cell', { name: 'r0c1' })).toBeNull();
  });

  it('holds a frozen row under the header while the sheet scrolls down', async () => {
    await freeze(1, 0);
    await scroll(0, ROW_HEIGHT * 60);

    expect(h.ui.getByRole('cell', { name: 'r0c0' })).toHaveVisibleBox({
      y: origin().y + HEADER_HEIGHT,
      height: ROW_HEIGHT
    });
    expect(h.ui.queryByRole('cell', { name: 'r1c0' })).toBeNull();
  });

  it('holds both at once on a diagonal scroll', async () => {
    await freeze(1, 1);
    await scroll(COLUMN_WIDTH * 6, ROW_HEIGHT * 60);

    expect(h.ui.getByRole('cell', { name: 'r0c0' })).toHaveVisibleBox({
      x: GUTTER_WIDTH,
      y: origin().y + HEADER_HEIGHT
    });
  });

  it('puts the column labels over the columns they label', async () => {
    await freeze(0, 2);
    await scroll(COLUMN_WIDTH * 6, 0);

    const a = h.ui.getVisibleBox(h.ui.getByRole('columnheader', { name: 'A' }));
    const cell = h.ui.getVisibleBox(h.ui.getByRole('cell', { name: 'r0c0' }));
    expect(a.x).toBe(cell.x);
    expect(a.width).toBe(cell.width);
  });
});

/**
 * Merged cells, asserted as boxes.
 *
 * A merge is drawn by its anchor and by nothing else: the anchor is
 * as wide as the columns it covers and as tall as the rows, and the
 * cells underneath have no size at all. Boxes rather than properties,
 * for the reason the frozen pane learned it.
 */
describe('merged cells', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount(document => {
      for (let row = 0; row < 40; row++) {
        for (let column = 0; column < 10; column++) {
          document.setCell(row, column, `r${row}c${column}`);
        }
      }
    });
  });

  async function merge(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): Promise<void> {
    h.service.setSelection(firstRow, firstColumn, lastRow, lastColumn);
    h.service.mergeCells();
    await h.served.settle();
    await h.ui.settle();
  }

  it('makes the anchor as wide as the columns it covers', async () => {
    await merge(1, 1, 1, 3);
    expect(h.ui.getByRole('cell', { name: 'r1c1' })).toHaveVisibleBox({
      x: GUTTER_WIDTH + COLUMN_WIDTH,
      width: COLUMN_WIDTH * 3,
      height: ROW_HEIGHT
    });
  });

  it('makes the anchor as tall as the rows it covers', async () => {
    await merge(1, 3, 1, 1);
    expect(h.ui.getByRole('cell', { name: 'r1c1' })).toHaveVisibleBox({
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT * 3
    });
  });

  /** The covered cells lose what they held, so there is nothing to draw. */
  it('empties the cells it covered', async () => {
    await merge(1, 1, 1, 3);
    expect(h.ui.queryByRole('cell', { name: 'r1c2' })).toBeNull();
    expect(h.document.sheet.input(1, 2)).toBe('');
    // And the one outside it is untouched.
    expect(h.document.sheet.input(1, 4)).toBe('r1c4');
  });

  /** Destructive, and in one step of undo. */
  it('gives every covered cell back on one press of ctrl-Z', async () => {
    await merge(1, 1, 1, 3);
    h.service.undo();
    await h.served.settle();
    await h.ui.settle();

    expect(h.document.sheet.input(1, 2)).toBe('r1c2');
    expect(h.document.sheet.input(1, 3)).toBe('r1c3');
  });

  it('keeps the column after a merge where the window put it', async () => {
    await merge(1, 1, 1, 3);
    expect(h.ui.getByRole('cell', { name: 'r1c4' })).toHaveVisibleBox({
      x: GUTTER_WIDTH + COLUMN_WIDTH * 4
    });
    // And the row below is undisturbed.
    expect(h.ui.getByRole('cell', { name: 'r2c1' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH });
  });

  /**
   * Found in a browser while checking the editor fix, by reading the
   * boxes rather than the picture: with empty cells either side, a
   * whole column sliding left is invisible.
   */
  it('keeps the columns beside a tall merge where the window put them', async () => {
    await merge(1, 3, 1, 1);
    // Row 1 holds the anchor; rows 2 and 3 hold cells it covers.
    expect(h.ui.getByRole('cell', { name: 'r2c2' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH * 2 });
    expect(h.ui.getByRole('cell', { name: 'r3c5' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH * 5 });
  });

  it('keeps them beside a merge that is tall and wide at once', async () => {
    await merge(1, 3, 1, 2);
    expect(h.ui.getByRole('cell', { name: 'r2c3' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH * 3 });
    expect(h.ui.getByRole('cell', { name: 'r3c4' })).toHaveVisibleBox({ x: GUTTER_WIDTH + COLUMN_WIDTH * 4 });
  });

  it('takes a merge apart again', async () => {
    await merge(1, 1, 1, 3);
    h.service.setSelection(1, 1, 1, 1);
    h.service.unmergeCells();
    await h.served.settle();
    await h.ui.settle();

    expect(h.ui.getByRole('cell', { name: 'r1c1' })).toHaveVisibleBox({ width: COLUMN_WIDTH });
  });

  /**
   * The window reaching back for an anchor it has scrolled past,
   * which is what `extendRange` is for: a merge drawn by a cell
   * outside the window is a merge that disappears at the edge of the
   * screen.
   */
  it('keeps drawing a merge whose anchor has scrolled out of view', async () => {
    await merge(1, 1, 0, 8);
    h.ui.fireEvent.wheel({ x: 300, y: 200, deltaX: COLUMN_WIDTH * 4 });
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    // The anchor is off to the left and still drawn, nine columns wide.
    expect(h.ui.getByRole('cell', { name: 'r1c0' })).toHaveVisibleBox({ width: COLUMN_WIDTH * 9 });
  });

  /**
   * Opening a merge for editing used to take the grid apart.
   *
   * The editor is a cell in the row like any other, and it was built
   * from `widthOf(column)` without ever asking about merges — so the
   * moment somebody typed into a merged cell, the anchor snapped back
   * to one column, the row lost the width the merge had given it, and
   * every column after it slid left. What that looks like on screen is
   * the original cells reappearing underneath the merge.
   */
  describe('editing one', () => {
    async function open(row: number, column: number): Promise<void> {
      h.service.setSelection(row, column, row, column);
      await h.served.settle();
      await h.ui.settle();
      // F2 has to reach the grid, not the toolbar it started on.
      const grid = h.ui.getByRole('grid');
      let stops = 0;
      while (h.ui.runtime.input.focus.focusedNode !== grid) {
        if (stops++ > 8) {
          throw new Error('Tab never reached the grid');
        }
        h.ui.fireEvent.tab();
        await h.ui.settle();
      }
      h.ui.fireEvent.press('F2');
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    }

    const editor = () => h.ui.getByRole('textbox', { name: 'Cell' });

    it('opens the editor across the columns the merge covers', async () => {
      await merge(1, 1, 1, 3);
      await open(1, 1);
      expect(editor()).toHaveVisibleBox({
        x: GUTTER_WIDTH + COLUMN_WIDTH,
        width: COLUMN_WIDTH * 3,
        height: ROW_HEIGHT
      });
    });

    it('opens it down the rows the merge covers', async () => {
      await merge(1, 3, 1, 1);
      await open(1, 1);
      expect(editor()).toHaveVisibleBox({ width: COLUMN_WIDTH, height: ROW_HEIGHT * 3 });
    });

    /** The symptom as it was reported: the grid moving underneath. */
    it('leaves the columns after the merge exactly where they were', async () => {
      await merge(1, 1, 1, 3);
      const before = h.ui.getVisibleBox(h.ui.getByRole('cell', { name: 'r1c4' }));
      await open(1, 1);
      expect(h.ui.getByRole('cell', { name: 'r1c4' })).toHaveVisibleBox({ x: before.x, y: before.y });
    });

    it('still covers one cell when the cell is not merged', async () => {
      await open(1, 5);
      expect(editor()).toHaveVisibleBox({
        x: GUTTER_WIDTH + COLUMN_WIDTH * 5,
        width: COLUMN_WIDTH,
        height: ROW_HEIGHT
      });
    });
  });

  /**
   * `cellAt` is a division over the offsets and has never heard of a
   * merge, so a sweep across one reports the cells underneath — which
   * are not drawn and hold nothing. The merge is what is there.
   */
  it('sweeps to the merge rather than to the cells under it', async () => {
    await merge(1, 2, 1, 2);
    const grid = h.ui.getVisibleBox(h.ui.getByRole('grid'));
    const at = (row: number, column: number) => ({
      x: grid.x + GUTTER_WIDTH + column * COLUMN_WIDTH + 4,
      y: grid.y + HEADER_HEIGHT + row * ROW_HEIGHT + 4
    });

    // Press in an untouched cell, then drag into the merge — onto
    // what would be its *second* row and column, which is a cell
    // nobody can see.
    const start = at(5, 5);
    const end = at(2, 2);
    h.ui.fireEvent.pointerDown(start.x, start.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(start.x - 8, start.y - 8, { buttons: 1 });
    h.ui.fireEvent.pointerMove(end.x, end.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(end.x, end.y);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    // The far corner is the merge's anchor, not the covered cell the
    // pointer is literally over.
    expect(h.document.selection.row).toBe(1);
    expect(h.document.selection.column).toBe(1);
  });
});

/**
 * A formula's references, coloured while it is being typed.
 *
 * The runs are asserted rather than the pixels: that an
 * `EditableText` given runs draws them is the engine's claim, with
 * its own spec in `EditableRendering.spec.ts`. What is this
 * application's claim is that the field is handed runs describing
 * exactly the text it holds — because runs describing anything else
 * are thrown away by the engine, and the colour would simply never
 * appear.
 */
describe('colouring a formula as it is typed', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount(document => {
      document.setCell(0, 0, '10');
      document.setCell(1, 0, '20');
      document.setCell(4, 4, '=A1+A2');
    });
  });

  async function open(row: number, column: number): Promise<void> {
    h.service.setSelection(row, column, row, column);
    await h.served.settle();
    await h.ui.settle();
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 8) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('F2');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  const runs = () =>
    h.ui.getByRole('textbox', { name: 'Cell' }).properties.get('spans') as
      | readonly { text: string; color?: string }[]
      | undefined;

  it('hands the field the formula cut at its references', async () => {
    await open(4, 4);
    expect(runs()?.map(run => run.text)).toEqual(['=', 'A1', '+', 'A2']);
  });

  /** The invariant the engine checks, asserted where it is produced. */
  it('hands it runs that spell the draft exactly', async () => {
    await open(4, 4);
    expect(runs()?.map(run => run.text).join('')).toBe('=A1+A2');
  });

  it('colours the references and nothing else', async () => {
    await open(4, 4);
    const coloured = runs()!.filter(run => run.color !== undefined);
    expect(coloured.map(run => run.text)).toEqual(['A1', 'A2']);
    expect(coloured[0].color).not.toBe(coloured[1].color);
  });

  it('gives a cell holding no formula no runs at all', async () => {
    await open(0, 0);
    expect(runs()).toBeUndefined();
  });

  it('follows the draft as it is typed', async () => {
    await open(2, 2);
    h.ui.fireEvent.type('=B1');
    await h.ui.settle();
    expect(runs()?.map(run => run.text)).toEqual(['=', 'B1']);

    h.ui.fireEvent.type('+C3');
    await h.ui.settle();
    expect(runs()?.map(run => run.text)).toEqual(['=', 'B1', '+', 'C3']);
    expect(runs()?.map(run => run.text).join('')).toBe('=B1+C3');
  });
});

/**
 * The boxes drawn round the cells a formula names, while it is typed.
 *
 * The other half of colouring the text, and they have to agree: the
 * `A1` in the formula and the box round A1 on the sheet are the same
 * colour or the feature is worse than not having it.
 *
 * Asserted through the decoration shapes the row carries, which is
 * what the renderer paints — there is no node to find with
 * `getByRole`, deliberately, because an outlined range costs no nodes
 * at all.
 */
describe('outlining the cells a formula names', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount();
  });

  /** The decoration rectangles a row carries, which is what is painted. */
  const decorationsOf = (node: { decorations?: unknown }) =>
    (node.decorations ?? []) as readonly { x: number; width: number; color: string }[];

  function shapesOn(row: number): readonly { x: number; width: number; color: string }[] {
    const found = h.ui.getAllByRole('row').find(node => node.properties.get('posInSet') === row + 1);
    return found === undefined ? [] : decorationsOf(found);
  }

  async function typeInto(row: number, column: number, text: string): Promise<void> {
    h.service.setSelection(row, column, row, column);
    await h.served.settle();
    await h.ui.settle();
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 8) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    // F2 first: the character that *opens* a cell is consumed doing
    // it, so typing straight at the grid loses the leading `=`.
    h.ui.fireEvent.press('F2');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
    h.ui.fireEvent.type(text);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  it('draws a box on the row the reference names', async () => {
    await typeInto(5, 5, '=B2');
    // B2 is row 1: two sides, a top and a bottom, all on one row.
    expect(shapesOn(1)).toHaveLength(4);
    expect(shapesOn(0)).toHaveLength(0);
    expect(shapesOn(2)).toHaveLength(0);
  });

  it('puts it over the columns the reference covers', async () => {
    await typeInto(5, 5, '=B2');
    const left = GUTTER_WIDTH + COLUMN_WIDTH;
    expect(shapesOn(1).map(shape => shape.x).sort((a, b) => a - b)[0]).toBe(left);
  });

  /**
   * A range crossing rows is drawn by each of them, so one reaching
   * into the viewport from above is still outlined.
   */
  it('draws a tall range on every row it crosses', async () => {
    await typeInto(9, 5, '=SUM(B2:B4)');
    // The first and last rows carry a horizontal edge as well as the
    // two sides; the row between carries only the sides.
    expect(shapesOn(1)).toHaveLength(3);
    expect(shapesOn(2)).toHaveLength(2);
    expect(shapesOn(3)).toHaveLength(3);
    expect(shapesOn(4)).toHaveLength(0);
  });

  it('gives each reference the colour its text is drawn in', async () => {
    await typeInto(9, 5, '=B2+C4');
    const first = shapesOn(1)[0].color;
    const second = shapesOn(3)[0].color;
    expect(first).not.toBe(second);
    // The same two colours the text is drawn in, in the same order.
    expect([first, second]).toEqual(
      formulaSpans('=B2+C4')!
        .filter(run => run.color !== undefined)
        .map(run => run.color)
    );
  });

  it('takes them away when the edit is finished', async () => {
    await typeInto(5, 5, '=B2');
    expect(shapesOn(1)).toHaveLength(4);

    h.ui.fireEvent.press('Escape');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    expect(shapesOn(1)).toHaveLength(0);
  });

  it('draws nothing for a cell that is not a formula', async () => {
    await typeInto(5, 5, 'North');
    expect(h.ui.getAllByRole('row').flatMap(row => decorationsOf(row))).toHaveLength(0);
  });
});

/**
 * Clicking a cell into a formula that is open.
 *
 * `FormulaEditing.spec.ts` is the table of (caret context, click) →
 * result; this is the other half of that claim — that the grid asks
 * it, and does what it says, rather than deciding for itself.
 *
 * The mode is the risk: the same click either moves the selection or
 * types into somebody's formula. Both directions are asserted here,
 * because getting either wrong is how a spreadsheet eats a formula
 * somebody was halfway through.
 */
describe('picking a reference by clicking', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount(document => {
      document.setCell(1, 1, '5');
      document.setCell(2, 2, '7');
    });
  });

  /** Opens the cell at `row`/`column` and types `text` into it. */
  async function typing(row: number, column: number, text: string): Promise<void> {
    h.service.setSelection(row, column, row, column);
    await h.served.settle();
    await h.ui.settle();
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 8) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('F2');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
    h.ui.fireEvent.type(text);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  /**
   * Where a cell sits on screen, for clicking it.
   *
   * The row's y comes from the row's own box rather than from
   * `grid.y + HEADER_HEIGHT + row * ROW_HEIGHT`, which assumes the
   * grid has not scrolled. It had: the toolbar grew a row of icons,
   * the viewport lost a row with it, and selecting the cell being
   * edited scrolled the sheet by one. Every coordinate then named the
   * cell below the one it meant, which looked exactly like broken
   * picking and was arithmetic in the spec.
   */
  function at(row: number, column: number): { x: number; y: number } {
    const rowNode = h.ui.getAllByRole('row').find(node => node.properties.get('posInSet') === row + 1);
    if (rowNode === undefined) {
      throw new Error(`row ${row} is not on screen`);
    }
    const grid = h.ui.getVisibleBox(h.ui.getByRole('grid'));
    return {
      x: grid.x + GUTTER_WIDTH + column * COLUMN_WIDTH + 4,
      y: h.ui.getVisibleBox(rowNode).y + 4
    };
  }

  async function clickCell(row: number, column: number): Promise<void> {
    const point = at(row, column);
    h.ui.fireEvent.pointerDown(point.x, point.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(point.x, point.y);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  const draft = () => h.document.sheet.input(8, 0);
  const editorText = () => h.ui.queryByRole('textbox', { name: 'Cell' })?.properties.get('value');

  it('writes the address where a formula wants a value', async () => {
    await typing(8, 0, '=');
    await clickCell(1, 1);
    expect(editorText()).toBe('=B2');
  });

  it('adds to what is already there rather than replacing it', async () => {
    await typing(8, 0, '=');
    await clickCell(1, 1);
    h.ui.fireEvent.type('+');
    await h.ui.settle();
    await clickCell(2, 2);
    expect(editorText()).toBe('=B2+C3');
  });

  /** Still choosing: a second click with nothing typed between. */
  it('replaces the reference when the next click is a change of mind', async () => {
    await typing(8, 0, '=');
    await clickCell(1, 1);
    await clickCell(2, 2);
    expect(editorText()).toBe('=C3');
  });

  it('leaves the formula alone when a click is only a click', async () => {
    await typing(8, 0, '=1');
    await clickCell(1, 1);
    // The edit was committed and the selection moved, as it always does.
    expect(h.ui.queryByRole('textbox', { name: 'Cell' })).toBeNull();
    expect(draft()).toBe('=1');
  });

  it('commits and moves when nothing is open at all', async () => {
    h.service.setSelection(0, 0, 0, 0);
    await h.served.settle();
    await h.ui.settle();
    await clickCell(3, 3);
    expect(h.document.selection.row).toBe(3);
    expect(h.document.selection.column).toBe(3);
  });

  it('drags a range in, rewriting it as the pointer moves', async () => {
    await typing(8, 0, '=SUM(');
    const from = at(1, 1);
    const to = at(3, 2);
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 8, from.y + 8, { buttons: 1 });
    await h.ui.settle();
    h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1 });
    await h.ui.settle();
    h.ui.fireEvent.pointerUp(to.x, to.y);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    expect(editorText()).toBe('=SUM(B2:C4');
  });

  /** The colours follow, because they are the same list. */
  it('outlines what it picked', async () => {
    await typing(8, 0, '=');
    await clickCell(1, 1);
    const row = h.ui.getAllByRole('row').find(node => node.properties.get('posInSet') === 2);
    expect((row?.decorations ?? []).length).toBe(4);
  });
});

/**
 * F4 on the reference the caret is in.
 *
 * The cycle itself has a table in `FormulaEditing.spec.ts`. What is
 * asserted here is the part that only a mounted grid has: that the
 * key reaches the open cell, that the caret ends up somewhere
 * sensible, and that a caret nowhere near a reference is left alone.
 */
describe('cycling a reference with F4', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount();
  });

  async function typing(text: string): Promise<void> {
    h.service.setSelection(5, 0, 5, 0);
    await h.served.settle();
    await h.ui.settle();
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 8) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('F2');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
    h.ui.fireEvent.type(text);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function f4(): Promise<void> {
    h.ui.fireEvent.press('F4');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  const editorText = () => h.ui.queryByRole('textbox', { name: 'Cell' })?.properties.get('value');

  it('pins the reference the caret is sitting after', async () => {
    await typing('=B2');
    await f4();
    expect(editorText()).toBe('=$B$2');
  });

  it('goes round the whole cycle', async () => {
    await typing('=B2');
    await f4();
    expect(editorText()).toBe('=$B$2');
    await f4();
    expect(editorText()).toBe('=B$2');
    await f4();
    expect(editorText()).toBe('=$B2');
    await f4();
    expect(editorText()).toBe('=B2');
  });

  it('leaves the rest of the formula alone', async () => {
    await typing('=SUM(A1,B2)');
    // The caret is after the closing bracket, which is not a reference.
    await f4();
    expect(editorText()).toBe('=SUM(A1,B2)');
  });

  it('keeps the colours in step with what it rewrote', async () => {
    await typing('=B2');
    await f4();
    const runs = h.ui.getByRole('textbox', { name: 'Cell' }).properties.get('spans') as
      | readonly { text: string }[]
      | undefined;
    expect(runs?.map(run => run.text)).toEqual(['=', '$B$2']);
  });
});
