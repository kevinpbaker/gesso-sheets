import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent, ShellService } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiKeyModifiers } from 'gesso-core';

import { COLUMN_WIDTH, GUTTER_WIDTH, HEADER_HEIGHT, ROW_HEIGHT } from './dimensions';
import { sheetChannel } from './sheetChannel';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { SheetService } from './SheetService';

/**
 * Clipboard and fill — the exit criterion for Phase 5.
 *
 * Two claims, and the second is the reason the first is not enough: a
 * round trip inside the sheet, *and* a block copied out of a real
 * spreadsheet landing correctly. A format that only talks to itself
 * would pass the round trip and be worthless.
 *
 * Every paste here goes through the engine's `Paste` event — the gap
 * this file's roadmap has carried since the first page — so what is
 * tested is the path a ctrl-V actually takes, not a command called by
 * hand.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
  copied: () => string;
}

async function mount(fill?: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);

  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 300 });

  // What the shell is asked to put on the system clipboard. A render
  // worker has no clipboard; `ShellService` is how it asks, and this
  // is the other end of the request. Installed after the runtime has
  // built, because the runtime installs its own handler as it starts
  // and setting one first is setting one that is about to be replaced.
  const writes: string[] = [];
  ui.runtime.services.get(ShellService).setHandler(request => {
    if (request.type === 'clipboard') {
      writes.push(request.text);
    }
  });
  await ui.settle();
  await served.settle();
  await ui.settle();

  const grid = ui.getByRole('grid');
  let stops = 0;
  while (ui.runtime.input.focus.focusedNode !== grid) {
    if (stops++ > 8) {
      throw new Error('Tab never reached the grid');
    }
    ui.fireEvent.tab();
    await ui.settle();
  }
  return { ui, served, document, copied: () => writes[writes.length - 1] ?? '' };
}

describe('clipboard and fill', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
    h.ui.fireEvent.press(key, modifiers);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  /** A real ctrl-V: text from the system, with no caret anywhere. */
  async function paste(text: string): Promise<void> {
    h.ui.fireEvent.paste(text);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  describe('selecting a range', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(0, 0, 'a');
        document.setCell(0, 1, 'b');
        document.setCell(1, 0, 'c');
        document.setCell(1, 1, 'd');
      });
    });

    it('extends with shift and an arrow', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });

      expect(h.document.selection).toEqual({ row: 1, column: 1, anchorRow: 0, anchorColumn: 0 });
    });

    it('copies the rectangle as tab-separated rows', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });
      await press('c', { ctrl: true });

      expect(h.copied()).toBe('a\tb\nc\td');
    });
  });

  /**
   * Sweeping a selection out with the pointer, which is how most
   * people make one. Shift+click and shift+arrows were built first and
   * this was simply missing — the cells had a click and nothing else,
   * so a drag across them selected the one it started on.
   */
  describe('dragging a selection', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(1, 1, 'a');
        document.setCell(3, 3, 'b');
        document.setCell(6, 6, 'far');
      });
    });

    async function sweep(from: { row: number; column: number }, to: { row: number; column: number }): Promise<void> {
      const grid = h.ui.getVisibleBox(h.ui.getByRole('grid'));
      const at = (cell: { row: number; column: number }) => ({
        x: grid.x + GUTTER_WIDTH + cell.column * COLUMN_WIDTH + 4,
        y: grid.y + HEADER_HEIGHT + cell.row * ROW_HEIGHT + 4
      });
      const start = at(from);
      const end = at(to);
      h.ui.fireEvent.pointerDown(start.x, start.y, { buttons: 1 });
      h.ui.fireEvent.pointerMove(start.x + 8, start.y + 8, { buttons: 1 });
      h.ui.fireEvent.pointerMove(end.x, end.y, { buttons: 1 });
      h.ui.fireEvent.pointerUp(end.x, end.y);
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    }

    it('selects the rectangle the pointer swept', async () => {
      await sweep({ row: 1, column: 1 }, { row: 3, column: 3 });

      expect(h.document.selection).toMatchObject({ anchorRow: 1, anchorColumn: 1, row: 3, column: 3 });
    });

    it('sweeps backwards as well as forwards', async () => {
      await sweep({ row: 3, column: 3 }, { row: 1, column: 1 });

      expect(h.document.selection).toMatchObject({ anchorRow: 3, anchorColumn: 3, row: 1, column: 1 });
    });

    it('copies what was swept', async () => {
      await sweep({ row: 1, column: 1 }, { row: 2, column: 2 });
      await press('c', { ctrl: true });

      expect(h.copied()).toBe('a\t\n\t');
    });

    /**
     * A click after a sweep. The sweep leaves a rectangle behind, and
     * a click is how a person puts it down again.
     */
    it('collapses to one cell when a cell is clicked afterwards', async () => {
      await sweep({ row: 1, column: 1 }, { row: 3, column: 3 });
      expect(h.document.selection).toMatchObject({ anchorRow: 1, row: 3 });

      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: 'b' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(h.document.selection).toEqual({ row: 3, column: 3, anchorRow: 3, anchorColumn: 3 });
    });

    /**
     * A click well away from the sweep, which is the case the spec
     * harness was too tidy to catch: in a browser the anchor stayed
     * where the sweep began, so clicking one cell selected everything
     * between it and there — and a paste afterwards landed at the
     * corner of that rectangle rather than where the click was.
     */
    it('puts the rectangle down, anchor and all', async () => {
      await sweep({ row: 1, column: 1 }, { row: 3, column: 3 });

      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: 'far' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(h.document.selection).toEqual({ row: 6, column: 6, anchorRow: 6, anchorColumn: 6 });
    });

    it('leaves the keyboard on the grid, so the selection can be used', async () => {
      await sweep({ row: 1, column: 1 }, { row: 3, column: 3 });
      await press('Delete');

      expect(h.document.sheet.input(1, 1)).toBe('');
      expect(h.document.sheet.input(3, 3)).toBe('');
    });
  });

  describe('a round trip inside the sheet', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(0, 0, '2');
        document.setCell(0, 1, '3');
        document.setCell(1, 0, '=A1*10');
      });
    });

    it('comes back as it went, values and all', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });
      await press('c', { ctrl: true });
      const text = h.copied();

      // Down to row 4, column A. Home first, because an arrow from a
      // range leaves the selection on the corner it was steering, not
      // on the corner it started from.
      await press('Home');
      await press('ArrowDown');
      await press('ArrowDown');
      await paste(text);

      expect(h.document.sheet.value(3, 0)).toBe(2);
      expect(h.document.sheet.value(3, 1)).toBe(3);
      expect(h.document.sheet.value(4, 0)).toBe(20);
    });

    /**
     * Copying formulas rather than the numbers they showed is what
     * makes this a spreadsheet's copy: the pasted block computes from
     * where it landed, not from where it came from.
     */
    /**
     * The formula moves with the block, so it computes from where it
     * landed. A2 holds `=A1*10` over a 2; copied one column right it
     * becomes `=B1*10` over a 3, and says 30 rather than 20.
     */
    it('moves the formulas with the block', async () => {
      await press('ArrowDown');
      await press('c', { ctrl: true });
      const text = h.copied();
      expect(text).toBe('=A1*10');

      await press('ArrowRight');
      await paste(text);

      expect(h.document.sheet.input(1, 1)).toBe('=(B1*10)');
      expect(h.document.sheet.value(1, 1)).toBe(30);
      // And the original is untouched.
      expect(h.document.sheet.value(1, 0)).toBe(20);
    });

    it('cuts, which copies and then empties', async () => {
      await press('ArrowDown', { shift: true });
      await press('x', { ctrl: true });

      expect(h.copied()).toBe('2\n=A1*10');
      expect(h.document.sheet.input(0, 0)).toBe('');
      expect(h.document.sheet.input(1, 0)).toBe('');
    });

    it('undoes a paste in one press', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });
      await press('c', { ctrl: true });
      const text = h.copied();
      await press('Home');
      await press('ArrowDown');
      await press('ArrowDown');
      await paste(text);
      expect(h.document.sheet.input(3, 0)).toBe('2');
      expect(h.document.sheet.input(3, 1)).toBe('3');

      await press('z', { ctrl: true });

      expect(h.document.sheet.input(3, 0)).toBe('');
      expect(h.document.sheet.input(3, 1)).toBe('');
    });

    it('empties a whole selection on Delete', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });
      await press('Delete');

      expect(h.document.sheet.input(0, 0)).toBe('');
      expect(h.document.sheet.input(1, 1)).toBe('');
    });
  });

  /**
   * The half a format that only talks to itself would fail. These are
   * real clipboard payloads: Excel and Google Sheets both write
   * tab-separated rows, Excel ends with a CRLF, and a cell holding a
   * tab or a newline arrives quoted the way CSV quotes it.
   */
  describe('a paste from a real spreadsheet', () => {
    beforeEach(async () => {
      h = await mount();
    });

    it('lands a block of values where the selection is', async () => {
      await paste('Region\tUnits\nNorth\t120\nSouth\t157\r\n');

      expect(h.document.sheet.value(0, 0)).toBe('Region');
      expect(h.document.sheet.value(1, 1)).toBe(120);
      expect(h.document.sheet.value(2, 0)).toBe('South');
      // The trailing CRLF Excel adds is not a fourth row of blanks.
      expect(h.document.sheet.input(3, 0)).toBe('');
    });

    it('keeps a cell that had a tab or a newline in it', async () => {
      await paste('"a\tb"\tplain\n"two\nlines"\tlast');

      expect(h.document.sheet.value(0, 0)).toBe('a\tb');
      expect(h.document.sheet.value(0, 1)).toBe('plain');
      expect(h.document.sheet.value(1, 0)).toBe('two\nlines');
      expect(h.document.sheet.value(1, 1)).toBe('last');
    });

    it('takes a formula pasted as text and computes it', async () => {
      await paste('5\t6\n=A1+B1\t');
      expect(h.document.sheet.value(1, 0)).toBe(11);
    });

    /**
     * References in text from somewhere else are left exactly as they
     * arrived. Moving them would be inventing an intent: they were
     * never relative to this sheet, and where they were copied from is
     * not something the clipboard says.
     */
    it('does not move references in text it did not copy', async () => {
      await press('ArrowDown');
      await press('ArrowDown');
      await paste('=A1');
      expect(h.document.sheet.input(2, 0)).toBe('=A1');
    });

    it('lands a single cell without needing a rectangle', async () => {
      await paste('hello');
      expect(h.document.sheet.value(0, 0)).toBe('hello');
    });

    it('selects what it pasted', async () => {
      await paste('a\tb\nc\td');
      expect(h.document.selection).toMatchObject({ row: 0, column: 0, anchorRow: 1, anchorColumn: 1 });
    });
  });

  /**
   * The fill handle: the square on the selection's far corner that
   * extends it, repeating what is in it and moving the formulas as it
   * goes. Dragged here, not called — the handle is a node with a
   * gesture on it and that is the thing being tested.
   */
  describe('the fill handle', () => {
    beforeEach(async () => {
      h = await mount(document => {
        for (let row = 0; row < 4; row++) {
          document.setCell(row, 1, String(4 + row));
          document.setCell(row, 2, String(5 + row));
        }
        document.setCell(0, 3, '=B1*C1');
      });
    });

    async function dragHandleTo(rows: number): Promise<void> {
      const handle = h.ui.getByRole('button', { name: 'Fill' });
      const box = h.ui.getVisibleBox(handle);
      const x = box.x + box.width / 2;
      const y = box.y + box.height / 2;
      h.ui.fireEvent.pointerDown(x, y, { buttons: 1 });
      h.ui.fireEvent.pointerMove(x, y + 6, { buttons: 1 });
      h.ui.fireEvent.pointerMove(x, y + rows * ROW_HEIGHT, { buttons: 1 });
      h.ui.fireEvent.pointerUp(x, y + rows * ROW_HEIGHT);
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
    }

    it('is on the selection, and moves with it', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '20' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
      expect(h.ui.getByRole('button', { name: 'Fill' })).toBeDefined();
    });

    it('repeats a formula down, moving its references', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '20' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      await dragHandleTo(3);

      expect(h.document.sheet.input(1, 3)).toBe('=(B2*C2)');
      expect(h.document.sheet.value(1, 3)).toBe(5 * 6);
      expect(h.document.sheet.value(3, 3)).toBe(7 * 8);
    });

    it('undoes the whole fill in one press', async () => {
      h.ui.fireEvent.click(h.ui.getByRole('cell', { name: '20' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();
      await dragHandleTo(3);
      expect(h.document.sheet.input(3, 3)).not.toBe('');

      await press('z', { ctrl: true });

      expect(h.document.sheet.input(1, 3)).toBe('');
      expect(h.document.sheet.input(3, 3)).toBe('');
      // And what was filled from is untouched.
      expect(h.document.sheet.input(0, 3)).toBe('=B1*C1');
    });
  });
});
