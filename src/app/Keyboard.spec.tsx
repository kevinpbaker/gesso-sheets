import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent, serve } from 'gesso-framework';
import { renderTest, serveForTest, textProperty, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { Sheet } from './SheetContract';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { SheetService } from './SheetService';

/**
 * The keyboard — the exit criterion for Phase 4.
 *
 * Navigate, type, commit and undo, entirely through the semantics
 * tree, with no synthetic mouse anywhere in the file. Nothing here
 * calls a handler, sets a value or focuses a node by hand: every
 * assertion is reached by pressing the key a person would press and
 * read back from what a screen reader would hear or from what the
 * application worker ended up holding.
 *
 * That is the shape of `gesso-components`' own `Keyboard.spec.ts`, and
 * the same reasoning: a spec that poked the editor directly would pass
 * while the sheet was unusable without a mouse.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
}

async function mount(fill?: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
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
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 300 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  // Tab from nothing, as somebody arriving at the page does. The grid
  // is the first stop after the formula bar.
  ui.fireEvent.tab();
  await ui.settle();
  return { ui, served, document };
}

describe('the sheet from the keyboard', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  /** The address the formula bar shows, which is where the selection is. */
  function address(): string | undefined {
    return textProperty(h.ui.getByRole('status'));
  }

  async function press(key: string, modifiers: Record<string, boolean> = {}): Promise<void> {
    h.ui.fireEvent.press(key, modifiers);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function type(text: string): Promise<void> {
    h.ui.fireEvent.type(text);
    await h.ui.settle();
  }

  describe('navigating', () => {
    beforeEach(async () => {
      h = await mount();
      // Focus reaches the grid, not something behind it.
      await press('ArrowDown');
      await press('ArrowUp');
    });

    it('starts on A1', () => {
      expect(address()).toBe('A1');
    });

    it('moves with the arrows', async () => {
      await press('ArrowDown');
      expect(address()).toBe('A2');
      await press('ArrowRight');
      expect(address()).toBe('B2');
      await press('ArrowUp');
      expect(address()).toBe('B1');
      await press('ArrowLeft');
      expect(address()).toBe('A1');
    });

    it('does not walk off the top or the left', async () => {
      await press('ArrowUp');
      await press('ArrowLeft');
      expect(address()).toBe('A1');
    });

    it('moves down on Enter and right on Tab', async () => {
      await press('Enter');
      expect(address()).toBe('A2');
      await press('Tab');
      expect(address()).toBe('B2');
      await press('Tab', { shift: true });
      expect(address()).toBe('A2');
    });

    it('jumps to the ends of a row and of the sheet', async () => {
      await press('End');
      expect(address()).toBe('T1');
      await press('Home');
      expect(address()).toBe('A1');
      await press('End', { control: true });
      expect(address()).toBe('T200');
      await press('Home', { control: true });
      expect(address()).toBe('A1');
    });

    /**
     * The cell scrolled to is one the window had not mounted, which is
     * the case a grid of a hundred rows would never exercise: the
     * offsets are arithmetic precisely so that a cell can be brought
     * into view before it exists as a node.
     */
    it('brings a selection outside the window into view', async () => {
      await press('PageDown');
      await press('PageDown');
      expect(address()).toBe('A49');
      expect(h.ui.getByRole('rowheader', { name: '49' })).toBeDefined();
      expect(h.ui.queryByRole('rowheader', { name: '1' })).toBeNull();
    });
  });

  describe('typing into a cell', () => {
    beforeEach(async () => {
      h = await mount();
      await press('ArrowDown');
      await press('ArrowUp');
    });

    /**
     * The character typed is the first character of the value, not a
     * keystroke spent opening the cell.
     */
    it('replaces the cell with what was typed, starting from the first key', async () => {
      // One press, not a press and an insertion. The grid has focus
      // and no editable does, so the character arrives as a key and
      // nothing else; the handler takes it and calls preventDefault,
      // which is also what stops the browser delivering the same
      // character again to the editor it just opened.
      await press('5');
      await press('Enter');

      expect(h.document.sheet.input(0, 0)).toBe('5');
      expect(h.document.sheet.value(0, 0)).toBe(5);
    });

    it('commits on Enter and lands on the cell below', async () => {
      await press('1');
      await press('Enter');
      expect(address()).toBe('A2');

      await press('2');
      await press('Enter');

      expect(h.document.sheet.value(0, 0)).toBe(1);
      expect(h.document.sheet.value(1, 0)).toBe(2);
      expect(address()).toBe('A3');
    });

    it('commits on Tab and lands on the cell to the right', async () => {
      await press('7');
      await press('Tab');

      expect(h.document.sheet.value(0, 0)).toBe(7);
      expect(address()).toBe('B1');
    });

    it('computes a formula that was typed', async () => {
      await press('=');
      await type('1+2*3');
      await press('Enter');

      expect(h.document.sheet.value(0, 0)).toBe(7);
      expect(h.document.sheet.input(0, 0)).toBe('=1+2*3');
    });
  });

  describe('opening a cell that already has something in it', () => {
    beforeEach(async () => {
      h = await mount(document => document.setCell(0, 0, '=1+2'));
      await press('ArrowDown');
      await press('ArrowUp');
    });

    it('shows the formula rather than the value while it is open', async () => {
      // The cell displays 3 and was typed `=1+2`.
      expect(h.ui.getByRole('cell', { name: '3' })).toBeDefined();

      await press('F2');

      expect(h.ui.getByRole('textbox', { name: 'Cell' })).toHaveText('=1+2');
    });

    it('keeps the edit when it is committed', async () => {
      // F2 opens it with `=1+2` and the caret at the end, so this is
      // appended rather than retyped.
      await press('F2');
      await type('+10');
      await press('Enter');

      expect(h.document.sheet.value(0, 0)).toBe(13);
    });

    /** The whole reason a draft is held on this side of the barrier. */
    it('puts back what was there on Escape', async () => {
      await press('F2');
      await type('=999');
      await press('Escape');

      expect(h.document.sheet.input(0, 0)).toBe('=1+2');
      expect(h.document.sheet.value(0, 0)).toBe(3);
      expect(h.ui.getByRole('cell', { name: '3' })).toBeDefined();
    });

    it('empties the cell on Delete without opening it', async () => {
      await press('Delete');
      expect(h.document.sheet.input(0, 0)).toBe('');
    });
  });

  describe('the formula bar', () => {
    beforeEach(async () => {
      h = await mount(document => document.setCell(0, 0, '=2*21'));
      await press('ArrowDown');
      await press('ArrowUp');
    });

    it('shows what the selected cell was typed as', () => {
      expect(h.ui.getByRole('textbox', { name: 'Formula' })).toHaveText('=2*21');
    });

    /**
     * The same buffer, not a copy of it. Two buffers kept in step
     * would be two answers to what Escape puts back, and the bar and
     * the cell would disagree for as long as it took a keystroke to
     * cross between them.
     */
    it('shows the draft while a cell is open, character by character', async () => {
      await press('F2');
      await type('+1');

      expect(h.ui.getByRole('textbox', { name: 'Formula' })).toHaveText('=2*21+1');
      expect(h.ui.getByRole('textbox', { name: 'Cell' })).toHaveText('=2*21+1');
    });

    it('goes back to the committed text when the edit is abandoned', async () => {
      await press('F2');
      await type('+0');
      await press('Escape');

      expect(h.ui.getByRole('textbox', { name: 'Formula' })).toHaveText('=2*21');
    });
  });

  describe('undo', () => {
    beforeEach(async () => {
      h = await mount(document => document.setCell(0, 0, 'before'));
      await press('ArrowDown');
      await press('ArrowUp');
    });

    it('takes back a committed edit and puts it forward again', async () => {
      await press('a');
      await type('fter');
      await press('Enter');
      expect(h.document.sheet.input(0, 0)).toBe('after');

      await press('z', { control: true });
      expect(h.document.sheet.input(0, 0)).toBe('before');

      await press('z', { control: true, shift: true });
      expect(h.document.sheet.input(0, 0)).toBe('after');
    });

    it('recalculates what the undone cell fed', async () => {
      h.document.setCell(0, 1, '=A1&"!"');
      h.document.sheet.recalculate();

      await press('x');
      await press('Enter');
      expect(h.document.sheet.value(0, 1)).toBe('x!');

      await press('z', { control: true });
      expect(h.document.sheet.value(0, 1)).toBe('before!');
    });
  });

  /**
   * A cell that is open and then scrolled out of the window.
   *
   * Nothing had exercised `EditableTextModel` inside a virtualised row
   * before this phase, and the failure mode is quiet: the node is
   * unmounted, and if the draft lived in it the edit would be gone
   * with no error anywhere.
   */
  describe('an open cell that scrolls away', () => {
    beforeEach(async () => {
      h = await mount();
      await press('ArrowDown');
      await press('ArrowUp');
    });

    it('keeps the draft, and commits it when it comes back', async () => {
      await press('=');
      await type('40+2');

      h.ui.fireEvent.wheel({ x: 300, y: 200, deltaY: 4000 });
      await h.ui.settle();
      expect(h.ui.queryByRole('textbox', { name: 'Cell' })).toBeNull();

      // The bar still has it, because the buffer is not in the cell.
      expect(h.ui.getByRole('textbox', { name: 'Formula' })).toHaveText('=40+2');

      await press('Enter');
      expect(h.document.sheet.value(0, 0)).toBe(42);
    });
  });
});
