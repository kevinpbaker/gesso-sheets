import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, textProperty, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiKeyModifiers } from 'gesso-core';

import { Sheet } from './SheetContract';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
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
    sheetChannel(service)
  ]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 700, height: 300 });
  await ui.settle();
  await served.settle();
  await ui.settle();

  // Tab from nothing, as somebody arriving at the page does, until the
  // sheet itself has the keyboard.
  //
  // This used to be a single Tab, which lands on the formula bar, and
  // the specs below passed anyway — because every key was going to the
  // bar and the bar was driving the grid. That was the bug a person
  // hit as soon as they tried to edit *in* the bar: Backspace emptied
  // the whole cell and the next character arrived twice. Asserting
  // which control the keyboard reached is what stops a spec proving
  // the sheet works through a path nobody uses.
  const grid = ui.getByRole('grid');
  let stops = 0;
  while (ui.runtime.input.focus.focusedNode !== grid) {
    if (stops++ > 8) {
      throw new Error(`Tab never reached the grid; it stopped on ${ui.debug()}`);
    }
    ui.fireEvent.tab();
    await ui.settle();
  }
  return { ui, served, document };
}

describe('the sheet from the keyboard', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  /**
   * The address the name box shows, which is where the selection is.
   *
   * It was `role: status` until Phase 8 made the address a field you
   * can type into. Still read through the semantics tree, which is
   * the point of this file.
   */
  function address(): string | undefined {
    return textProperty(h.ui.getByRole('textbox', { name: 'Name box' }));
  }

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
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
      await press('End', { ctrl: true });
      expect(address()).toBe('T200');
      await press('Home', { ctrl: true });
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
    });

    it('takes back a committed edit and puts it forward again', async () => {
      await press('a');
      await type('fter');
      await press('Enter');
      expect(h.document.sheet.input(0, 0)).toBe('after');

      await press('z', { ctrl: true });
      expect(h.document.sheet.input(0, 0)).toBe('before');

      await press('z', { ctrl: true, shift: true });
      expect(h.document.sheet.input(0, 0)).toBe('after');
    });

    it('recalculates what the undone cell fed', async () => {
      h.document.setCell(0, 1, '=A1&"!"');
      h.document.sheet.recalculate();

      await press('x');
      await press('Enter');
      expect(h.document.sheet.value(0, 1)).toBe('x!');

      await press('z', { ctrl: true });
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

  /**
   * The keys Excel users already have — Phase 19.
   *
   * Each of these is a key a hand presses before it has decided to,
   * and before this phase each one did nothing or did something else.
   */
  describe('the keys a hand that knows Excel reaches for', () => {
    const selection = () => h.document.selection;

    describe('Ctrl+Arrow', () => {
      beforeEach(async () => {
        h = await mount(d => {
          // A1:A5 filled, a gap, then A10:A12; C1 alone across a gap.
          for (const r of [0, 1, 2, 3, 4, 9, 10, 11]) {
            d.setCell(r, 0, String(r + 1));
          }
          d.setCell(0, 2, 'far');
        });
      });

      it('goes to the end of the run it is in', async () => {
        await press('ArrowDown', { ctrl: true });
        expect(address()).toBe('A5');
      });

      it('then across the gap to the start of the next run', async () => {
        await press('ArrowDown', { ctrl: true });
        await press('ArrowDown', { ctrl: true });
        expect(address()).toBe('A10');
        await press('ArrowDown', { ctrl: true });
        expect(address()).toBe('A12');
      });

      it('goes to the edge of the sheet past the last of the data', async () => {
        for (let i = 0; i < 4; i++) {
          await press('ArrowDown', { ctrl: true });
        }
        expect(address()).toBe('A200');
        await press('ArrowUp', { ctrl: true });
        expect(address()).toBe('A12');
      });

      it('goes sideways over empty cells to the next filled one', async () => {
        await press('ArrowRight', { ctrl: true });
        expect(address()).toBe('C1');
        await press('ArrowRight', { ctrl: true });
        expect(address()).toBe('T1');
        await press('ArrowLeft', { ctrl: true });
        expect(address()).toBe('C1');
      });

      it('selects as far as it goes with Shift', async () => {
        await press('ArrowDown', { ctrl: true, shift: true });
        expect(selection()).toEqual({ row: 4, column: 0, anchorRow: 0, anchorColumn: 0 });
      });

      it('counts a cell an array spilled into as filled', async () => {
        h.ui.unmount();
        h.served.dispose();
        h = await mount(d => d.setCell(0, 1, '=SEQUENCE(4)'));
        await press('ArrowRight');
        await press('ArrowDown', { ctrl: true });
        expect(address()).toBe('B4');
      });
    });

    describe('selecting a line', () => {
      beforeEach(async () => {
        h = await mount();
        await press('ArrowRight');
        await press('ArrowDown');
      });

      it('takes the column on Ctrl+Space', async () => {
        await press(' ', { ctrl: true });
        expect(selection()).toEqual({ row: 0, column: 1, anchorRow: 199, anchorColumn: 1 });
      });

      it('takes the row on Shift+Space', async () => {
        await press(' ', { shift: true });
        expect(selection()).toEqual({ row: 1, column: 0, anchorRow: 1, anchorColumn: 19 });
      });
    });

    describe('Tab, Tab, Enter', () => {
      beforeEach(async () => {
        h = await mount();
      });

      it('comes back to the column the Tabs started from', async () => {
        await press('ArrowRight');
        await press('a');
        await press('Tab');
        await press('b');
        await press('Tab');
        await press('c');
        await press('Enter');

        expect(address()).toBe('B2');
        expect([0, 1, 2].map(c => h.document.sheet.input(0, c + 1))).toEqual(['a', 'b', 'c']);
      });

      it('goes straight down when something other than Tab moved it', async () => {
        await press('Tab');
        await press('ArrowRight');
        await press('Enter');
        expect(address()).toBe('C2');
      });
    });

    describe('Ctrl+Enter', () => {
      it('puts what was typed in every selected cell, its references moved', async () => {
        h = await mount(d => {
          for (let r = 0; r < 3; r++) {
            d.setCell(r, 0, String((r + 1) * 10));
          }
        });
        await press('ArrowRight');
        await press('ArrowDown', { shift: true });
        await press('ArrowDown', { shift: true });
        // Typed in the cell the cursor is on, which is B3: the corner
        // Shift moved, and the one the editor opens in.
        await press('=');
        await type('A3*2');
        await press('Enter', { ctrl: true });

        expect([0, 1, 2].map(r => h.document.sheet.input(r, 1))).toEqual(['=A1*2', '=A2*2', '=A3*2']);
        expect(h.document.sheet.value(2, 1)).toBe(60);
        // The selection is still the three cells, and it is one undo.
        expect(selection()).toEqual({ row: 2, column: 1, anchorRow: 0, anchorColumn: 1 });
        await press('z', { ctrl: true });
        expect([0, 1, 2].map(r => h.document.sheet.input(r, 1))).toEqual(['', '', '']);
      });
    });

    describe('Alt+Enter', () => {
      it('puts a line break in the cell instead of committing it', async () => {
        h = await mount();
        await press('a');
        await press('Enter', { alt: true });
        await type('b');
        expect(h.ui.getByRole('textbox', { name: 'Cell' })).toBeDefined();
        await press('Enter');

        expect(h.document.sheet.input(0, 0)).toBe('a\nb');
        expect(address()).toBe('A2');
      });

      it('turns wrap on, so the second line is seen, and takes both back on one undo', async () => {
        h = await mount();
        await press('a');
        await press('Enter', { alt: true });
        await type('b');
        await press('Enter');
        expect(h.document.formatAt(0, 0).paint.wrap).toBe(true);

        await press('z', { ctrl: true });
        expect(h.document.sheet.input(0, 0)).toBe('');
        expect(h.document.formatAt(0, 0).paint.wrap).toBe(false);
      });
    });

    describe('Ctrl+;', () => {
      beforeEach(async () => {
        h = await mount();
      });

      it("types today's date, which the sheet reads as a date", async () => {
        await press(';', { ctrl: true });
        await press('Enter');

        const now = new Date();
        const typed = h.document.sheet.input(0, 0);
        expect(typed).toMatch(/^\d{4}-\d{2}-\d{2}$/);
        expect(typed.startsWith(String(now.getFullYear()))).toBe(true);
        expect(typeof h.document.sheet.value(0, 0)).toBe('number');
      });

      it('types the time with Shift, however the keyboard spells it', async () => {
        await press(':', { ctrl: true, shift: true });
        await press('Enter');
        expect(h.document.sheet.input(0, 0)).toMatch(/^\d{2}:\d{2}$/);
        expect(typeof h.document.sheet.value(0, 0)).toBe('number');
      });

      it('goes in at the caret of a cell already open', async () => {
        await press('D');
        await type('ue ');
        await press(';', { ctrl: true });
        await press('Enter');
        expect(h.document.sheet.input(0, 0)).toMatch(/^Due \d{4}-\d{2}-\d{2}$/);
      });
    });
  });
});
