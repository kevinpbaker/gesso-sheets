import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, textProperty, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiKeyModifiers } from 'gesso-core';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * The chrome — the exit criterion for Phase 8.
 *
 * The whole top bar driven from the keyboard, through the semantics
 * tree, with no pointer event anywhere in the file. That is Phase 4's
 * standard applied to the part of the application that is usually
 * exempted from it: a grid that can be used without a mouse and a
 * menu bar that cannot is an application that cannot be used without
 * a mouse.
 *
 * Nothing here calls a handler or focuses a node by hand. Every
 * assertion is reached by pressing the key a person would press and
 * read back from what a screen reader would hear, or from what the
 * application worker ended up holding.
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
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served, document };
}

describe('the top bar from the keyboard', () => {
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

  async function type(text: string): Promise<void> {
    h.ui.fireEvent.type(text);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  /** Tab until the keyboard is on the sheet, as somebody arriving does. */
  async function reachTheGrid(): Promise<void> {
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 10) {
        throw new Error(`Tab never reached the grid; it stopped on ${h.ui.debug()}`);
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
  }

  const address = () => textProperty(h.ui.getByRole('textbox', { name: 'Name box' }));
  const focused = () => h.ui.runtime.input.focus.focusedNode;

  describe('the menu bar', () => {
    beforeEach(async () => {
      h = await mount();
      await reachTheGrid();
    });

    it('is a single tab stop, not one per menu', async () => {
      // Whatever the bar grows to, Tab past it costs one press. A bar
      // of seven menus that is seven stops is a bar people Tab
      // through rather than into.
      const bar = h.ui.getByRole('menubar');
      expect(bar).toBeDefined();
      expect(h.ui.getAllByRole('menubar')).toHaveLength(1);
    });

    it('is reached from the sheet with F10', async () => {
      await press('F10');
      expect(focused()).toBe(h.ui.getByRole('menubar'));
    });

    it('opens a menu with ArrowDown and lists its commands', async () => {
      await press('F10');
      await press('ArrowDown');

      expect(h.ui.getByRole('menu', { name: 'Edit' })).toBeDefined();
      expect(h.ui.getByRole('menuitem', { name: 'Undo' })).toBeDefined();
      expect(h.ui.getByRole('menuitem', { name: 'Select all' })).toBeDefined();
    });

    it('opens a menu by its letter', async () => {
      await press('F10');
      await press('d');
      expect(h.ui.getByRole('menu', { name: 'Data' })).toBeDefined();
      expect(h.ui.getByRole('menuitem', { name: 'Fill down' })).toBeDefined();
    });

    /**
     * The case a focus-trapped popup cannot serve, and the reason the
     * bar is not built out of `Menu`: while a menu is showing, left
     * and right walk the *bar*.
     */
    it('walks to the menu next door with a menu still open', async () => {
      await press('F10');
      await press('ArrowDown');
      expect(h.ui.getByRole('menu', { name: 'Edit' })).toBeDefined();

      await press('ArrowRight');
      expect(h.ui.queryByRole('menu', { name: 'Edit' })).toBeNull();
      expect(h.ui.getByRole('menu', { name: 'Data' })).toBeDefined();

      await press('ArrowLeft');
      expect(h.ui.getByRole('menu', { name: 'Edit' })).toBeDefined();
    });

    it('runs the command Enter lands on', async () => {
      // Something to take back, so Undo is worth choosing.
      await press('5');
      await press('Enter');
      expect(h.document.sheet.value(0, 0)).toBe(5);

      await press('F10');
      await press('ArrowDown');
      await press('Enter');

      expect(h.document.sheet.input(0, 0)).toBe('');
    });

    it('shows the accelerator beside the command', async () => {
      await press('F10');
      await press('ArrowDown');
      // Drawn in the row, so the text is under the menu item rather
      // than in its name. A menu that did not print its keys would be
      // a menu people never graduate from.
      expect(h.ui.getByText('Ctrl+Z')).toBeDefined();
      expect(h.ui.getByText('Delete')).toBeDefined();
    });

    it('greys a command that cannot be run', async () => {
      await press('F10');
      await press('ArrowDown');
      // Nothing has been typed, so there is nothing to undo.
      expect(h.ui.getByRole('menuitem', { name: 'Undo', disabled: true })).toBeDefined();
    });

    /**
     * The first Escape closes the menu and leaves the bar focused;
     * only the second hands the keyboard back. Closing straight to
     * the sheet loses the place of anybody who opened the wrong menu.
     */
    it('takes two Escapes to get back to the sheet', async () => {
      await press('F10');
      await press('ArrowDown');
      await press('Escape');
      expect(h.ui.queryByRole('menu', { name: 'Edit' })).toBeNull();
      expect(focused()).toBe(h.ui.getByRole('menubar'));

      await press('Escape');
      expect(focused()).toBe(h.ui.getByRole('grid'));
    });

    it('gives the keyboard back to the sheet after a command', async () => {
      await press('F10');
      await press('d');
      await press('Enter');
      expect(focused()).toBe(h.ui.getByRole('grid'));
    });
  });

  describe('the name box', () => {
    beforeEach(async () => {
      h = await mount();
      await reachTheGrid();
    });

    it('says where the selection is', async () => {
      expect(address()).toBe('A1');
      await press('ArrowDown');
      expect(address()).toBe('A2');
    });

    it('says what a range is, not just its corner', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowRight', { shift: true });
      expect(address()).toBe('A1:B2');
    });

    /** The one navigation a keyboard cannot otherwise do. */
    it('jumps to an address typed into it', async () => {
      await press('g', { ctrl: true });
      expect(focused()).toBe(h.ui.getByRole('textbox', { name: 'Name box' }));

      await type('C120');
      await press('Enter');

      expect(address()).toBe('C120');
      expect(focused()).toBe(h.ui.getByRole('grid'));
    });

    it('selects a range typed into it, with its first cell active', async () => {
      await press('g', { ctrl: true });
      await type('B2:D5');
      await press('Enter');
      expect(address()).toBe('B2:D5');
      // The active cell is the near corner, which is where somebody
      // who asked for B2:D5 expects to start typing.
      expect(h.document.selection.row).toBe(1);
      expect(h.document.selection.column).toBe(1);
    });

    /**
     * `B7+1` is a formula typed into the wrong box. Jumping to B7
     * would be reading half of what somebody wrote and acting on it.
     */
    it('stays put when what was typed is not an address', async () => {
      await press('g', { ctrl: true });
      await type('not a cell');
      await press('Enter');
      expect(address()).toBe('A1');
    });

    it('puts the address back on Escape', async () => {
      await press('g', { ctrl: true });
      await type('Z9');
      await press('Escape');
      expect(address()).toBe('A1');
      expect(focused()).toBe(h.ui.getByRole('grid'));
    });
  });

  describe('find and replace', () => {
    const book = (d: SheetDocument) => {
      d.setCell(0, 0, 'apple');
      d.setCell(4, 2, 'pineapple');
      d.setCell(9, 1, 'pear');
    };

    beforeEach(async () => {
      h = await mount(book);
      await reachTheGrid();
    });

    it('opens on Ctrl+F with the keyboard in the field', async () => {
      await press('f', { ctrl: true });
      expect(focused()).toBe(h.ui.getByRole('searchbox', { name: 'Find' }));
    });

    it('moves the selection to the match', async () => {
      await press('f', { ctrl: true });
      await type('pine');
      await press('Enter');
      expect(address()).toBe('C5');
    });

    it('steps through the matches on Enter and wraps', async () => {
      await press('f', { ctrl: true });
      await type('apple');
      await press('Enter');
      expect(address()).toBe('A1');
      await press('Enter');
      expect(address()).toBe('C5');
      await press('Enter');
      expect(address()).toBe('A1');
    });

    it('says how it is going', async () => {
      await press('f', { ctrl: true });
      await type('apple');
      await press('Enter');
      expect(h.ui.getByText('1/2')).toBeDefined();
    });

    it('says so when nothing matches', async () => {
      await press('f', { ctrl: true });
      await type('quince');
      await press('Enter');
      expect(h.ui.getByText('None')).toBeDefined();
    });

    it('closes on Escape and gives the keyboard back', async () => {
      await press('f', { ctrl: true });
      await press('Escape');
      expect(h.ui.queryByRole('searchbox', { name: 'Find' })).toBeNull();
      expect(focused()).toBe(h.ui.getByRole('grid'));
    });

    it('shows the replace half only when replace was asked for', async () => {
      await press('f', { ctrl: true });
      expect(h.ui.queryByRole('searchbox', { name: 'Replace with' })).toBeNull();
      await press('Escape');

      await press('h', { ctrl: true });
      expect(h.ui.getByRole('searchbox', { name: 'Replace with' })).toBeDefined();
    });
  });

  describe('the status bar', () => {
    beforeEach(async () => {
      h = await mount(d => {
        d.setCell(0, 0, '10');
        d.setCell(1, 0, '20');
        d.setCell(2, 0, 'apples');
      });
      await reachTheGrid();
    });

    it('says nothing about a single empty cell', async () => {
      await press('ArrowRight');
      expect(h.ui.queryByText(/Sum/)).toBeNull();
      expect(h.ui.queryByText(/Count/)).toBeNull();
    });

    it('adds up the selection', async () => {
      await press('ArrowDown', { shift: true });
      expect(h.ui.getByText(/Sum 30/)).toBeDefined();
      expect(h.ui.getByText(/Average 15/)).toBeDefined();
    });

    /**
     * `Sum 0` over a column of names is a true statement about the
     * empty set that reads as a false one about the names.
     */
    it('counts text without adding it', async () => {
      await press('ArrowDown', { shift: true });
      await press('ArrowDown', { shift: true });
      expect(h.ui.getByText(/Count 3/)).toBeDefined();
      expect(h.ui.getByText(/Sum 30/)).toBeDefined();
    });
  });

  describe('the accelerators', () => {
    beforeEach(async () => {
      h = await mount(d => {
        d.setCell(0, 0, '4');
        d.setCell(0, 1, '=A1*3');
      });
      await reachTheGrid();
    });

    it('fills down on Ctrl+D', async () => {
      await press('ArrowRight');
      await press('ArrowDown', { shift: true });
      await press('d', { ctrl: true });
      expect(h.document.sheet.input(1, 1)).toBe('=(A2*3)');
    });

    /**
     * Escape has to be routed to the chrome rather than left to
     * bubble: a dialog lives in the overlay layer, the grid is not
     * its child, and `Dialog` traps a keyboard it never takes. Found
     * by pressing Escape in a browser — every spec passed without it,
     * because asserting a dialog is open never asks what has the
     * keyboard.
     */
    it('closes the shortcut sheet on Escape', async () => {
      await press('/', { ctrl: true });
      expect(h.ui.queryByRole('dialog')).not.toBeNull();
      await press('Escape');
      expect(h.ui.queryByRole('dialog')).toBeNull();
    });

    it('closes the find bar on Escape from the sheet', async () => {
      await press('f', { ctrl: true });
      // Back to the sheet first, so the key is arriving where it does
      // when somebody has gone back to typing and changed their mind.
      await press('Escape');
      expect(h.ui.queryByRole('searchbox', { name: 'Find' })).toBeNull();
    });

    /** With a cell open, Escape is the cell's: it puts back what was there. */
    it('leaves Escape to an open cell', async () => {
      await press('9');
      await press('Escape');
      expect(h.document.sheet.value(0, 0)).toBe(4);
    });

    it('opens the shortcut sheet on Ctrl+/', async () => {
      await press('/', { ctrl: true });
      const dialog = h.ui.getByRole('dialog');
      expect(dialog).toBeDefined();
      // Generated from the command table, so a rebinding cannot leave
      // the help page advertising the old key.
      expect(h.ui.getByText('Ctrl+Z')).toBeDefined();
      expect(h.ui.getByText('F10')).toBeDefined();
    });

    /**
     * The browser will not hand a worker the clipboard, so Paste is
     * the one menu item that cannot do what it says. It says so
     * rather than silently doing nothing.
     */
    it('explains Paste rather than pretending', async () => {
      await press('F10');
      await press('ArrowDown');
      await press('p');
      await press('Enter');
      expect(h.ui.getByRole('dialog')).toBeDefined();
      expect(h.ui.getByText(/Ctrl\+V/)).toBeDefined();
    });
  });
});
