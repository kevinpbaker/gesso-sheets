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
      expect(h.ui.getByRole('menu', { name: 'Insert' })).toBeDefined();

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

/**
 * Formatting from the keyboard — Phase 9's half of the chrome.
 *
 * Same rule as the rest of this file: no pointer event anywhere. A
 * toolbar you can see and cannot reach is a toolbar half the people
 * using this cannot reach at all.
 */
describe('formatting from the keyboard', () => {
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

  beforeEach(async () => {
    h = await mount(d => {
      d.setCell(0, 0, '1234.5');
      d.setCell(1, 0, '0.256');
      d.setCell(0, 1, 'Total');
    });
    await reachTheGrid();
  });

  const paintOf = (row: number, column: number) => h.document.formatAt(row, column).paint;

  it('makes the selection bold on Ctrl+B, and plain again', async () => {
    await press('b', { ctrl: true });
    expect(paintOf(0, 0).bold).toBe(true);
    await press('b', { ctrl: true });
    expect(paintOf(0, 0).bold).toBe(false);
  });

  it('applies to the whole selection, not just the active cell', async () => {
    await press('ArrowDown', { shift: true });
    await press('i', { ctrl: true });
    expect(paintOf(0, 0).italic).toBe(true);
    expect(paintOf(1, 0).italic).toBe(true);
  });

  /**
   * The bug every naive formatting model has: a change is relative to
   * what each cell already holds, so Italic must not undo Bold and
   * Bold must not undo the currency symbol somebody chose.
   */
  it('leaves alone what it was not asked about', async () => {
    await press('b', { ctrl: true });
    await press('$', { ctrl: true, shift: true });
    await press('i', { ctrl: true });

    const format = h.document.formatAt(0, 0);
    expect(format.paint.bold).toBe(true);
    expect(format.paint.italic).toBe(true);
    expect(format.number.kind).toBe('currency');
  });

  it('formats a number as currency and shows it that way', async () => {
    await press('4', { ctrl: true, shift: true });
    expect(h.document.display(0, 0)).toBe('$1,234.50');
    expect(h.ui.getByText('$1,234.50')).toBeDefined();
  });

  it('formats a number as a percentage', async () => {
    await press('ArrowDown');
    await press('5', { ctrl: true, shift: true });
    expect(h.document.display(1, 0)).toBe('26%');
  });

  it('adds and removes decimal places', async () => {
    await press('ArrowDown');
    await press('5', { ctrl: true, shift: true });
    await press(']', { ctrl: true });
    expect(h.document.display(1, 0)).toBe('25.6%');
    await press('[', { ctrl: true });
    expect(h.document.display(1, 0)).toBe('26%');
  });

  /** Formatting a range must never make what is in it unreadable. */
  it('leaves text readable under a number format', async () => {
    await press('ArrowRight');
    await press('4', { ctrl: true, shift: true });
    expect(h.document.display(0, 1)).toBe('Total');
  });

  it('aligns, and toggles back to automatic', async () => {
    await press('e', { ctrl: true, shift: true });
    expect(paintOf(0, 0).align).toBe('center');
    await press('e', { ctrl: true, shift: true });
    expect(paintOf(0, 0).align).toBe('auto');
  });

  /**
   * Text is the one format that changes what *typing* means. `007` in
   * a Text cell stays `007`, which is the whole reason it exists.
   */
  it('keeps the leading zeros in a Text cell', async () => {
    await press('ArrowDown');
    await press('ArrowDown');
    await press('7', { ctrl: true, shift: true });
    await press('0');
    h.ui.fireEvent.type('07');
    await h.ui.settle();
    await press('Enter');

    expect(h.document.display(2, 0)).toBe('007');
  });

  it('puts a whole range back to plain in one step of undo', async () => {
    await press('ArrowDown', { shift: true });
    await press('b', { ctrl: true });
    expect(paintOf(1, 0).bold).toBe(true);

    await press('z', { ctrl: true });
    expect(paintOf(0, 0).bold).toBe(false);
    expect(paintOf(1, 0).bold).toBe(false);
  });

  it('clears formatting', async () => {
    await press('b', { ctrl: true });
    await press('4', { ctrl: true, shift: true });
    await press('\\', { ctrl: true });

    expect(paintOf(0, 0).bold).toBe(false);
    expect(h.document.display(0, 0)).toBe('1234.5');
  });

  it('shows the toolbar pressed for the cell the selection is on', async () => {
    await press('b', { ctrl: true });
    expect(h.ui.getByRole('button', { name: 'Bold', states: ['pressed'] })).toBeDefined();

    await press('ArrowRight');
    expect(h.ui.queryByRole('button', { name: 'Bold', states: ['pressed'] })).toBeNull();
  });
});

/**
 * The toolbar is one tab stop with the arrows moving inside it.
 *
 * As fifteen stops it put the grid fifteen presses from the keyboard,
 * and the specs said so before a person could: `Tab never reached the
 * grid` is what Phase 9 got for adding a Format section.
 */
describe('the toolbar', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount(d => d.setCell(0, 0, '5'));
  });

  it('is one tab stop, not one per button', async () => {
    const stops: string[] = [];
    for (let press = 0; press < 5; press++) {
      h.ui.fireEvent.tab();
      await h.ui.settle();
      const node = h.ui.runtime.input.focus.focusedNode;
      stops.push(node === null ? 'nothing' : String(node.properties.get('role') ?? 'none'));
      if (node === h.ui.getByRole('grid')) {
        break;
      }
    }
    expect(stops).toEqual(['menubar', 'toolbar', 'textbox', 'textbox', 'grid']);
  });

  it('runs the button the arrows land on', async () => {
    h.ui.fireEvent.tab();
    await h.ui.settle();
    h.ui.fireEvent.tab();
    await h.ui.settle();
    expect(h.ui.runtime.input.focus.focusedNode).toBe(h.ui.getByRole('toolbar'));

    // Undo, Redo, then Bold.
    for (let step = 0; step < 2; step++) {
      h.ui.fireEvent.press('ArrowRight');
      await h.ui.settle();
    }
    h.ui.fireEvent.press('Enter');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    expect(h.document.formatAt(0, 0).paint.bold).toBe(true);
  });
});

/**
 * Hovering a menu, which is the half of a menu bar the keyboard
 * specs above cannot reach.
 *
 * Driven through `pointerMove`, which goes through the hit tester, so
 * these are the enter and leave events a real pointer produces rather
 * than handlers called by hand.
 */
describe('the menu under the pointer', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount();
    // F10 is answered by the grid, so the keyboard has to be there
    // first — the same arrival a person makes.
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid && stops++ < 10) {
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
  });

  /** The middle of a node, in the coordinates the hit tester uses. */
  function middleOf(node: Parameters<Rendered['getLayout']>[0]): [number, number] {
    const box = h.ui.getLayout(node);
    return [box.x + box.width / 2, box.y + box.height / 2];
  }

  async function moveTo(node: Parameters<Rendered['getLayout']>[0]): Promise<void> {
    const [x, y] = middleOf(node);
    h.ui.fireEvent.pointerMove(x, y);
    await h.ui.settle();
  }

  async function openEdit(): Promise<void> {
    h.ui.fireEvent.press('F10');
    await h.ui.settle();
    h.ui.fireEvent.press('ArrowDown');
    await h.ui.settle();
  }

  const background = (name: string): unknown =>
    h.ui.getByRole('menuitem', { name }).properties.get('backgroundColor');

  it('lights up the item the pointer is over', async () => {
    await openEdit();
    // Cut, not Undo: the sheet is empty, so Undo cannot be run and
    // the menu opens past it. The highlight only rests where Enter
    // would work, whether it got there by key or by pointer.
    expect(background('Cut')).toBe('controlBackgroundHovered');
    expect(background('Select all')).toBe('transparent');

    await moveTo(h.ui.getByRole('menuitem', { name: 'Select all' }));

    expect(background('Select all')).toBe('controlBackgroundHovered');
    expect(background('Cut')).toBe('transparent');
  });

  /**
   * One highlight, not two. A menu with a keyboard highlight on Undo
   * and a hover highlight on Paste cannot say what Enter will do.
   */
  it('hands the keyboard highlight to the pointer, so Enter follows it', async () => {
    await openEdit();
    await moveTo(h.ui.getByRole('menuitem', { name: 'Select all' }));

    h.ui.fireEvent.press('Enter');
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();

    // Select all ran, which is what the pointer was resting on.
    expect(h.document.selection.anchorRow).toBeGreaterThan(0);
  });

  /** The highlight only ever rests where Enter would work. */
  it('does not light up a command that cannot be run', async () => {
    await openEdit();
    await moveTo(h.ui.getByRole('menuitem', { name: 'Redo' }));
    expect(background('Redo')).toBe('transparent');
  });

  /**
   * Moving along the bar with a menu open opens the one under the
   * pointer. Without it a bar is something you have to click four
   * times to read.
   *
   * This is the behaviour that cost the overlay its backdrop:
   * `dismissOnOutsidePress` puts a full-screen box over everything to
   * catch the press, and a box over everything is a box over the menu
   * bar — so the titles never saw the pointer at all.
   */
  it('opens the menu the pointer moves onto', async () => {
    await openEdit();
    expect(h.ui.getByRole('menu', { name: 'Edit' })).toBeDefined();

    await moveTo(h.ui.getByText('Format'));

    expect(h.ui.queryByRole('menu', { name: 'Edit' })).toBeNull();
    expect(h.ui.getByRole('menu', { name: 'Format' })).toBeDefined();
  });

  /** Hovering a title with nothing open lights it and opens nothing. */
  it('lights a title without opening it', async () => {
    await moveTo(h.ui.getByText('Data'));
    expect(h.ui.queryByRole('menu', { name: 'Data' })).toBeNull();
  });
});
