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
 * The tab strip, from the keyboard.
 *
 * Phase 4's standard again: no pointer event in the file, every
 * assertion reached by pressing the key a person would press and read
 * back from what a screen reader would hear or from what the
 * application worker ended up holding.
 *
 * The strip is one tab stop with the arrows moving inside it — the
 * third time this application has made that decision, and here for
 * the plainest reason of the three: a workbook holds up to 256
 * sheets, and as tab stops that is a grid 256 presses away.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
  service: SheetService;
}

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served, document, service };
}

describe('the tab strip', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  beforeEach(async () => {
    h = await mount();
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

  /**
   * Alt+F10 from the sheet, which is how a person gets to the strip.
   *
   * **Not Tab**, and the difference matters. Tab moves the selection
   * one cell right, so the grid consumes it and never hands the
   * keyboard on — which a spec that drives focus directly cannot
   * see, and a browser finds in one press. So the route asserted
   * here is the route that exists.
   */
  async function reachTheStrip(): Promise<void> {
    await reachTheGrid();
    await press('F10', { alt: true });
    if (h.ui.runtime.input.focus.focusedNode !== h.ui.getByRole('tablist')) {
      throw new Error('Alt+F10 did not reach the tab strip');
    }
  }

  /** Every tab's name, as the accessibility tree carries them. */
  const names = (): string[] => h.ui.getAllByRole('tab').map(node => String(node.properties.get('label') ?? ''));
  const selected = (): string | undefined =>
    h.ui
      .getAllByRole('tab')
      .filter(node => (node.properties.get('states') as readonly string[] | undefined)?.includes('selected'))
      .map(node => String(node.properties.get('label') ?? ''))[0];

  /** Tab until the keyboard is on the sheet, which is where F10 is answered. */
  async function reachTheGrid(): Promise<void> {
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
  }

  async function menu(mnemonic: string, item: string): Promise<void> {
    await reachTheGrid();
    await press('F10');
    await press(mnemonic);
    const entry = h.ui.getByRole('menuitem', { name: item });
    h.ui.fireEvent.click(entry);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  it('shows one tab for a new workbook', () => {
    expect(names()).toEqual(['Sheet1']);
    expect(selected()).toBe('Sheet1');
  });

  /**
   * The strip is a region, and a region a keyboard cannot reach is
   * one that only half the people using this have.
   */
  it('is reached from the sheet with Alt+F10', async () => {
    await reachTheGrid();
    await press('F10', { alt: true });
    expect(h.ui.runtime.input.focus.focusedNode).toBe(h.ui.getByRole('tablist'));
  });

  /** Plain Tab is the grid's: it moves one cell right, as it should. */
  it('is not reached by Tab, because the sheet answers that', async () => {
    await reachTheGrid();
    const before = h.document.selection.column;
    await press('Tab');
    expect(h.document.selection.column).toBe(before + 1);
    expect(h.ui.runtime.input.focus.focusedNode).toBe(h.ui.getByRole('grid'));
  });

  it('is one tab stop, not one per sheet', async () => {
    await menu('s', 'Insert sheet');
    await menu('s', 'Insert sheet');
    expect(names()).toHaveLength(3);
    expect(h.ui.getAllByRole('tablist')).toHaveLength(1);
  });

  it('says how many sheets there are', async () => {
    const line = () =>
      h.ui
        .allNodes()
        .map(node => String(node.properties.get('text') ?? ''))
        .filter(text => text.endsWith('sheet') || text.endsWith('sheets'));
    expect(line()).toContain('1 sheet');
    await menu('s', 'Insert sheet');
    expect(line()).toContain('2 sheets');
  });

  describe('the arrows inside it', () => {
    beforeEach(async () => {
      await menu('s', 'Insert sheet');
      await menu('s', 'Insert sheet');
      await menu('s', 'Previous sheet');
      await menu('s', 'Previous sheet');
      await reachTheStrip();
    });

    /** Arrowing onto a tab shows it, which is what Ctrl+PageDown taught. */
    it('shows the next sheet on the right arrow', async () => {
      expect(selected()).toBe('Sheet1');
      await press('ArrowRight');
      expect(selected()).toBe('Sheet2');
      await press('ArrowRight');
      expect(selected()).toBe('Sheet3');
    });

    it('goes back on the left arrow', async () => {
      await press('End');
      await press('ArrowLeft');
      expect(selected()).toBe('Sheet2');
    });

    it('jumps to the ends with Home and End', async () => {
      await press('End');
      expect(selected()).toBe('Sheet3');
      await press('Home');
      expect(selected()).toBe('Sheet1');
    });

    /**
     * Past the last tab is the add button, which is not a sheet — so
     * the arrow stops there rather than wrapping, and the sheet in
     * view does not change under somebody reaching for it.
     */
    it('stops on the add button rather than wrapping past it', async () => {
      await press('End');
      await press('ArrowRight');
      expect(selected()).toBe('Sheet3');
      await press('ArrowRight');
      expect(selected()).toBe('Sheet3');
    });

    it('adds a sheet with Enter on the add button', async () => {
      await press('End');
      await press('ArrowRight');
      await press('Enter');
      expect(names()).toEqual(['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']);
    });

    it('leaves the strip on Tab', async () => {
      h.ui.fireEvent.tab();
      await h.ui.settle();
      expect(h.ui.runtime.input.focus.focusedNode).not.toBe(h.ui.getByRole('tablist'));
    });
  });

  describe('renaming', () => {
    it('opens a box on F2 with the name in it', async () => {
      await reachTheStrip();
      await press('F2');

      const box = h.ui.getByRole('textbox', { name: 'Sheet name' });
      expect(textProperty(box)).toBe('Sheet1');
      expect(h.ui.runtime.input.focus.focusedNode).toBe(box);
    });

    it('takes the new name on Enter', async () => {
      await reachTheStrip();
      await press('F2');
      await type('Figures');
      await press('Enter');

      expect(names()).toEqual(['Figures']);
      expect(h.document.sheets()[0].name).toBe('Figures');
    });

    it('puts the old name back on Escape', async () => {
      await reachTheStrip();
      await press('F2');
      await type('Figures');
      await press('Escape');

      expect(names()).toEqual(['Sheet1']);
      expect(h.ui.queryByRole('textbox', { name: 'Sheet name' })).toBeNull();
    });

    /** From the sheet, without going near the strip or the menu. */
    it('opens on Shift+F2 from the grid', async () => {
      await reachTheGrid();
      await press('F2', { shift: true });

      const box = h.ui.getByRole('textbox', { name: 'Sheet name' });
      expect(h.ui.runtime.input.focus.focusedNode).toBe(box);
      // Plain F2 is the cell editor and has to stay that way.
      await press('Escape');
      await reachTheGrid();
      await press('F2');
      expect(h.ui.queryByRole('textbox', { name: 'Sheet name' })).toBeNull();
      expect(h.ui.getByRole('textbox', { name: 'Cell' })).toBeDefined();
    });

    /** The menu is the other way in, and has to reach the same box. */
    it('opens the same box from the Sheet menu', async () => {
      await menu('s', 'Rename sheet…');
      expect(h.ui.getByRole('textbox', { name: 'Sheet name' })).toBeDefined();
    });

    it('keeps the name when the box is left empty', async () => {
      await reachTheStrip();
      await press('F2');
      await type('   ');
      await press('Enter');
      expect(names()).toEqual(['Sheet1']);
    });
  });

  describe('the Sheet menu', () => {
    it('inserts a sheet and shows it', async () => {
      await menu('s', 'Insert sheet');
      expect(names()).toEqual(['Sheet1', 'Sheet2']);
      expect(selected()).toBe('Sheet2');
    });

    it('duplicates one, cells and all', async () => {
      h.service.setCell(0, 0, 'copy me');
      await h.served.settle();
      await menu('s', 'Duplicate sheet');

      expect(names()).toEqual(['Sheet1', 'Sheet1 copy']);
      expect(h.document.sheet.input(0, 0)).toBe('copy me');
    });

    it('moves one along the strip', async () => {
      await menu('s', 'Insert sheet');
      await menu('s', 'Move sheet left');
      expect(names()).toEqual(['Sheet2', 'Sheet1']);
      expect(selected()).toBe('Sheet2');
    });

    it('does not move one off the end', async () => {
      await menu('s', 'Move sheet left');
      expect(names()).toEqual(['Sheet1']);
    });

    it('steps through the sheets, wrapping at the ends', async () => {
      await menu('s', 'Insert sheet');
      await menu('s', 'Next sheet');
      expect(selected()).toBe('Sheet1');
      await menu('s', 'Previous sheet');
      expect(selected()).toBe('Sheet2');
    });

    it('colours a tab, and takes the colour off again', async () => {
      await menu('s', 'Tab colour: red');
      expect(h.document.sheets()[0].colour).toBe('#ea4335');
      await menu('s', 'Tab colour: none');
      expect(h.document.sheets()[0].colour).toBeNull();
    });
  });

  describe('deleting', () => {
    /** The one destructive thing here, so the one thing that asks. */
    it('asks before it deletes', async () => {
      await menu('s', 'Insert sheet');
      await reachTheStrip();
      await press('Delete');

      expect(h.ui.getByRole('dialog', { name: 'Delete sheet' })).toBeDefined();
      expect(names()).toHaveLength(2);
    });

    it('deletes when the question is answered', async () => {
      await menu('s', 'Insert sheet');
      await reachTheStrip();
      await press('Delete');

      h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Delete sheet' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(names()).toEqual(['Sheet1']);
    });

    it('leaves it alone when the question is declined', async () => {
      await menu('s', 'Insert sheet');
      await reachTheStrip();
      await press('Delete');

      h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Cancel' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(names()).toEqual(['Sheet1', 'Sheet2']);
    });

    /** Nothing in the application is written to survive a workbook of none. */
    it('says so rather than emptying the workbook', async () => {
      await reachTheStrip();
      await press('Delete');
      h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Delete sheet' }));
      await h.ui.settle();
      await h.served.settle();
      await h.ui.settle();

      expect(names()).toEqual(['Sheet1']);
    });
  });

  /**
   * The strip is a view of the sheets, so anything that changes them
   * reaches it — including a command that came from somewhere else
   * entirely.
   */
  it('follows a sheet activated from outside the strip', async () => {
    await menu('s', 'Insert sheet');
    h.service.activateSheet(0);
    await h.served.settle();
    await h.ui.settle();
    expect(selected()).toBe('Sheet1');
  });
});
