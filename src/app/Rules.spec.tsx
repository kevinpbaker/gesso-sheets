import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiKeyModifiers } from 'gesso-core';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';
import type { SheetValidation } from './SheetContract';

/**
 * Formats that think, end to end.
 *
 * The rules themselves are specced in node; this is the path through
 * the barrier — a rule made over the selection, resolved for the
 * window, and arriving as a palette index the render worker draws
 * with. Every assertion is read from what a screen reader would hear
 * or from what the application worker ended up holding.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
  service: SheetService;
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
  return { ui, served, document, service };
}

describe('rules over a selection', () => {
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
    h.ui.fireEvent.click(h.ui.getByRole('menuitem', { name: item }));
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function click(label: string): Promise<void> {
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: label }));
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  /** What the render worker would paint a cell, after everything. */
  const fillOf = (row: number, column: number): string => {
    let entries: readonly { fill: string }[] = [];
    let ids: Record<string, Record<string, number>> = {};
    h.service.palette.subscribe(palette => (entries = palette.entries)).unsubscribe();
    h.service.formats.subscribe(formats => (ids = formats.cells as never)).unsubscribe();
    return entries[ids[row]?.[column] ?? 0]?.fill ?? '';
  };

  const validation = (): SheetValidation => {
    let seen: SheetValidation | undefined;
    h.service.validation.subscribe(value => (seen = value)).unsubscribe();
    return seen!;
  };

  describe('a conditional format', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(0, 0, '1');
        document.setCell(1, 0, '9');
      });
      h.service.setSelection(0, 0, 4, 0);
      await h.served.settle();
      await h.ui.settle();
    });

    it('is offered in the Format menu', async () => {
      await reachTheGrid();
      await press('F10');
      await press('o');
      expect(h.ui.getByRole('menuitem', { name: 'Conditional formatting…' })).toBeDefined();
    });

    it('opens a bar with the keyboard in its field', async () => {
      await menu('o', 'Conditional formatting…');
      const field = h.ui.getByRole('textbox', { name: 'Value' });
      expect(h.ui.runtime.input.focus.focusedNode).toBe(field);
    });

    it('paints the cells that match and leaves the rest', async () => {
      await menu('o', 'Conditional formatting…');
      await type('5');
      await press('Enter');

      expect(fillOf(1, 0)).toBe('#fce8e6');
      expect(fillOf(0, 0)).toBe('');
    });

    it('repaints when the value changes underneath it', async () => {
      await menu('o', 'Conditional formatting…');
      await type('5');
      await press('Enter');
      expect(fillOf(0, 0)).toBe('');

      h.service.setCell(0, 0, '7');
      await h.served.settle();
      await h.ui.settle();
      expect(fillOf(0, 0)).toBe('#fce8e6');
    });

    it('takes a colour scale over the range', async () => {
      await menu('o', 'Conditional formatting…');
      h.ui.fireEvent.click(h.ui.getByRole('radio', { name: 'Colour scale' }));
      await h.ui.settle();
      await click('Apply');

      // The smallest is the scale's first stop and the largest its last.
      expect(fillOf(0, 0)).toBe('#ffffff');
      expect(fillOf(1, 0)).toBe('#c5221f');
    });

    it('takes a custom formula, moved to each cell', async () => {
      await menu('o', 'Conditional formatting…');
      h.ui.fireEvent.click(h.ui.getByRole('radio', { name: 'Custom formula' }));
      await h.ui.settle();
      await type('=A1>5');
      await press('Enter');

      expect(fillOf(1, 0)).toBe('#fce8e6');
      expect(fillOf(0, 0)).toBe('');
    });

    it('is taken back in one press of ctrl-Z', async () => {
      await menu('o', 'Conditional formatting…');
      await type('5');
      await press('Enter');
      expect(fillOf(1, 0)).toBe('#fce8e6');

      h.service.undo();
      await h.served.settle();
      await h.ui.settle();
      expect(fillOf(1, 0)).toBe('');
    });

    it('is cleared from the Data menu', async () => {
      await menu('o', 'Conditional formatting…');
      await type('5');
      await press('Enter');
      await menu('d', 'Clear rules from this sheet');
      expect(fillOf(1, 0)).toBe('');
    });
  });

  describe('a validation', () => {
    beforeEach(async () => {
      h = await mount(document => {
        document.setCell(0, 0, 'North');
        document.setCell(1, 0, 'Nowhere');
      });
      h.service.setSelection(0, 0, 4, 0);
      await h.served.settle();
      await h.ui.settle();
    });

    it('marks the cells in view that break it', async () => {
      await menu('d', 'Data validation…');
      await type('North, South');
      await press('Enter');

      expect(validation().cells[1]?.[0]).toContain('North, South');
      expect(validation().cells[0]).toBeUndefined();
    });

    /** Only the window, which is the whole shape of this phase. */
    it('marks nothing outside the window', async () => {
      h.document.sheet.setCell(150, 0, 'Nowhere');
      await menu('d', 'Data validation…');
      await type('North, South');
      await press('Enter');
      expect(validation().cells[150]).toBeUndefined();
    });

    it('offers the list to the cell that has one', async () => {
      await menu('d', 'Data validation…');
      await type('North, South');
      await press('Enter');

      h.service.setSelection(0, 0, 0, 0);
      await h.served.settle();
      expect(validation().list).toEqual(['North', 'South']);
    });

    /**
     * Marking is the default and refusing is asked for, because a
     * rule that silently refuses what somebody typed looks like a
     * broken keyboard.
     */
    it('marks but does not refuse, unless it was told to', async () => {
      await menu('d', 'Data validation…');
      await type('North, South');
      await press('Enter');

      h.service.setCell(2, 0, 'Elsewhere');
      await h.served.settle();
      expect(h.document.sheet.input(2, 0)).toBe('Elsewhere');
    });

    it('refuses when it was told to, and says why', async () => {
      await menu('d', 'Data validation…');
      h.ui.fireEvent.click(h.ui.getByRole('checkbox', { name: 'Refuse anything else' }));
      await h.ui.settle();
      await type('North, South');
      await press('Enter');

      h.service.setCell(2, 0, 'Elsewhere');
      await h.served.settle();
      expect(h.document.sheet.input(2, 0)).toBe('');
      expect(validation().refused).toContain('North, South');
    });

    /** Emptying a cell is how a mistake is taken back. */
    it('lets a cell be emptied whatever it says', async () => {
      await menu('d', 'Data validation…');
      h.ui.fireEvent.click(h.ui.getByRole('checkbox', { name: 'Refuse anything else' }));
      await h.ui.settle();
      await type('North, South');
      await press('Enter');

      h.service.setCell(0, 0, '');
      await h.served.settle();
      expect(h.document.sheet.input(0, 0)).toBe('');
    });
  });
});

/**
 * The dropdown, which only a list has.
 *
 * It is the one kind of rule where the acceptable values are few and
 * known — which is also what makes it the kind worth enforcing.
 */
describe('the list a cell may choose from', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function press(key: string): Promise<void> {
    h.ui.fireEvent.press(key);
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

  async function openCell(): Promise<void> {
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    await press('F2');
  }

  /**
   * Every option on screen, or none.
   *
   * `getAllByRole` throws when nothing matches, and "nothing matches"
   * is half of what this describe block asserts.
   */
  const options = (): string[] =>
    h.ui
      .allNodes()
      .filter(node => node.properties.get('role') === 'option')
      .map(node => String(node.properties.get('label') ?? ''));

  beforeEach(async () => {
    const document = new SheetDocument();
    document.addValidation({
      range: { start: { row: 0, column: 0, rowAbsolute: false, columnAbsolute: false },
               end: { row: 9, column: 0, rowAbsolute: false, columnAbsolute: false } },
      rule: { kind: 'list', values: ['North', 'South', 'East'] }
    });
    const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
    await ui.settle();
    await served.settle();
    await ui.settle();
    h = { ui, served, document, service };
    service.setSelection(0, 0, 0, 0);
    await served.settle();
    await ui.settle();
  });

  it('offers the values when the cell is opened', async () => {
    await openCell();
    expect(options()).toEqual(['North', 'South', 'East']);
  });

  it('narrows to what has been typed', async () => {
    await openCell();
    await type('S');
    expect(options()).toEqual(['South']);
  });

  /** An empty box over the sheet says less than no box at all. */
  it('closes when nothing matches', async () => {
    await openCell();
    await type('Q');
    expect(options()).toEqual([]);
  });

  it('takes the one the arrows landed on', async () => {
    await openCell();
    await press('ArrowDown');
    await press('Enter');
    await h.served.settle();
    expect(h.document.sheet.input(0, 0)).toBe('South');
  });

  it('says nothing over a cell with no list', async () => {
    h.service.setSelection(0, 5, 0, 5);
    await h.served.settle();
    await h.ui.settle();
    await openCell();
    expect(options()).toEqual([]);
  });
});
