import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiKeyModifiers } from 'gesso-core';

import { colourName, describeConditional, describeValidation, paintName } from './RulesBar';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';
import { SHEET, seed } from './SheetSeed';
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

  /** One of a dropdown's options, chosen as a person would: open it, click the option. */
  async function choose(dropdown: string, option: string): Promise<void> {
    h.ui.fireEvent.click(h.ui.getByRole('combobox', { name: dropdown }));
    await h.ui.settle();
    h.ui.fireEvent.click(h.ui.getByRole('option', { name: option }));
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  /** What a dropdown shows chosen. */
  const shown = (dropdown: string) => h.ui.getByRole('combobox', { name: dropdown }).properties.get('valueText');

  /** Types into a field, after putting the keyboard there. */
  async function typeInto(field: string, text: string): Promise<void> {
    h.ui.fireEvent.focus(h.ui.getByRole('textbox', { name: field }));
    await h.ui.settle();
    await type(text);
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
      await choose('Colour', 'white to red');
      await click('Add rule');

      // The smallest is the scale's first stop and the largest its last.
      expect(fillOf(0, 0)).toBe('#ffffff');
      expect(fillOf(1, 0)).toBe('#c5221f');
    });

    it('takes a custom formula, moved to each cell', async () => {
      await menu('o', 'Conditional formatting…');
      await choose('Condition', 'makes a formula true');
      await typeInto('Rule formula', '=A1>5');
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

    it('opens on the rule the active cell is already under', async () => {
      await menu('o', 'Conditional formatting…');
      await choose('Condition', 'is less than');
      await choose('Colour', 'green fill');
      await typeInto('Value', '3');
      await press('Enter');
      await press('Escape');

      // Off the range, where there is no rule, and back onto it.
      h.service.setSelection(10, 5, 10, 5);
      await h.served.settle();
      await h.ui.settle();
      await menu('o', 'Conditional formatting…');
      expect(shown('Condition')).toBe('is greater than');
      expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('');
      await press('Escape');

      h.service.setSelection(1, 0, 1, 0);
      await h.served.settle();
      await h.ui.settle();
      await menu('o', 'Conditional formatting…');
      expect(shown('Condition')).toBe('is less than');
      expect(shown('Colour')).toBe('green fill');
      expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('3');
      // And it is the rule being changed, not a new one to add.
      expect(h.ui.getByRole('button', { name: 'Save changes' })).toBeDefined();
    });

    /**
     * The bar is wider than a window of 900 pixels, and its one field
     * was the only thing in it allowed to shrink — so on anything much
     * narrower than a full-width monitor the value a rule had been
     * given was drawn zero pixels wide, and the bar looked as if it had
     * not read the rule at all. It wraps now, and the field keeps its
     * width.
     */
    it('keeps its value field whole on a window too narrow for one line', async () => {
      await menu('o', 'Conditional formatting…');
      const field = h.ui.getLayout(h.ui.getByRole('textbox', { name: 'Value' }));
      const form = h.ui.getLayout(h.ui.getByRole('form'));
      expect(field.width).toBe(110);
      expect(field.x + field.width).toBeLessThanOrEqual(900);
      // And the grid moves down for the second line rather than being
      // drawn over by it.
      expect(h.ui.getLayout(h.ui.getByRole('grid')).y).toBeGreaterThanOrEqual(form.y + form.height);
    });

    it('opens on a colour scale as a colour scale', async () => {
      await menu('o', 'Conditional formatting…');
      h.ui.fireEvent.click(h.ui.getByRole('radio', { name: 'Colour scale' }));
      await h.ui.settle();
      await choose('Colour', 'white to blue');
      await click('Add rule');
      await press('Escape');

      h.service.setSelection(10, 5, 10, 5);
      h.service.setSelection(0, 0, 0, 0);
      await h.served.settle();
      await h.ui.settle();
      await menu('o', 'Conditional formatting…');
      expect(h.ui.getByRole('radio', { name: 'Colour scale' })).toHaveSemantics({ states: ['checked'] });
      expect(shown('Colour')).toBe('white to blue');
    });

    /** The bar says which cells, and takes others typed in its place. */
    it('says the cells it is for, and takes a range written in their place', async () => {
      await menu('o', 'Conditional formatting…');
      expect(h.ui.getByRole('textbox', { name: 'Cells' })).toHaveText('A1:A5');
      await typeInto('Value', '5');
      h.ui.fireEvent.focus(h.ui.getByRole('textbox', { name: 'Cells' }));
      await press('a', { ctrl: true });
      await type('A2:A3');
      await press('Enter');
      expect(h.document.conditional.map(rule => rule.range)).toMatchObject([
        { start: { row: 1, column: 0 }, end: { row: 2, column: 0 } }
      ]);
    });

    it('refuses a range that is not one, and says how to write it', async () => {
      await menu('o', 'Conditional formatting…');
      await typeInto('Value', '5');
      h.ui.fireEvent.focus(h.ui.getByRole('textbox', { name: 'Cells' }));
      await press('a', { ctrl: true });
      await type('the top ones');
      await press('Enter');
      expect(h.document.conditional).toHaveLength(0);
      expect(h.ui.getAllByText(/is not a range on this sheet/)).toHaveLength(1);
    });

    /** A1 is 1 and A2 is 9, under A1:A5. */
    it('counts the cells the rule would colour as it is written', async () => {
      await menu('o', 'Conditional formatting…');
      await typeInto('Value', '5');
      expect(h.ui.getAllByText('1 of 5 cells match')).toHaveLength(1);
      await press('Backspace');
      await type('0');
      expect(h.ui.getAllByText('2 of 5 cells match')).toHaveLength(1);
    });

    it('changes the rule it opened on rather than adding another', async () => {
      await menu('o', 'Conditional formatting…');
      await typeInto('Value', '5');
      await press('Enter');
      expect(h.ui.getByRole('button', { name: 'Save changes' })).toBeDefined();
      await press('Backspace');
      await type('0');
      await press('Enter');
      expect(h.document.conditional).toHaveLength(1);
      expect(h.document.conditional[0].test).toEqual({ kind: 'greaterThan', value: 0 });
    });

    it('lists the sheet’s rules as sentences, to change or remove one at a time', async () => {
      h.service.addConditional({ test: { kind: 'greaterThan', value: 5 }, paint: { fill: '#fce8e6', color: '#c5221f' } });
      h.service.addConditional({ test: { kind: 'lessThan', value: 2 }, paint: { fill: '#e6f4ea', color: '#137333' } });
      await h.served.settle();
      await menu('o', 'Conditional formatting…');
      await click('Rules on this sheet');
      const first = 'Highlight · A1:A5: greater than 5, in red fill';
      const second = 'Highlight · A1:A5: less than 2, in green fill';
      expect(h.ui.getByRole('listitem', { name: first })).toBeDefined();
      expect(h.ui.getByRole('listitem', { name: second })).toBeDefined();

      await click(`Edit ${first}`);
      expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('5');
      expect(h.ui.getByRole('button', { name: 'Save changes' })).toBeDefined();

      await click(`Delete ${second}`);
      expect(h.document.conditional.map(rule => rule.test)).toEqual([{ kind: 'greaterThan', value: 5 }]);
    });

    /**
     * The bar opens on the rule on top, and a second rule under it — red
     * below a target, green above — was left for nobody to find but the
     * cell's own colour. So a cell under more than one lists them all.
     */
    it('lists every rule the active cell is under, when there is more than one', async () => {
      h.service.addConditional({ test: { kind: 'greaterThan', value: 5 }, paint: { fill: '#fce8e6', color: '#c5221f' } });
      await h.served.settle();
      await menu('o', 'Conditional formatting…');
      // One rule is the sentence itself; there is nothing more to list.
      expect(h.ui.queryByRole('listitem')).toBeNull();
      await press('Escape');

      h.service.addConditional({ test: { kind: 'lessThan', value: 2 }, paint: { fill: '#e6f4ea', color: '#137333' } });
      await h.served.settle();
      await menu('o', 'Conditional formatting…');
      const first = 'Highlight · A1:A5: greater than 5, in red fill';
      const second = 'Highlight · A1:A5: less than 2, in green fill';
      expect(h.ui.getAllByRole('listitem').map(item => item.properties.get('label'))).toEqual([first, second]);
      // It opened on the one on top, and the other is a click away.
      expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('2');
      await click(`Edit ${first}`);
      expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('5');

      await click(`Delete ${second}`);
      expect(h.document.conditional.map(rule => rule.test)).toEqual([{ kind: 'greaterThan', value: 5 }]);
      expect(h.ui.queryByRole('listitem')).toBeNull();
    });

    it('keeps changing the same rule when one before it is removed', async () => {
      h.service.addConditional({ test: { kind: 'greaterThan', value: 5 }, paint: { fill: '#fce8e6', color: '#c5221f' } });
      h.service.addConditional({ test: { kind: 'lessThan', value: 2 }, paint: { fill: '#e6f4ea', color: '#137333' } });
      await h.served.settle();
      await menu('o', 'Conditional formatting…');
      await click('Delete Highlight · A1:A5: greater than 5, in red fill');
      // Typed after the 2 it opened with.
      await typeInto('Value', '3');
      await click('Save changes');
      expect(h.document.conditional.map(rule => rule.test)).toEqual([{ kind: 'lessThan', value: 32 }]);
    });

    /** A highlight being changed must not be saved as a scale by switching the tab. */
    it('starts a new rule when the kind changes, rather than turning one into another', async () => {
      await menu('o', 'Conditional formatting…');
      await typeInto('Value', '5');
      await press('Enter');
      h.ui.fireEvent.click(h.ui.getByRole('radio', { name: 'Colour scale' }));
      await h.ui.settle();
      expect(h.ui.getByRole('button', { name: 'Add rule' })).toBeDefined();
      await click('Add rule');
      expect(h.document.conditional.map(rule => (rule.scale === undefined ? 'highlight' : 'scale'))).toEqual(['highlight', 'scale']);
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

    /** `, 10` is "at most ten": the empty first part is the minimum nobody set. */
    it('takes a number rule with only a maximum, and shows it back that way', async () => {
      await menu('d', 'Data validation…');
      await choose('Allow', 'a number');
      await typeInto('Allowed values', ', 10');
      await press('Enter');
      await press('Escape');
      expect(h.document.validationAt(0, 0)?.rule).toEqual({ kind: 'number', max: 10 });

      h.service.setSelection(10, 5, 10, 5);
      h.service.setSelection(0, 0, 0, 0);
      await h.served.settle();
      await h.ui.settle();
      await menu('d', 'Data validation…');
      expect(h.ui.getByRole('textbox', { name: 'Allowed values' })).toHaveText(', 10');
    });

    it('opens on the validation the active cell is already under', async () => {
      await menu('d', 'Data validation…');
      await type('North, South');
      await press('Enter');

      h.service.setSelection(10, 5, 10, 5);
      h.service.setSelection(0, 0, 0, 0);
      await h.served.settle();
      await h.ui.settle();
      await menu('d', 'Data validation…');
      expect(shown('Allow')).toBe('one of a list');
      expect(h.ui.getByRole('textbox', { name: 'Allowed values' })).toHaveText('North, South');
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
      await typeInto('Allowed values', 'North, South');
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
      await typeInto('Allowed values', 'North, South');
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
describe('the rules bar over the seeded workbook', () => {
  let h: Harness;

  beforeEach(async () => {
    // Onto the sales sheet, where the seed's rules are; it opens on the dashboard.
    h = await mount(document => {
      seed(document);
      document.activate(SHEET.sales);
    });
  });

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function settle(): Promise<void> {
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function openFormat(): Promise<void> {
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('F10');
    await settle();
    h.ui.fireEvent.press('o');
    await settle();
    h.ui.fireEvent.click(h.ui.getByRole('menuitem', { name: 'Conditional formatting…' }));
    await settle();
  }

  /**
   * The revenue column's scale is pink to yellow to green, which is
   * none of the bar's swatches — so it is shown as itself, and not as
   * the nearest one.
   */
  it('shows a scale in colours of its own as its own, down the column', async () => {
    h.service.setSelection(4, 6, 4, 6);
    await settle();
    await openFormat();
    for (const row of [4, 5, 6]) {
      h.service.setSelection(row, 6, row, 6);
      await settle();
      expect(h.ui.getByRole('radio', { name: 'Colour scale' })).toHaveSemantics({ states: ['checked'] });
      // Its own colours, named: none of the bar's.
      expect(h.ui.getByRole('combobox', { name: 'Colour' }).properties.get('valueText')).toBe('pale red to pale yellow to pale green');
    }
  });

  it('puts the scale back as it was when applied unchanged', async () => {
    h.service.setSelection(4, 6, 4, 6);
    await settle();
    const before = h.document.conditionalAt(4, 6)?.scale;
    await openFormat();
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Save changes' }));
    await settle();
    expect(h.document.conditionalAt(4, 6)?.scale).toEqual(before);
  });

  it('lists both of the dashboard’s rules on a percentage of target', async () => {
    h.document.activate(SHEET.dashboard);
    h.service.setSelection(8, 6, 8, 6);
    await settle();
    await openFormat();
    expect(h.ui.getAllByRole('listitem').map(item => item.properties.get('label'))).toEqual([
      expect.stringContaining('less than 1'),
      expect.stringContaining('greater than 0.999')
    ]);
  });

  it('shows bold coloured text as the current colours, not a fill', async () => {
    h.service.setSelection(4, 8, 4, 8);
    await settle();
    await openFormat();
    expect(h.ui.getByRole('combobox', { name: 'Condition' }).properties.get('valueText')).toBe('contains the text');
    expect(h.ui.getByRole('textbox', { name: 'Value' })).toHaveText('Below');
    expect(h.ui.getByRole('combobox', { name: 'Colour' }).properties.get('valueText')).toBe('bold red text');
  });
});

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

describe('a rule, in words', () => {
  it('names a colour for a person, not as a code', () => {
    expect(colourName('#fde2e2')).toBe('pale red');
    expect(colourName('#c5221f')).toBe('red');
    expect(colourName('#1967d2')).toBe('blue');
    expect(colourName('#f1f3f4')).toBe('pale grey');
    expect(colourName('#ffffff')).toBe('white');
  });

  it('says what a highlight paints', () => {
    expect(paintName({ fill: '#fce8e6', color: '#c5221f' })).toBe('red fill');
    expect(paintName({ color: '#c5221f', bold: true })).toBe('bold red text');
  });

  it('reads a rule as a sentence', () => {
    expect(describeConditional({ range: 'E4:E27', test: { kind: 'between', low: 1, high: 9 }, paint: { fill: '#e6f4ea', color: '#137333' } })).toBe(
      'Highlight · E4:E27: between 1 and 9, in green fill'
    );
    expect(describeConditional({ range: 'G4:G27', test: null, scale: { from: '#ffffff', to: '#137333' } })).toBe(
      'Colour scale · G4:G27: shaded white to green, lowest to highest'
    );
    expect(describeValidation({ range: 'J4:J27', rule: { kind: 'list', values: ['Yes', 'No'] }, strict: true })).toBe(
      'Limit · J4:J27: one of Yes, No, and nothing else'
    );
  });
});
