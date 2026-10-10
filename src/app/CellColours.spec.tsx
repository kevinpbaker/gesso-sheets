import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { darkColors, type UiKeyModifiers } from 'gesso-core';

import { colourName } from './colourNames';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * A cell's text colour and fill, from the toolbar and the Format menu.
 *
 * The format had both from the start and nothing in the chrome could set
 * either. Two buttons now, each with a bar under it in the active cell's
 * own colour, each opening a palette beside it.
 */
describe('colouring a cell', () => {
  let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument };

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function mount(): Promise<void> {
    const document = new SheetDocument();
    document.setCell(1, 1, 'Hello');
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1200, height: 500 });
    h = { ui, served, service, document };
    await settle();
    h.service.setSelection(1, 1, 1, 1);
    await settle();
  }

  async function settle(): Promise<void> {
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
    h.ui.fireEvent.press(key, modifiers);
    await settle();
  }

  async function click(role: string, name: string): Promise<void> {
    h.ui.fireEvent.click(h.ui.getByRole(role as never, { name }));
    await settle();
  }

  /** The colour of the bar under a toolbar button, as it is drawn. */
  const bar = (button: string): unknown => {
    const node = h.ui.getByRole('button', { name: button });
    const found = h.ui.allNodes().find(each => {
      for (let at = each.parent; at !== null; at = at.parent) {
        if (at === node) {
          return each.properties.get('height') === 4;
        }
      }
      return false;
    });
    return found?.properties.get('backgroundColor');
  };

  it('fills the selection with a colour from the palette', async () => {
    await mount();
    await click('button', 'Fill colour');
    expect(h.ui.getByRole('listbox', { name: 'Fill colour' })).toBeDefined();
    await click('option', 'light cornflower blue 3');
    expect(h.document.formatAt(1, 1).paint.fill).toBe('#c9daf8');
    // And the button says so, in the cell's own colour.
    expect(bar('Fill colour')).toBe('#c9daf8');
  });

  it('colours the text, and takes it back to automatic', async () => {
    await mount();
    await click('button', 'Text colour');
    await click('option', 'dark red 1');
    expect(h.document.formatAt(1, 1).paint.color).toBe('#cc0000');
    expect(bar('Text colour')).toBe('#cc0000');
    await click('button', 'Text colour');
    await click('option', 'Automatic');
    expect(h.document.formatAt(1, 1).paint.color).toBe('');
  });

  it('takes the fill away with No fill, and one undo puts it back', async () => {
    await mount();
    await click('button', 'Fill colour');
    await click('option', 'yellow');
    await click('button', 'Fill colour');
    await click('option', 'No fill');
    expect(h.document.formatAt(1, 1).paint.fill).toBe('');
    h.document.undo();
    expect(h.document.formatAt(1, 1).paint.fill).toBe('#ffff00');
  });

  it('offers the colour just used again, at the top', async () => {
    await mount();
    await click('button', 'Fill colour');
    await click('option', 'green');
    await click('button', 'Text colour');
    expect(h.ui.getByRole('option', { name: 'recent #00ff00' })).toBeDefined();
  });

  it('opens from the Format menu, and chooses with the keyboard', async () => {
    await mount();
    h.ui.fireEvent.focus(h.ui.getByRole('grid'));
    await settle();
    await press('F10');
    await press('o');
    await click('menuitem', 'Fill colour…');
    expect(h.ui.getByRole('listbox', { name: 'Fill colour' })).toBeDefined();
    // From No fill, down into the greys, and along one.
    await press('ArrowDown');
    await press('ArrowRight');
    await press('Enter');
    expect(h.document.formatAt(1, 1).paint.fill).toBe('#434343');
  });

  it('takes any colour at all, from Custom colour…', async () => {
    await mount();
    await click('button', 'Fill colour');
    await click('option', 'Custom colour…');
    const field = h.ui.getByRole('textbox', { name: 'Fill colour as hex' });
    h.ui.fireEvent.focus(field);
    await press('a', { ctrl: true });
    h.ui.fireEvent.type('#2e7d6b');
    await settle();
    // Nothing is painted until the colour is kept.
    expect(h.document.formatAt(1, 1).paint.fill).toBe('');
    await click('button', 'Use colour');
    expect(h.document.formatAt(1, 1).paint.fill).toBe('#2e7d6b');
    // And it is one of the recent colours from then on.
    await click('button', 'Text colour');
    expect(h.ui.getByRole('option', { name: 'recent #2e7d6b' })).toBeDefined();
  });

  /**
   * A pale header on a dark sheet is drawn dark, in white ink, and is
   * still pale in the document — the toolbar's bar says so.
   */
  it('draws a light fill darkened in the dark theme, and keeps the colour it was given', async () => {
    await mount();
    h.service.format({ fill: '#f1f3f4' });
    await settle();
    const cell = () => h.ui.getByRole('cell', { name: 'Hello' });
    expect(cell().properties.get('backgroundColor')).toBe('#f1f3f4');

    h.ui.fireEvent.focus(h.ui.getByRole('grid'));
    await settle();
    await press('F10');
    await press('v');
    await click('menuitemcheckbox', 'Theme: dark');

    expect(cell().properties.get('backgroundColor')).toBe('#272c2e');
    expect(cell().properties.get('color')).toBe(darkColors.text);
    expect(h.document.formatAt(1, 1).paint.fill).toBe('#f1f3f4');
    expect(bar('Fill colour')).toBe('#f1f3f4');
  });

  it('names the colour a cell has, in the tooltip', () => {
    expect(colourName('#c9daf8')).toBe('pale blue');
  });
});
