import { describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';
import type { UiKeyModifiers, UiNode } from 'gesso-core';

import { relativeRef } from '../sheet/A1';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Insert ▸ Names: every name, to read, change and take away — ranges
 * and formulas in one list, through the render tree.
 */
describe('the names dialog', () => {
  let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument };

  async function mount(): Promise<void> {
    const document = new SheetDocument();
    document.setCell(0, 0, '10');
    document.setCell(1, 0, '20');
    document.defineName('Sales', { start: relativeRef(0, 0), end: relativeRef(1, 0) });
    document.defineFormulaName('Double', '=LAMBDA(x, x*2)');
    document.setCell(0, 2, '=SUM(Sales)');
    document.setCell(1, 2, '=Double(A1)');
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1000, height: 700 });
    h = { ui, served, service, document };
    await settle();
    h.ui.fireEvent.focus(h.ui.getByRole('grid'));
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

  function shown(text: string | RegExp): UiNode[] {
    try {
      return h.ui.getAllByText(text);
    } catch {
      return [];
    }
  }

  async function openNames(): Promise<void> {
    await press('F10');
    await press('i');
    h.ui.fireEvent.click(h.ui.getByRole('menuitem', { name: 'Names…' }));
    await settle();
  }

  async function replace(field: string, text: string): Promise<void> {
    h.ui.fireEvent.focus(h.ui.getByRole('textbox', { name: field }));
    await press('a', { ctrl: true });
    h.ui.fireEvent.type(text);
    await settle();
  }

  async function click(label: string): Promise<void> {
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: label }));
    await settle();
  }

  it('lists ranges and formulas together, with what each holds', async () => {
    await mount();
    await openNames();
    expect(shown('=$A$1:$A$2')).toHaveLength(1);
    // In the list, and in the field, because the first in order is
    // chosen, ready to edit.
    expect(shown('=LAMBDA(x, x*2)')).toHaveLength(2);
    expect(h.ui.getByRole('textbox', { name: 'Name' }).properties.get('value')).toBe('Double');
  });

  it('changes what a function does, and every caller follows', async () => {
    await mount();
    await openNames();
    await replace('Refers to', '=LAMBDA(x, x*3)');
    await click('Save');
    h.document.sheet.recalculate();
    expect(h.document.sheet.value(1, 2)).toBe(30);
  });

  it('turns a formula into a range when a range is written', async () => {
    await mount();
    await openNames();
    await click('New name');
    await replace('Name', 'Figures');
    await replace('Refers to', '=A1:A2');
    await click('Save');
    expect(h.document.sheet.names.rangeOf('Figures')).toMatchObject({ start: { row: 0 }, end: { row: 1 } });
  });

  it('renames in one step that one undo takes back', async () => {
    await mount();
    await openNames();
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Sales' }));
    await settle();
    await replace('Name', 'Takings');
    await click('Save');
    expect(h.document.sheet.names.rangeOf('Takings')).not.toBeNull();
    expect(h.document.sheet.names.rangeOf('Sales')).toBeNull();
    h.document.undo();
    expect(h.document.sheet.names.rangeOf('Sales')).not.toBeNull();
    expect(h.document.sheet.names.rangeOf('Takings')).toBeNull();
  });

  it('deletes one', async () => {
    await mount();
    await openNames();
    await click('Delete');
    expect(h.document.sheet.names.formulaOf('Double')).toBeNull();
    expect(shown('=LAMBDA(x, x*2)')).toHaveLength(0);
  });

  it('says why it will not save, and keeps what was typed', async () => {
    await mount();
    await openNames();
    await replace('Refers to', '=LAMBDA(x, x*');
    await click('Save');
    expect(shown(/does not parse/)).toHaveLength(1);
    expect(h.document.sheet.names.formulaOf('Double')).not.toBeNull();
    expect(h.ui.getByRole('textbox', { name: 'Refers to' }).properties.get('value')).toBe('=LAMBDA(x, x*');
  });
});
