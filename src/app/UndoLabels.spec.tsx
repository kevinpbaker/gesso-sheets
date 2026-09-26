import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent, OverlayService } from 'gesso-framework';
import type { UiKeyModifiers } from 'gesso-core';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Undo says what it undoes — Phase 21.
 *
 * The history always knew; the menu said "Undo" and made somebody
 * press it to find out. Asked at the service, where every command
 * names its own step, and then through the menu and the tooltip, which
 * are where it is read.
 */

function service(fill?: (document: SheetDocument) => void): { service: SheetService; document: SheetDocument } {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  return { service: new SheetService(document, { rowCount: 100, columnCount: 10 }), document };
}

function select(s: SheetService, firstRow: number, firstColumn: number, lastRow = firstRow, lastColumn = firstColumn): void {
  s.setSelection(firstRow, firstColumn, lastRow, lastColumn);
}

describe('what Undo is called', () => {
  it('is nothing with nothing to undo', () => {
    const { document } = service();
    expect(document.undoLabel).toBe('');
    expect(document.redoLabel).toBe('');
  });

  it('names the cell typed into', () => {
    const { service: s, document } = service();
    s.setCell(3, 1, 'hello');
    expect(document.undoLabel).toBe('typing in B4');
  });

  it('still calls a typed date typing, though it brought a format with it', () => {
    const { service: s, document } = service();
    s.setCell(0, 0, '2026-09-26');
    expect(document.undoLabel).toBe('typing in A1');
  });

  it('names each command by what it did', () => {
    const { service: s, document } = service(d => {
      for (let r = 0; r < 4; r++) {
        d.setCell(r, 0, String(4 - r));
      }
    });
    const cases: [string, () => void][] = [
      ['sort', () => (select(s, 0, 0, 3, 0), s.sortRange(0, true, false))],
      ['fill', () => (select(s, 0, 0), s.fill(0, 2))],
      ['clear', () => (select(s, 0, 1, 0, 2), s.clearRange())],
      ['borders', () => (select(s, 0, 0), s.setBorders('all', 1, ''))],
      ['formatting', () => s.format({ bold: true })],
      ['clear formatting', () => s.clearFormat()],
      ['merge', () => (select(s, 5, 0, 5, 1), s.setCell(5, 1, 'x'), s.mergeCells())],
      ['insert 2 rows', () => s.insertRows(10, 2)],
      ['delete column', () => s.deleteColumns(8, 1)],
      ['typing in 3 cells', () => (select(s, 20, 0, 22, 0), s.writeSelection('same'))]
    ];
    for (const [label, run] of cases) {
      run();
      expect(document.undoLabel, label).toBe(label);
    }
  });

  it('names a paste by how it was pasted', () => {
    const { service: s, document } = service(d => d.setCell(0, 0, '7'));
    select(s, 0, 0);
    s.copy(false);
    select(s, 5, 0);
    s.paste('7');
    expect(document.undoLabel).toBe('paste');
    select(s, 6, 0);
    s.pasteSpecial('values');
    expect(document.undoLabel).toBe('paste values');
  });

  it('calls a pasted cut a move', () => {
    const { service: s, document } = service(d => d.setCell(0, 0, '7'));
    select(s, 0, 0);
    s.copy(true);
    select(s, 5, 0);
    s.paste('7');
    expect(document.undoLabel).toBe('move');
  });

  it('hands the name to Redo when the step is undone, and back again', () => {
    const { service: s, document } = service(d => {
      d.setCell(0, 0, '2');
      d.setCell(1, 0, '1');
    });
    select(s, 0, 0, 1, 0);
    s.sortRange(0, true, false);
    s.undo();
    expect(document.redoLabel).toBe('sort');
    // What is left to undo is the second of the two cells typed above.
    expect(document.undoLabel).toBe('typing in A2');
    s.redo();
    expect(document.undoLabel).toBe('sort');
  });

  it('is what the status says, so the chrome can read it', () => {
    const { service: s } = service();
    let seen = '';
    s.status.subscribe(status => (seen = status.undoLabel));
    s.setCell(0, 0, 'x');
    expect(seen).toBe('typing in A1');
  });
});

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
}

describe('what the chrome says', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function mount(): Promise<void> {
    const document = new SheetDocument();
    document.setCell(0, 0, '2');
    document.setCell(1, 0, '1');
    document.sheet.recalculate();
    const s = new SheetService(document, { rowCount: 100, columnCount: 10 });
    const served = serveForTest([sheetChannel(s)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
    h = { ui, served, service: s };
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

  it('puts it in the Edit menu', async () => {
    await mount();
    h.service.setSelection(0, 0, 1, 0);
    h.service.sortRange(0, true, false);
    await settle();
    // The keyboard on the sheet, where F10 is answered.
    h.ui.fireEvent.focus(h.ui.getByRole('grid'));
    await settle();

    await press('F10');
    await press('ArrowRight');
    await press('ArrowDown');

    expect(h.ui.getByRole('menuitem', { name: 'Undo sort' })).toBeDefined();
    expect(h.ui.getByRole('menuitem', { name: 'Redo', disabled: true })).toBeDefined();
  });

  it('puts it in the toolbar button’s tooltip, with the key', async () => {
    await mount();
    h.service.setCell(4, 2, 'x');
    await settle();

    const button = h.ui.getByRole('button', { name: 'Undo' });
    const box = h.ui.getLayout(button);
    h.ui.fireEvent.pointerMove(box.x + box.width / 2, box.y + box.height / 2);
    await new Promise(resolve => setTimeout(resolve, 450));
    await settle();

    expect(h.ui.runtime.services.get(OverlayService).entries.value.length).toBeGreaterThan(0);
    expect(h.ui.queryByText('Undo typing in C5 (Ctrl+Z)')).not.toBeNull();
  });
});
