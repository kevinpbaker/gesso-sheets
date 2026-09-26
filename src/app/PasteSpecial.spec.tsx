import { afterEach, describe, expect, it } from 'vitest';

import { createComponent, ShellService } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import type { UiKeyModifiers } from 'gesso-core';

import { GENERAL, PLAIN, type CellFormat } from '../sheet/Format';
import { SheetApp } from './SheetApp';
import type { SheetMarked } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Phase 20: paste that asks.
 *
 * Every paste through the same harness the clipboard specs use — the
 * text the shell was asked to put on the clipboard, handed back as the
 * Paste event a browser would send — and each undone in one step.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  document: SheetDocument;
  service: SheetService;
  copied: () => string;
}

const money: CellFormat = { number: { kind: 'currency', places: 2, symbol: '$' }, paint: { ...PLAIN, bold: true } };

async function mount(fill?: (document: SheetDocument) => void): Promise<Harness> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 360 });
  const writes: string[] = [];
  ui.runtime.services.get(ShellService).setHandler(request => {
    if (request.type === 'clipboard') {
      writes.push(request.text);
    }
  });
  await ui.settle();
  await served.settle();
  await ui.settle();
  const grid = ui.getByRole('grid');
  let stops = 0;
  while (ui.runtime.input.focus.focusedNode !== grid) {
    if (stops++ > 12) {
      throw new Error('Tab never reached the grid');
    }
    ui.fireEvent.tab();
    await ui.settle();
  }
  return { ui, served, document, service, copied: () => writes[writes.length - 1] ?? '' };
}

let h: Harness;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
});

async function settle(): Promise<void> {
  await h.ui.settle();
  await h.served.settle();
  await h.ui.settle();
}

async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
  h.ui.fireEvent.press(key, modifiers);
  await settle();
}

async function select(row: number, column: number, anchorRow = row, anchorColumn = column): Promise<void> {
  h.service.setSelection(row, column, anchorRow, anchorColumn);
  await settle();
}

async function paste(text: string): Promise<void> {
  h.ui.fireEvent.paste(text);
  await settle();
}

const marked = (): SheetMarked | null => {
  let seen: SheetMarked | null = null;
  h.service.clipboard.subscribe(clipboard => (seen = clipboard.marked)).unsubscribe();
  return seen;
};

const numbers = (document: SheetDocument) => {
  document.setFormat(0, 0, money);
  document.setCell(0, 0, '2');
  document.setCell(1, 0, '=A1*10');
};

describe('pasting less than everything', () => {
  it('pastes values only: a formula lands as its answer, into the formats already there', async () => {
    h = await mount(numbers);
    await select(0, 0, 1, 0);
    await press('c', { ctrl: true });
    await select(0, 2);
    h.service.pasteSpecial('values');
    await settle();
    expect(h.document.sheet.input(0, 2)).toBe('2');
    expect(h.document.sheet.input(1, 2)).toBe('20');
    expect(h.document.formatAt(0, 2).number).toEqual(GENERAL);

    h.service.undo();
    await settle();
    expect(h.document.sheet.input(1, 2)).toBe('');
  });

  it('pastes values with Ctrl+Shift+V, which is the browser’s paste as plain text', async () => {
    h = await mount(numbers);
    await select(0, 0, 1, 0);
    await press('c', { ctrl: true });
    await select(0, 3);
    h.ui.fireEvent.keyDown('V', { ctrl: true, shift: true });
    await paste(h.copied());
    expect(h.document.sheet.input(1, 3)).toBe('20');
    // And the next paste is an ordinary one again.
    await select(5, 3);
    await paste(h.copied());
    expect(h.document.sheet.input(6, 3)).toBe('=D6*10');
  });

  it('pastes formats only, and leaves what the cells hold', async () => {
    h = await mount(d => {
      numbers(d);
      d.setCell(0, 4, '7');
    });
    await select(0, 0);
    await press('c', { ctrl: true });
    await select(0, 4);
    h.service.pasteSpecial('formats');
    await settle();
    expect(h.document.sheet.input(0, 4)).toBe('7');
    expect(h.document.formatAt(0, 4)).toEqual(money);
  });

  it('pastes transposed, each formula moved by how far its own cell went', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '1');
      d.setCell(0, 1, '2');
      d.setCell(1, 0, '=A1+B1');
    });
    await select(0, 0, 1, 1);
    await press('c', { ctrl: true });
    await select(4, 4);
    h.service.pasteSpecial('transposed');
    await settle();
    // A1 → E5, B1 → E6, A2 → F5, and the formula moved as far as its
    // cell did, as any paste moves it.
    expect(h.document.sheet.input(4, 4)).toBe('1');
    expect(h.document.sheet.input(5, 4)).toBe('2');
    expect(h.document.sheet.input(4, 5)).toBe('=F4+G4');
    expect(h.document.selection).toEqual({ row: 4, column: 4, anchorRow: 5, anchorColumn: 5 });
  });

  it('turns text from somewhere else on its side too', async () => {
    h = await mount();
    await select(0, 0);
    h.service.paste('a\tb\tc', 'transposed');
    await settle();
    expect([0, 1, 2].map(row => h.document.sheet.input(row, 0))).toEqual(['a', 'b', 'c']);
  });

  it('says why there are no formats to paste from somewhere else', async () => {
    h = await mount();
    await select(0, 0);
    h.service.paste('x', 'formats');
    await settle();
    expect(h.document.sheet.input(0, 0)).toBe('');
  });
});

describe('a cut, pasted', () => {
  it('moves the cells, and a formula elsewhere reads them in their new place', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '5');
      d.setCell(1, 0, '6');
      d.setCell(0, 5, '=SUM(A1:A2)');
      d.setCell(1, 5, '=A2*2');
    });
    await select(0, 0, 1, 0);
    await press('x', { ctrl: true });
    // Marked, and still there until it is pasted.
    expect(marked()).toMatchObject({ firstRow: 0, lastRow: 1, cut: true });
    expect(h.document.sheet.input(0, 0)).toBe('5');

    await select(4, 2);
    await paste(h.copied());
    expect(h.document.sheet.input(0, 0)).toBe('');
    expect(h.document.sheet.input(4, 2)).toBe('5');
    expect(h.document.sheet.input(0, 5)).toBe('=SUM(C5:C6)');
    expect(h.document.sheet.input(1, 5)).toBe('=C6*2');
    expect(h.document.sheet.value(0, 5)).toBe(11);
    expect(marked()).toBeNull();

    // All of it back in one step.
    h.service.undo();
    await settle();
    expect(h.document.sheet.input(0, 0)).toBe('5');
    expect(h.document.sheet.input(4, 2)).toBe('');
    expect(h.document.sheet.input(0, 5)).toBe('=SUM(A1:A2)');
  });

  it('keeps a range that was only partly cut where it was', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '5');
      d.setCell(1, 0, '6');
      d.setCell(0, 5, '=SUM(A1:A2)');
    });
    await select(0, 0);
    await press('x', { ctrl: true });
    await select(9, 9);
    await paste(h.copied());
    expect(h.document.sheet.input(0, 5)).toBe('=SUM(A1:A2)');
    expect(h.document.sheet.input(9, 9)).toBe('5');
  });

  it('carries a moved formula’s other references to the sheet it came from', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '5');
      d.setCell(1, 0, '=A1+B9');
      d.addSheet('Two');
      // A new sheet is the active one; the cut is from the first.
      d.activate(0);
    });
    await select(1, 0);
    await press('x', { ctrl: true });
    h.service.activateSheet(1);
    await settle();
    await select(0, 0);
    await paste(h.copied());
    h.document.activate(1);
    expect(h.document.sheet.input(0, 0)).toBe('=Sheet1!A1+Sheet1!B9');
  });

  it('is called off by Escape, and then pastes nothing', async () => {
    h = await mount(d => d.setCell(0, 0, '5'));
    await select(0, 0);
    await press('x', { ctrl: true });
    await press('Escape');
    expect(marked()).toBeNull();
    const text = h.copied();
    await select(3, 3);
    await paste(text);
    // Pasted as text from anywhere: a copy, not a move.
    expect(h.document.sheet.input(0, 0)).toBe('5');
  });
});

describe('the copied range, marked', () => {
  it('is marked by a copy, and stays marked through a paste', async () => {
    h = await mount(numbers);
    await select(0, 0, 1, 0);
    await press('c', { ctrl: true });
    expect(marked()).toMatchObject({ sheet: 0, firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 0, cut: false });
    await select(0, 3);
    await paste(h.copied());
    expect(marked()).not.toBeNull();
  });

  it('is unmarked by an edit, and by Escape', async () => {
    h = await mount(numbers);
    await select(0, 0);
    await press('c', { ctrl: true });
    h.service.setCell(5, 5, 'typed');
    await settle();
    expect(marked()).toBeNull();

    await press('c', { ctrl: true });
    expect(marked()).not.toBeNull();
    await press('Escape');
    expect(marked()).toBeNull();
  });
});

describe('the format painter', () => {
  const brush = () => h.ui.getByRole('button', { name: 'Format painter' });

  it('takes the selection’s formats to the next thing clicked, and goes out', async () => {
    h = await mount(d => {
      numbers(d);
      d.setCell(0, 3, 'target');
    });
    await select(0, 0);
    h.ui.fireEvent.click(brush());
    await settle();
    h.ui.fireEvent.click(h.ui.getByRole('cell', { name: 'target' }));
    await settle();
    expect(h.document.formatAt(0, 3)).toEqual(money);
    expect(h.document.sheet.input(0, 3)).toBe('target');
    expect(brush()).toHaveSemantics({ states: [] });
  });

  it('stays lit after a double-click, until Escape', async () => {
    h = await mount(numbers);
    await select(0, 0);
    h.ui.fireEvent.doubleClick(brush());
    await settle();
    expect(brush()).toHaveSemantics({ states: ['pressed'] });
    await press('Escape');
    expect(brush()).toHaveSemantics({ states: [] });
  });

  it('leaves the clipboard as it was', async () => {
    h = await mount(numbers);
    await select(1, 0);
    await press('c', { ctrl: true });
    await select(0, 0);
    h.ui.fireEvent.click(brush());
    await settle();
    expect(marked()).toMatchObject({ firstRow: 1 });
  });
});
