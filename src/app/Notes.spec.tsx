import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';
import type { UiKeyModifiers } from 'gesso-core';

import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Notes on a cell — Phase 22.
 *
 * A few words left on a cell: a corner mark, the text beside the cell
 * on hover and on the selected cell, Shift+F2 to write one. Saved with
 * the document and read from an .xlsx's comments.
 */

describe('a document’s notes', () => {
  it('are one step of undo each way', () => {
    const document = new SheetDocument();
    document.setNote(1, 1, 'Ask Priya');
    expect(document.undoLabel).toBe('note');
    document.undo();
    expect(document.noteAt(1, 1)).toBe('');
    document.redo();
    expect(document.noteAt(1, 1)).toBe('Ask Priya');
  });

  it('move with their cells, and come back when the move is undone', () => {
    const document = new SheetDocument();
    document.setNote(5, 0, 'on row 6');
    document.setNote(2, 0, 'on row 3');
    document.applyShift({ axis: 'row', at: 2, by: -1 });
    expect(document.noteAt(4, 0)).toBe('on row 6');
    expect(document.noteAt(2, 0)).toBe('');

    document.undo();
    expect(document.noteAt(5, 0)).toBe('on row 6');
    expect(document.noteAt(2, 0)).toBe('on row 3');
    document.redo();
    expect(document.noteAt(4, 0)).toBe('on row 6');
  });

  it('belong to their sheet', () => {
    const document = new SheetDocument();
    document.setNote(0, 0, 'first');
    document.addSheet('Second');
    document.activate(1);
    expect(document.noteAt(0, 0)).toBe('');
    document.activate(0);
    expect(document.noteAt(0, 0)).toBe('first');
  });

  it('survive being saved and opened again', () => {
    const document = new SheetDocument();
    document.setNote(3, 2, 'Line one\nline two');
    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(snapshotOf(document)), 10)!);
    expect(reopened.noteAt(3, 2)).toBe('Line one\nline two');
  });

  it('are none in a file written before there were notes', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, 'x');
    const written = JSON.parse(JSON.stringify(snapshotOf(document))) as { sheets: Record<string, unknown>[] };
    delete written.sheets[0].notes;
    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(written), 10)!);
    expect(reopened.notes.size).toBe(0);
  });
});

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

describe('notes on the sheet', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function mount(fill?: (document: SheetDocument) => void): Promise<void> {
    const document = new SheetDocument();
    document.setCell(0, 0, 'Region');
    document.setCell(1, 1, '42');
    fill?.(document);
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
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

  const shown = (cell: string) => h.ui.queryByRole('status', { name: `Note on ${cell}` });

  it('is written with Shift+F2, and saved with Ctrl+Enter', async () => {
    await mount();
    await press('ArrowDown');
    await press('ArrowRight');
    await press('F2', { shift: true });

    const field = h.ui.getByRole('textbox', { name: 'Note' });
    expect(h.ui.runtime.input.focus.focusedNode).toBe(field);
    h.ui.fireEvent.type('Checked against the invoice');
    await settle();
    await press('Enter', { ctrl: true });

    expect(h.document.noteAt(1, 1)).toBe('Checked against the invoice');
    expect(h.ui.queryByRole('textbox', { name: 'Note' })).toBeNull();
    // The keyboard is back on the sheet, and the note is beside its cell.
    expect(h.ui.runtime.input.focus.focusedNode).toBe(h.ui.getByRole('grid'));
    expect(shown('B2')).not.toBeNull();
  });

  it('opens on what the note says, and takes it away with Delete note', async () => {
    await mount(d => d.setNote(0, 0, 'Old words'));
    await press('F2', { shift: true });
    expect(h.ui.getByRole('textbox', { name: 'Note' })).toHaveText('Old words');

    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Delete note' }));
    await settle();
    expect(h.document.noteAt(0, 0)).toBe('');
    expect(shown('A1')).toBeNull();
  });

  it('leaves the note alone on Escape', async () => {
    await mount(d => d.setNote(0, 0, 'Kept'));
    await press('F2', { shift: true });
    h.ui.fireEvent.type(' and more');
    await settle();
    await press('Escape');
    expect(h.document.noteAt(0, 0)).toBe('Kept');
  });

  it('shows beside the selected cell, and only that one', async () => {
    await mount(d => {
      d.setNote(0, 0, 'About the regions');
      d.setNote(1, 1, 'About the number');
    });
    expect(shown('A1')).not.toBeNull();
    expect(shown('B2')).toBeNull();

    await press('ArrowDown');
    expect(shown('A1')).toBeNull();
  });

  it('shows beside the cell the pointer rests on', async () => {
    await mount(d => d.setNote(1, 1, 'Hovered'));
    await press('ArrowDown');
    await press('ArrowDown');
    expect(shown('B2')).toBeNull();

    const cell = h.ui.getLayout(h.ui.getByRole('cell', { name: '42' }));
    h.ui.fireEvent.pointerMove(cell.x + 10, cell.y + 5);
    await settle();
    expect(shown('B2')).not.toBeNull();
  });

  it('is not shown over a cell being typed into', async () => {
    await mount(d => d.setNote(0, 0, 'Hidden while typing'));
    await press('F2');
    expect(shown('A1')).toBeNull();
  });

  it('is on the cell’s right-click menu', async () => {
    await mount();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('cell', { name: 'Region' }));
    await settle();
    expect(h.ui.getByRole('menuitem', { name: 'Note…' })).toBeDefined();
  });
});
