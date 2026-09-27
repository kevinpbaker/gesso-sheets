import { Worker } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';
import type { UiKeyModifiers, UiNode } from 'gesso-core';

import type { ScriptWorker } from '../script/ScriptHost';
import { TEMPLATE } from './ScriptDialog';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Data ▸ Scripts — Phase 30's editor, through the render tree, with
 * the real script worker behind it on a Node thread.
 */

function spawn(): ScriptWorker {
  const worker = new Worker(new URL('../script/nodeWorker.mjs', import.meta.url));
  return {
    post: message => worker.postMessage(message),
    onMessage: listener => worker.on('message', listener),
    terminate: () => void worker.terminate()
  };
}

describe('the script editor', () => {
  let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument };

  async function mount(fill?: (document: SheetDocument) => void): Promise<void> {
    const document = new SheetDocument();
    document.setCell(0, 0, 'Price');
    document.setCell(1, 0, '10');
    document.setCell(2, 0, '20');
    fill?.(document);
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10, scripts: spawn });
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

  /** Settles until `done` holds; a run finishes on another thread, so a few settles may pass first. */
  async function until(done: () => boolean): Promise<void> {
    for (let tries = 0; tries < 200 && !done(); tries++) {
      await new Promise(resolve => setTimeout(resolve, 10));
      await settle();
    }
    expect(done()).toBe(true);
  }

  /** How many nodes say this; none is zero rather than a throw. */
  function shown(text: string | RegExp): UiNode[] {
    try {
      return h.ui.getAllByText(text);
    } catch {
      return [];
    }
  }

  async function openScripts(): Promise<void> {
    await press('F10');
    await press('d');
    h.ui.fireEvent.click(h.ui.getByRole('menuitem', { name: 'Scripts…' }));
    await settle();
  }

  async function replaceSource(text: string): Promise<void> {
    const field = h.ui.getByRole('textbox', { name: 'Script' });
    h.ui.fireEvent.focus(field);
    await press('a', { ctrl: true });
    h.ui.fireEvent.type(text);
    await settle();
  }

  it('opens from Data ▸ Scripts with a script to start from, inside the dialog', async () => {
    await mount();
    await openScripts();
    const field = h.ui.getByRole('textbox', { name: 'Script' });
    expect(h.ui.runtime.input.focus.focusedNode).toBe(field);
    expect(field).toHaveText(TEMPLATE);
    // A long line wraps; it does not push the field past the dialog's edge.
    const dialog = h.ui.getLayout(h.ui.getByRole('dialog', { name: 'Scripts' }));
    const box = h.ui.getLayout(field);
    expect(box.x + box.width).toBeLessThanOrEqual(dialog.x + dialog.width);
  });

  it('runs with Ctrl+Enter, says what it did, and is undone in one step', async () => {
    await mount();
    await openScripts();
    await replaceSource('const v = sheet.range("A2:A3").values;\nsheet.range("B2:B3").write(v.map(([n]) => [n + 1]));\nsheet.range("B2:B3").format({ bold: true });');
    await press('Enter', { ctrl: true });
    await until(() => h.document.sheet.input(2, 1) === '21');

    expect(h.document.sheet.input(1, 1)).toBe('11');
    expect(h.document.formatAt(1, 1).paint.bold).toBe(true);
    // Said in the editor, and in the status bar for when the editor is closed.
    await until(() => shown('Script 1 changed 4 cells.').length === 2);

    await press('Escape');
    expect(h.ui.queryByRole('dialog', { name: 'Scripts' })).toBeNull();
    await press('z', { ctrl: true });
    expect(h.document.sheet.input(1, 1)).toBe('');
    expect(h.document.formatAt(1, 1).paint.bold).toBe(false);
  });

  it('shows why a script failed, and what it logged', async () => {
    await mount();
    await openScripts();
    await replaceSource('console.log("looking"); workbook.sheet("Nowhere");');
    await press('Enter', { ctrl: true });
    await until(() => shown(/stopped with an error/).length > 0);
    expect(shown(/no sheet called "Nowhere"/).length).toBeGreaterThan(0);
    expect(h.ui.queryByText('looking')).not.toBeNull();
  });

  it('asks before running a script that came with a file, each time', async () => {
    await mount(document => {
      document.scripts = [{ name: 'Tidy', source: 'sheet.write("C1", "tidied")', origin: { kind: 'file', file: 'budget.gsheet' } }];
    });
    await openScripts();
    expect(h.ui.queryByText(/Came with budget\.gsheet/)).not.toBeNull();

    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Run' }));
    await settle();
    expect(h.ui.queryByText(/This script came with budget\.gsheet/)).not.toBeNull();
    expect(h.document.sheet.input(0, 2)).toBe('');

    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Run it' }));
    await until(() => h.document.sheet.input(0, 2) === 'tidied');

    // And again next time: saying yes once is not saying yes for good.
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Run' }));
    await settle();
    expect(h.ui.queryByRole('button', { name: 'Run it' })).not.toBeNull();
  });
});
