import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent, performShellStorage, ShellService, type ShellLocalStore } from 'gesso-framework';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * The status bar's figures, chosen — Phase 22.
 *
 * Sum, Average and Count stay the default. The readout offers the rest
 * — Numerical count, Min, Max — and a click on a figure copies it,
 * whole, because the number somebody checked is usually the number
 * they wanted to put somewhere.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  copied: () => string;
}

let h: Harness;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
});

/**
 * The browser's `localStorage`, as the shell would answer for it: kept
 * across mounts, which is what a reload is from here.
 */
function memoryStore(): ShellLocalStore {
  const held = new Map<string, string>();
  return {
    get length() {
      return held.size;
    },
    key: at => [...held.keys()][at] ?? null,
    getItem: key => held.get(key) ?? null,
    setItem: (key, value) => void held.set(key, value),
    removeItem: key => void held.delete(key)
  };
}

async function mount(store?: ShellLocalStore): Promise<void> {
  const document = new SheetDocument();
  document.setCell(0, 0, '1');
  document.setCell(1, 0, '2');
  document.setCell(2, 0, '3.14159');
  document.setCell(3, 0, 'text');
  document.setCell(0, 1, 'a');
  document.setCell(1, 1, 'b');
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
  const served = serveForTest([sheetChannel(service)]);
  const writes: string[] = [];
  // The shell's answers, in place before the first frame as a real
  // shell's are: the status bar asks for its stored figures as it
  // mounts, and a handler installed after that would miss the question.
  const ui = renderTest(createComponent(SheetApp), {
    channels: served.registry,
    width: 900,
    height: 420,
    onCreate: runtime => {
      const shell = runtime.services.get(ShellService);
      runtime.onShellRequest(request => {
        if (request.type === 'clipboard') {
          writes.push(request.text);
        }
        if (request.type === 'storage') {
          shell.settleStorage(request.id, performShellStorage(request, () => store ?? null));
        }
      });
    }
  });
  h = { ui, served, service, copied: () => writes[writes.length - 1] ?? '' };
  await settle();
}

async function settle(): Promise<void> {
  await h.ui.settle();
  await h.served.settle();
  await h.ui.settle();
}

async function select(firstRow: number, firstColumn: number, lastRow: number, lastColumn: number): Promise<void> {
  h.service.setSelection(firstRow, firstColumn, lastRow, lastColumn);
  await settle();
}

const figure = (name: string) => h.ui.queryByRole('button', { name });

async function choose(label: string): Promise<void> {
  h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Choose what the status bar shows' }));
  await settle();
  const item = h.ui.getAllByRole('menuitem').find(node => String(node.properties.get('label')).trim().endsWith(label));
  if (item === undefined) {
    throw new Error(`no menu item for ${label}`);
  }
  h.ui.fireEvent.click(item);
  await settle();
}

describe('the status bar’s figures', () => {
  it('are Sum, Average and Count until somebody chooses otherwise', async () => {
    await mount();
    await select(0, 0, 3, 0);

    expect(figure('Sum 6.1416')).not.toBeNull();
    expect(figure('Average 2.0472')).not.toBeNull();
    expect(figure('Count 4')).not.toBeNull();
    expect(figure('Min 1')).toBeNull();
  });

  it('take Min, Max and Numerical count from the menu, and let Sum go', async () => {
    await mount();
    await select(0, 0, 3, 0);
    await choose('Min');
    await choose('Max');
    await choose('Numerical count');
    await choose('Sum');

    expect(figure('Min 1')).not.toBeNull();
    expect(figure('Max 3.1416')).not.toBeNull();
    expect(figure('Numerical count 3')).not.toBeNull();
    expect(figure('Sum 6.1416')).toBeNull();
  });

  it('mark what is chosen in the menu', async () => {
    await mount();
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Choose what the status bar shows' }));
    await settle();
    const labels = h.ui.getAllByRole('menuitem').map(node => String(node.properties.get('label')));

    expect(labels.find(label => label.endsWith('Sum'))?.startsWith('✓')).toBe(true);
    expect(labels.find(label => label.endsWith('Min'))?.startsWith('✓')).toBe(false);
  });

  it('keep the choice as the selection moves', async () => {
    await mount();
    await choose('Max');
    await select(0, 0, 1, 0);
    expect(figure('Max 2')).not.toBeNull();
  });

  it('say only how many over a selection with no numbers in it', async () => {
    await mount();
    await choose('Min');
    await select(0, 1, 1, 1);

    expect(figure('Count 2')).not.toBeNull();
    expect(figure('Min 0')).toBeNull();
    expect(figure('Sum 0')).toBeNull();
  });

  it('copy a figure, whole, on a click', async () => {
    await mount();
    await select(0, 0, 3, 0);
    h.ui.fireEvent.click(figure('Sum 6.1416')!);
    await settle();

    expect(h.copied()).toBe('6.14159');
    expect(h.ui.queryByText('Sum copied')).not.toBeNull();
  });

  /** Phase 27's exit: chosen, reloaded, and still chosen. */
  it('are kept for the next visit', async () => {
    const store = memoryStore();
    await mount(store);
    await choose('Min');
    // A choice is written once it has settled, not on every click.
    await new Promise(resolve => setTimeout(resolve, 350));
    h.ui.unmount();
    h.served.dispose();

    await mount(store);
    await select(0, 0, 3, 0);
    expect(figure('Min 1')).not.toBeNull();
    expect(figure('Sum 6.1416')).not.toBeNull();
  });

  it('come back as the default when what was stored is not a choice', async () => {
    const store = memoryStore();
    store.setItem('gessosheet:status-figures', JSON.stringify(['sum', 'nonsense']));
    await mount(store);
    await select(0, 0, 3, 0);
    expect(figure('Sum 6.1416')).not.toBeNull();
    expect(figure('Count 4')).not.toBeNull();
  });
});
