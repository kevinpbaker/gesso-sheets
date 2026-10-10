import { afterEach, describe, expect, it } from 'vitest';

import { createComponent } from 'gesso-framework';
import { renderTest, serveForTest, textProperty, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import type { UiKeyModifiers } from 'gesso-core';

import type { SheetDocumentView } from './SheetContract';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { FIRST_DOCUMENT, InMemorySheetLibrary } from './SheetLibrary';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';
import { matching, saveSaid, whenSaid } from './whereItIs';

/**
 * Phase 36 on the screen: the name and the save state at the top, the
 * home screen, and version history — through the render tree, with the
 * keyboard and the semantics a screen reader would use.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  library: InMemorySheetLibrary;
}

let h: Harness | undefined;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
  h = undefined;
});

async function settle(): Promise<void> {
  for (let round = 0; round < 3; round++) {
    await h!.ui.settle();
    await h!.served.settle();
    await h!.service.settled;
  }
}

async function mount(): Promise<Harness> {
  const library = new InMemorySheetLibrary();
  const service = new SheetService(new SheetDocument(), { library, rowCount: 200, columnCount: 20 });
  service.openDocument('');
  await service.settled;
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1400, height: 700 });
  h = { ui, served, service, library };
  await settle();
  return h;
}

async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
  h!.ui.fireEvent.press(key, modifiers);
  await settle();
}

async function menu(mnemonic: string, item: string): Promise<void> {
  await openMenu(mnemonic);
  h!.ui.fireEvent.click(h!.ui.getByRole('menuitem', { name: item }));
  await settle();
}

async function openMenu(mnemonic: string): Promise<void> {
  const grid = h!.ui.getByRole('grid');
  let stops = 0;
  while (h!.ui.runtime.input.focus.focusedNode !== grid) {
    if (stops++ > 12) {
      throw new Error('Tab never reached the grid');
    }
    h!.ui.fireEvent.tab();
    await h!.ui.settle();
  }
  await press('F10');
  await press(mnemonic);
}

async function click(name: string): Promise<void> {
  h!.ui.fireEvent.click(h!.ui.getByRole('button', { name }));
  await settle();
}

const said = (): string => textProperty(h!.ui.getByRole('status', { name: 'Where this workbook is saved' })) ?? '';
const title = (): string => textProperty(h!.ui.getByRole('textbox', { name: 'Workbook name' })) ?? '';

describe('the title bar', () => {
  it('shows the workbook’s name, and where it is kept', async () => {
    await mount();
    expect(title()).toBe('Untitled');
    expect(said()).toBe('Saved in this browser');
  });

  it('is renamed in place from File ▸ Rename, with Enter', async () => {
    await mount();
    await menu('f', 'Rename…');
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(h!.ui.getByRole('textbox', { name: 'Workbook name' }));
    h!.ui.fireEvent.type('Q3 budget');
    await press('Enter');
    expect(title()).toBe('Q3 budget');
    expect((await h!.library.entries())[0].name).toBe('Q3 budget');
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(h!.ui.getByRole('grid'));
  });

  it('puts the name back on Escape', async () => {
    await mount();
    await menu('f', 'Rename…');
    h!.ui.fireEvent.type('Not this');
    await press('Escape');
    expect(title()).toBe('Untitled');
    expect((await h!.library.entries())[0].name).toBe('Untitled');
  });

  it('says so when the browser will not keep an edit', async () => {
    await mount();
    h!.library.repository(FIRST_DOCUMENT).failing = true;
    h!.service.setCell(0, 0, 'lost');
    await settle();
    expect(said()).toBe('Not saved: this browser refused to store it');
  });
});

describe('where the work is, in words', () => {
  const base: SheetDocumentView = {
    id: 'd1',
    name: 'Q3',
    file: null,
    edited: false,
    elsewhere: false,
    saving: 'saved',
    autosave: false
  };

  it('names the browser and the file, and what is behind', () => {
    expect(saveSaid(base)).toEqual({ text: 'Saved in this browser', tone: 'muted' });
    expect(saveSaid({ ...base, saving: 'saving' }).text).toBe('Saving…');
    const file = { handle: 4, name: 'Q3.gsheet' };
    expect(saveSaid({ ...base, file }).text).toBe('Saved in this browser and to Q3.gsheet');
    expect(saveSaid({ ...base, file, edited: true }).text).toBe('Saved in this browser · Q3.gsheet is behind');
    expect(saveSaid({ ...base, file, autosave: true }).text).toBe('Saved to Q3.gsheet');
    expect(saveSaid({ ...base, file, autosave: true, edited: true }).text).toBe('Saving to Q3.gsheet…');
    expect(saveSaid({ ...base, file: { handle: null, name: 'Q3.gsheet' } }).text).toBe('Saved in this browser · downloaded as Q3.gsheet');
  });

  it('says it in red when the work is not being kept', () => {
    expect(saveSaid({ ...base, elsewhere: true })).toEqual({ text: 'Open in another tab, so not saved here', tone: 'danger' });
    expect(saveSaid({ ...base, saving: 'memory' }).tone).toBe('danger');
    expect(saveSaid({ ...base, saving: 'failed' }).tone).toBe('danger');
  });

  it('says when by distance, then by the calendar', () => {
    const now = new Date(2026, 9, 10, 15, 30).getTime();
    expect(whenSaid(now - 20_000, now)).toBe('Just now');
    expect(whenSaid(now - 60_000, now)).toBe('1 minute ago');
    expect(whenSaid(now - 25 * 60_000, now)).toBe('25 minutes ago');
    expect(whenSaid(new Date(2026, 9, 10, 9, 5).getTime(), now)).toBe('Today, 09:05');
    expect(whenSaid(new Date(2026, 9, 9, 18, 0).getTime(), now)).toBe('Yesterday, 18:00');
    expect(whenSaid(new Date(2026, 6, 2, 8, 0).getTime(), now)).toBe('2 Jul, 08:00');
    expect(whenSaid(new Date(2025, 11, 24, 8, 0).getTime(), now)).toBe('24 Dec 2025');
  });

  it('finds a workbook by every word of its name, case aside', () => {
    const entries = [{ name: 'Q3 Budget' }, { name: 'Holiday plan' }, { name: 'Budget 2027' }];
    expect(matching(entries, 'budget').map(entry => entry.name)).toEqual(['Q3 Budget', 'Budget 2027']);
    expect(matching(entries, 'BUDGET q3').map(entry => entry.name)).toEqual(['Q3 Budget']);
    expect(matching(entries, '  ')).toHaveLength(3);
  });
});

describe('the home screen', () => {
  async function withTwo(): Promise<void> {
    await mount();
    h!.service.renameDocument(FIRST_DOCUMENT, 'Budget');
    h!.service.openDocument('new');
    await settle();
    h!.service.renameDocument((await h!.library.entries()).find(each => each.id !== FIRST_DOCUMENT)!.id, 'Holiday plan');
    await settle();
  }

  it('lists every workbook, and opens on the search field', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    expect(h!.ui.getByRole('dialog', { name: 'Workbooks' })).toBeDefined();
    expect(h!.ui.getByRole('button', { name: 'Open Budget' })).toBeDefined();
    expect(h!.ui.getByRole('button', { name: 'Open Holiday plan' })).toBeDefined();
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(h!.ui.getByRole('searchbox', { name: 'Search workbooks' }));
  });

  it('narrows the list as the search is typed', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    h!.ui.fireEvent.type('holi');
    await settle();
    expect(h!.ui.queryByRole('button', { name: 'Open Budget' })).toBeNull();
    expect(h!.ui.getByRole('button', { name: 'Open Holiday plan' })).toBeDefined();
  });

  it('opens one, and closes', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    await click('Open Budget');
    expect(h!.ui.queryByRole('dialog', { name: 'Workbooks' })).toBeNull();
    expect(title()).toBe('Budget');
  });

  it('renames one in its row', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    await click('Rename Budget');
    const field = h!.ui.getByRole('textbox', { name: 'New name for Budget' });
    expect(h!.ui.runtime.input.focus.focusedNode).toBe(field);
    await press('a', { ctrl: true });
    h!.ui.fireEvent.type('Budget 2027');
    await press('Enter');
    expect(h!.ui.getByRole('button', { name: 'Open Budget 2027' })).toBeDefined();
    expect((await h!.library.entries()).find(each => each.id === FIRST_DOCUMENT)!.name).toBe('Budget 2027');
  });

  it('duplicates one', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    await click('Duplicate Budget');
    expect(h!.ui.getByRole('button', { name: 'Open Copy of Budget' })).toBeDefined();
  });

  it('asks before deleting, and keeps it when told to', async () => {
    await withTwo();
    await menu('f', 'All workbooks…');
    await click('Delete Budget');
    expect(h!.ui.getByText('Delete “Budget”? Its versions go with it, and it cannot be undone.')).toBeDefined();
    await click('Keep it');
    expect(h!.ui.getByRole('button', { name: 'Open Budget' })).toBeDefined();
    expect(await h!.library.entries()).toHaveLength(2);

    await click('Delete Budget');
    h!.ui.fireEvent.click(h!.ui.getAllByRole('button', { name: 'Delete Budget' })[0]);
    await settle();
    expect(h!.ui.queryByRole('button', { name: 'Open Budget' })).toBeNull();
    expect((await h!.library.entries()).map(each => each.name)).toEqual(['Holiday plan']);
  });

  it('starts a blank workbook', async () => {
    await mount();
    await menu('f', 'All workbooks…');
    await click('New blank workbook');
    expect(h!.ui.queryByRole('dialog', { name: 'Workbooks' })).toBeNull();
    expect(await h!.library.entries()).toHaveLength(2);
  });

  it('closes on Escape', async () => {
    await mount();
    await menu('f', 'All workbooks…');
    await press('Escape');
    expect(h!.ui.queryByRole('dialog', { name: 'Workbooks' })).toBeNull();
  });
});

describe('version history', () => {
  it('lists the versions kept, and restores one', async () => {
    await mount();
    h!.service.setCell(0, 0, 'first');
    h!.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await settle();
    h!.service.setCell(0, 0, 'second');
    await settle();

    await menu('f', 'Version history…');
    expect(h!.ui.getByRole('dialog', { name: 'Version history' })).toBeDefined();
    expect(h!.ui.getByText('Saved to a file')).toBeDefined();
    h!.ui.fireEvent.click(h!.ui.getByRole('button', { name: /^Restore the version from / }));
    await settle();
    expect(h!.ui.queryByRole('dialog', { name: 'Version history' })).toBeNull();
    expect(h!.service.snapshot().sheets[0].cells).toContainEqual({ row: 0, column: 0, input: 'first' });
  });

  it('says how versions come to be kept, when there are none', async () => {
    await mount();
    await menu('f', 'Version history…');
    expect(h!.ui.getByText(/^No versions yet\./)).toBeDefined();
  });
});

describe('keep saving to this file', () => {
  it('is ticked in the File menu while it is on', async () => {
    await mount();
    h!.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await settle();
    h!.service.setAutosave(true);
    await settle();
    expect(said()).toBe('Saved to Q3.gsheet');
    await openMenu('f');
    expect(h!.ui.getByRole('menuitemcheckbox', { name: 'Keep saving to this file' })).toHaveSemantics({ states: ['checked'] });
  });
});
