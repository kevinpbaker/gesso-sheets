import { describe, expect, it } from 'vitest';

import type { SheetDocumentView, SheetTransfer } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { FIRST_DOCUMENT, InMemorySheetLibrary } from './SheetLibrary';
import { SheetService, type Schedule } from './SheetService';

/**
 * Documents: which workbook a tab shows, and the file it belongs to.
 *
 * A tab is a document. Two tabs over one library are two services over
 * the same shelf, which is what a reload is too, so that is what these
 * build: a library, and as many services over it as the story needs.
 */

function tab(library: InMemorySheetLibrary, seed?: (document: SheetDocument) => void) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  let clock = 1_000;
  const service = new SheetService(new SheetDocument(), {
    schedule,
    library,
    seed,
    now: () => clock++,
    rowCount: 100,
    columnCount: 10
  });
  const drain = () => {
    while (queue.length > 0) {
      queue.shift()!();
    }
  };
  const view = (): SheetDocumentView => latest(service.documentView);
  const transfer = (): SheetTransfer => latest(service.transfer);
  /** Whatever the service shows, read through the window it publishes. */
  const cell = (row: number, column: number): string | undefined => {
    service.setViewport(latest(service.sheets).active, 0, 20, 0, 5);
    drain();
    return latest(service.window).cells[row]?.[column];
  };
  return { service, drain, view, transfer, cell };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

describe('the first document', () => {
  it('is seeded, and kept where every earlier build kept the workbook', async () => {
    const library = new InMemorySheetLibrary();
    const first = tab(library, document => document.setCell(0, 0, 'seeded'));
    first.service.openDocument('');
    await first.service.settled;

    expect(first.view()).toEqual({ id: FIRST_DOCUMENT, name: 'Untitled', file: null, edited: false, elsewhere: false });
    expect(first.cell(0, 0)).toBe('seeded');
    expect(library.repository(FIRST_DOCUMENT).peek()).not.toBeNull();
  });

  it('opens a workbook an older build left there, rather than seeding over it', async () => {
    const library = new InMemorySheetLibrary();
    const old = tab(library);
    old.service.openDocument('');
    await old.service.settled;
    old.service.setCell(0, 0, 'from Phase 6');
    await old.service.flush();
    // Forget the index, as a profile from before this phase has none.
    const bare = new InMemorySheetLibrary();
    bare.repository(FIRST_DOCUMENT).save(library.repository(FIRST_DOCUMENT).peek()!);

    const now = tab(bare, document => document.setCell(0, 0, 'seeded'));
    now.service.openDocument('');
    await now.service.settled;
    expect(now.cell(0, 0)).toBe('from Phase 6');
  });
});

describe('opening documents', () => {
  it('opens a blank one on new, and comes back to the last used on a reload', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library, document => document.setCell(0, 0, 'seeded'));
    one.service.openDocument('');
    await one.service.settled;

    one.service.openDocument('new');
    await one.service.settled;
    const blank = one.view().id;
    expect(blank).not.toBe(FIRST_DOCUMENT);
    expect(one.cell(0, 0) ?? '').toBe('');
    one.service.setCell(2, 2, 'in the new one');
    await one.service.flush();

    const reload = tab(library);
    reload.service.openDocument('');
    await reload.service.settled;
    expect(reload.view().id).toBe(blank);
    expect(reload.cell(2, 2)).toBe('in the new one');
  });

  it('opens one by id, and leaves the other as it was', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.setCell(0, 0, 'first');
    one.service.openDocument('new');
    await one.service.settled;
    one.service.setCell(0, 0, 'second');

    one.service.openDocument(FIRST_DOCUMENT);
    await one.service.settled;
    expect(one.cell(0, 0)).toBe('first');
  });

  it('falls back to the last used for an id it does not know', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.openDocument('nobody-has-this');
    await one.service.settled;
    expect(one.view().id).toBe(FIRST_DOCUMENT);
  });

  it('forgets what it knew about the last workbook', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.setCell(0, 0, 'needle');
    one.service.find('needle', false, false, false);
    one.service.openDocument('new');
    await one.service.settled;
    expect(latest(one.service.findView).matches).toBe(0);
    expect(one.service.snapshot().sheets[0].cells).toEqual([]);
  });

  it('is marked edited by the first change, and not before', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    expect(one.view().edited).toBe(false);
    one.service.setCell(0, 0, 'x');
    expect(one.view().edited).toBe(true);
  });
});

describe('one document in two tabs', () => {
  /** Two copies kept at once would each write over the other's edits, and the last would win. */
  it('is kept by the tab that opened it first, and only shown by the second', async () => {
    const library = new InMemorySheetLibrary();
    const first = tab(library);
    first.service.openDocument('');
    await first.service.settled;
    first.service.setCell(0, 0, 'first tab');
    await first.service.flush();

    const second = tab(library);
    second.service.openDocument(FIRST_DOCUMENT);
    await second.service.settled;
    expect(second.view().elsewhere).toBe(true);
    expect(second.cell(0, 0)).toBe('first tab');
    expect(second.transfer().report).toBe('Close the other tab and reload this one to edit here.');

    second.service.setCell(0, 0, 'second tab');
    await second.service.flush();
    expect(library.repository(FIRST_DOCUMENT).peek()!.sheets[0].cells).toContainEqual({ row: 0, column: 0, input: 'first tab' });
  });

  it('is free again once the first tab moves to another document', async () => {
    const library = new InMemorySheetLibrary();
    const first = tab(library);
    first.service.openDocument('');
    await first.service.settled;
    first.service.openDocument('new');
    await first.service.settled;

    const second = tab(library);
    second.service.openDocument(FIRST_DOCUMENT);
    await second.service.settled;
    expect(second.view().elsewhere).toBe(false);
  });
});

describe('a document that cannot be read', () => {
  /**
   * The data-loss path, closed: a read that fails is not an empty
   * document, and an empty one must never be saved over it. It failed
   * in a browser whenever two tabs touched the file at once.
   */
  it('is left as it was, with nothing saved over it', async () => {
    const library = new InMemorySheetLibrary();
    const keeper = tab(library);
    keeper.service.openDocument('');
    await keeper.service.settled;
    keeper.service.setCell(0, 0, 'precious');
    keeper.service.openDocument('new');
    await keeper.service.settled;
    await keeper.service.flush();
    const repository = library.repository(FIRST_DOCUMENT);
    const before = repository.peek();
    repository.load = () => Promise.reject(new Error('held by another tab'));
    const saves = repository.saves;

    const reader = tab(library);
    reader.service.openDocument(FIRST_DOCUMENT);
    await reader.service.settled;
    expect(reader.transfer().report).toBe('Untitled could not be read, so it was left as it was. Reload to try again.');
    reader.service.setCell(5, 5, 'typed into the blank');
    await reader.service.flush();

    expect(repository.saves).toBe(saves);
    expect(repository.peek()).toBe(before);
  });
});

describe('saving a workbook to a file', () => {
  it('builds a .gsheet to save as, and remembers the file it went to', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.setCell(0, 0, '6');
    one.service.setCell(0, 1, '=A1*7');

    one.service.saveDocument(false);
    const download = one.transfer().download!;
    expect(download).toMatchObject({ kind: 'workbook', name: 'Untitled.gsheet', mediaType: 'application/json', handle: null });
    expect(JSON.parse(download.text).sheets[0].cells).toContainEqual({ row: 0, column: 1, input: '=A1*7' });

    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;
    expect(one.view()).toMatchObject({ name: 'Q3', file: { handle: 4, name: 'Q3.gsheet' }, edited: false });
    expect(one.transfer().report).toBe('Saved Q3.gsheet.');
  });

  it('saves back to that file next time, and asks again for a save as', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;

    one.service.saveDocument(false);
    expect(one.transfer().download).toMatchObject({ name: 'Q3.gsheet', handle: 4 });
    one.service.saveDocument(true);
    expect(one.transfer().download).toMatchObject({ name: 'Q3.gsheet', handle: null });
  });

  it('keeps no handle for a download, since there is nothing to save back to', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.fileSaved('workbook', 'Q3.gsheet', null, 'download');
    await one.service.settled;
    one.service.saveDocument(false);
    expect(one.transfer().download!.handle).toBeNull();
    expect(one.transfer().report).toBe('Downloaded Q3.gsheet.');
  });
});

describe('opening a file', () => {
  /** The exit criterion, in node: a CSV in, formulas added, saved, closed, reopened. */
  it('brings a saved workbook back, and recalculates the formulas added since the CSV', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.importCsv('orders.csv', 'Units,Price\n3,4\n5,6\n');
    one.service.setCell(1, 2, '=A2*B2');
    one.service.setCell(2, 2, '=A3*B3');
    one.drain();
    one.service.saveDocument(false);
    const saved = one.transfer().download!.text;

    // A different browser profile, with none of this in it: only the file.
    const elsewhere = tab(new InMemorySheetLibrary());
    elsewhere.service.openDocument('');
    await elsewhere.service.settled;
    elsewhere.service.openFile('orders.gsheet', saved, 9);
    await elsewhere.service.settled;
    elsewhere.drain();

    expect(elsewhere.view()).toMatchObject({ name: 'orders', file: { handle: 9, name: 'orders.gsheet' }, edited: false });
    expect(elsewhere.transfer().report).toBe('Opened orders.gsheet.');
    elsewhere.service.activateSheet(1);
    expect(elsewhere.cell(1, 2)).toBe('12');
    expect(elsewhere.cell(2, 2)).toBe('30');
  });

  it('shows the same document again for the same file, with what the file holds', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.setCell(0, 0, 'on disk');
    one.service.saveDocument(false);
    const text = one.transfer().download!.text;

    one.service.openFile('a.gsheet', text, 3);
    await one.service.settled;
    const first = one.view().id;
    one.service.setCell(0, 0, 'edited here, never saved');
    one.service.openDocument('new');
    await one.service.settled;

    one.service.openFile('a.gsheet', text, 3);
    await one.service.settled;
    expect(one.view().id).toBe(first);
    expect(one.cell(0, 0)).toBe('on disk');
  });

  it('turns down a file that is not a workbook, and stays where it was', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.openFile('broken.gsheet', '{not json', null);
    await one.service.settled;
    expect(one.view().id).toBe(FIRST_DOCUMENT);
    expect(one.transfer().report).toBe('broken.gsheet was not opened: it is not a workbook this can read.');
  });

  it('sends a CSV to the importer', async () => {
    const library = new InMemorySheetLibrary();
    const one = tab(library);
    one.service.openDocument('');
    await one.service.settled;
    one.service.openFile('a.csv', 'x\n', 5);
    expect(latest(one.service.sheets).entries.map(sheet => sheet.name)).toEqual(['Sheet1', 'a']);
  });
});
