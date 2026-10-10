import { describe, expect, it } from 'vitest';

import type { SheetDocumentView, SheetLibraryView, SheetTransfer, SheetVersionsView } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { FIRST_DOCUMENT, InMemorySheetLibrary } from './SheetLibrary';
import { SheetService, type Schedule } from './SheetService';
import { VERSION_EVERY_MS } from './SheetVersions';

/**
 * Phase 38: where a person's work is, and what became of it.
 *
 * The save state the title bar reads, the library the home screen
 * lists and changes, *Keep saving to this file*, and versions — each
 * through the service, over an in-memory library, with a clock the
 * spec moves.
 */

function tab(library: InMemorySheetLibrary, seed?: (document: SheetDocument) => void) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const clock = { now: 1_000_000 };
  const service = new SheetService(new SheetDocument(), {
    schedule,
    library,
    seed,
    now: () => clock.now,
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
  const listed = (): SheetLibraryView => latest(service.libraryView);
  const versions = (): SheetVersionsView => latest(service.versionsView);
  const cell = (row: number, column: number): string | undefined => {
    service.setViewport(latest(service.sheets).active, 0, 20, 0, 5);
    drain();
    return latest(service.window).cells[row]?.[column];
  };
  return { service, drain, view, transfer, listed, versions, cell, clock };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

async function opened(library = new InMemorySheetLibrary()) {
  const one = tab(library);
  one.service.openDocument('');
  await one.service.settled;
  return one;
}

describe('the save state', () => {
  it('says saved once an edit is in the repository', async () => {
    const one = await opened();
    one.service.setCell(0, 0, 'kept');
    expect(one.view().saving).toBe('saved');
  });

  it('says failed when the repository refuses a save', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    library.repository(FIRST_DOCUMENT).failing = true;
    one.service.setCell(0, 0, 'lost');
    expect(one.view().saving).toBe('failed');
  });

  it('is off for a document another tab holds', async () => {
    const library = new InMemorySheetLibrary();
    await opened(library);
    const second = tab(library);
    second.service.openDocument(FIRST_DOCUMENT);
    await second.service.settled;
    expect(second.view().saving).toBe('off');
  });
});

describe('the first workbook', () => {
  it('is called what the seed is, and a blank one after it Untitled', async () => {
    const library = new InMemorySheetLibrary();
    const service = new SheetService(new SheetDocument(), {
      library,
      seed: document => document.setCell(0, 0, 'example'),
      seedName: 'Northwind Trading',
      rowCount: 100,
      columnCount: 10
    });
    service.openDocument('');
    await service.settled;
    expect(latest(service.documentView).name).toBe('Northwind Trading');
    service.openDocument('new');
    await service.settled;
    expect(latest(service.documentView).name).toBe('Untitled');
  });
});

describe('renaming', () => {
  it('renames the open document, and the library keeps the name', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.renameDocument(FIRST_DOCUMENT, '  Q3   budget ');
    await one.service.settled;
    expect(one.view().name).toBe('Q3 budget');
    expect((await library.entries())[0].name).toBe('Q3 budget');
  });

  it('turns down an empty name', async () => {
    const one = await opened();
    one.service.renameDocument(FIRST_DOCUMENT, '   ');
    await one.service.settled;
    expect(one.view().name).toBe('Untitled');
  });

  it('suggests the new name for a save as, and keeps it through a save to the old file', async () => {
    const one = await opened();
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;
    one.service.renameDocument(FIRST_DOCUMENT, 'Q3 final');
    await one.service.settled;
    one.service.saveDocument(true);
    expect(one.transfer().download).toMatchObject({ name: 'Q3 final.gsheet', handle: null });
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;
    expect(one.view().name).toBe('Q3 final');
  });
});

describe('the library', () => {
  it('lists every document, most recently edited first, with the open one marked', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.renameDocument(FIRST_DOCUMENT, 'Older');
    one.clock.now += 60_000;
    one.service.openDocument('new');
    await one.service.settled;
    one.service.renameDocument(one.view().id, 'Newer');
    one.service.listDocuments();
    await one.service.settled;
    expect(one.listed().entries.map(entry => [entry.name, entry.open])).toEqual([
      ['Newer', true],
      ['Older', false]
    ]);
  });

  it('copies a document as Copy of, with its cells, without opening it', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.setCell(0, 0, 'original');
    one.service.renameDocument(FIRST_DOCUMENT, 'Budget');
    one.service.duplicateDocument(FIRST_DOCUMENT);
    await one.service.settled;
    const copy = one.listed().entries.find(entry => entry.name === 'Copy of Budget')!;
    expect(copy.open).toBe(false);
    expect(one.view().id).toBe(FIRST_DOCUMENT);
    expect(one.transfer().report).toBe('Made Copy of Budget.');

    one.service.openDocument(copy.id);
    await one.service.settled;
    expect(one.cell(0, 0)).toBe('original');
  });

  it('deletes another document, with its versions', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.openDocument('new');
    await one.service.settled;
    const blank = one.view().id;
    one.service.openDocument(FIRST_DOCUMENT);
    await one.service.settled;
    one.service.deleteDocument(blank);
    await one.service.settled;
    expect((await library.entries()).map(entry => entry.id)).toEqual([FIRST_DOCUMENT]);
    expect(one.transfer().report).toBe('Deleted Untitled.');
  });

  it('moves to the next most recent when the open one is deleted', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.renameDocument(FIRST_DOCUMENT, 'Keep');
    one.clock.now += 1_000;
    one.service.openDocument('new');
    await one.service.settled;
    const doomed = one.view().id;
    one.service.setCell(0, 0, 'going');
    one.service.deleteDocument(doomed);
    await one.service.settled;
    expect(one.view()).toMatchObject({ id: FIRST_DOCUMENT, name: 'Keep' });
    expect(library.repository(doomed).peek()).toBeNull();
  });

  it('opens a blank one when the last is deleted', async () => {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.deleteDocument(FIRST_DOCUMENT);
    await one.service.settled;
    expect(one.view().id).not.toBe(FIRST_DOCUMENT);
    expect((await library.entries()).map(entry => entry.id)).toEqual([one.view().id]);
  });

  it('will not delete a document another tab has open', async () => {
    const library = new InMemorySheetLibrary();
    const holder = await opened(library);
    holder.service.openDocument('new');
    await holder.service.settled;
    const held = holder.view().id;

    const other = tab(library);
    other.service.openDocument(FIRST_DOCUMENT);
    await other.service.settled;
    other.service.deleteDocument(held);
    await other.service.settled;
    expect(other.transfer().report).toBe('Untitled is open in another tab. Close it there first.');
    expect((await library.entries()).some(entry => entry.id === held)).toBe(true);
  });
});

describe('keeping a file up to date', () => {
  it('writes every saved change to the file, quietly', async () => {
    const one = await opened();
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;
    one.service.setAutosave(true);
    await one.service.settled;
    expect(one.view().autosave).toBe(true);
    expect(one.transfer().report).toBe('Every change now goes to Q3.gsheet as well.');

    one.service.setCell(0, 0, '42');
    const download = one.transfer().download!;
    expect(download).toMatchObject({ kind: 'workbook', name: 'Q3.gsheet', handle: 4, quiet: true });
    expect(JSON.parse(download.text).sheets[0].cells).toContainEqual({ row: 0, column: 0, input: '42' });

    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file', true);
    await one.service.settled;
    expect(one.view().edited).toBe(false);
    // Nothing said: it happens on every edit.
    expect(one.transfer().report).toBe('');
  });

  it('asks where first for a document with no file, and keeps that one up to date', async () => {
    const one = await opened();
    one.service.setAutosave(true);
    expect(one.transfer().download).toMatchObject({ name: 'Untitled.gsheet', handle: null });
    one.service.fileSaved('workbook', 'Plan.gsheet', 9, 'file');
    await one.service.settled;
    expect(one.view()).toMatchObject({ autosave: true, name: 'Plan', file: { handle: 9, name: 'Plan.gsheet' } });
    expect(one.transfer().report).toBe('Saved Plan.gsheet. Every change now goes to it as well.');
  });

  it('says so when the browser can only download', async () => {
    const one = await opened();
    one.service.setAutosave(true);
    one.service.fileSaved('workbook', 'Plan.gsheet', null, 'download');
    await one.service.settled;
    expect(one.view().autosave).toBe(false);
    expect(one.transfer().report).toBe(
      'Downloaded Plan.gsheet. This browser cannot write to a file again, so changes stay here until the next download.'
    );
  });

  it('says what to do when a quiet write is refused', async () => {
    const one = await opened();
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    await one.service.settled;
    one.service.fileNotSaved('the browser did not allow it.');
    expect(one.transfer().report).toBe('Q3.gsheet was not kept up to date: the browser did not allow it. Press Ctrl+S to save it.');
  });
});

describe('versions', () => {
  async function withWork() {
    const library = new InMemorySheetLibrary();
    const one = await opened(library);
    one.service.setCell(0, 0, 'monday');
    // Opened again, as next week: the first edit keeps it as it was.
    one.service.openDocument('new');
    await one.service.settled;
    one.clock.now += 1_000;
    one.service.openDocument(FIRST_DOCUMENT);
    await one.service.settled;
    return { library, one };
  }

  it('keeps nothing for a document only read', async () => {
    const { one } = await withWork();
    one.service.listVersions();
    await one.service.settled;
    // The first session opened a blank, which is not worth keeping.
    expect(one.versions().entries).toEqual([]);
  });

  it('keeps the document as it was opened at the first edit, then one every ten minutes', async () => {
    const { one } = await withWork();
    const openedAt = one.clock.now;
    one.service.setCell(0, 0, 'tuesday');
    one.clock.now += VERSION_EVERY_MS - 1;
    one.service.setCell(0, 0, 'wednesday');
    one.clock.now += 1;
    one.service.setCell(0, 0, 'thursday');
    one.service.listVersions();
    await one.service.settled;
    const entries = one.versions().entries;
    expect(entries.map(version => version.reason)).toEqual(['auto', 'opened']);
    expect(entries[1].at).toBe(openedAt);
  });

  it('puts a version back, and keeps what it replaced', async () => {
    const { one } = await withWork();
    one.service.setCell(0, 0, 'tuesday');
    one.service.listVersions();
    await one.service.settled;
    const monday = one.versions().entries[0];

    one.clock.now += 5_000;
    one.service.restoreVersion(monday.at);
    await one.service.settled;
    expect(one.cell(0, 0)).toBe('monday');
    expect(one.versions().entries.map(version => version.reason)).toEqual(['restored', 'opened']);
    expect(one.transfer().report).toBe('Restored an earlier version. What it replaced is in the list as Before a restore.');

    one.service.restoreVersion(one.versions().entries[0].at);
    await one.service.settled;
    expect(one.cell(0, 0)).toBe('tuesday');
  });

  it('keeps one for every save to a file', async () => {
    const { one } = await withWork();
    one.service.fileSaved('workbook', 'Q3.gsheet', 4, 'file');
    one.service.listVersions();
    await one.service.settled;
    expect(one.versions().entries[0].reason).toBe('saved');
  });
});
