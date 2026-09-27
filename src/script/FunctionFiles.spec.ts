import { describe, expect, it } from 'vitest';

import type { SheetEditor, SheetScripts, SheetTransfer } from '../app/SheetContract';
import { SheetDocument } from '../app/SheetDocument';
import { InMemorySheetLibrary } from '../app/SheetLibrary';
import { SheetService, type Schedule } from '../app/SheetService';
import { xlsxOfDocument } from '../app/SheetXlsxOut';
import { xlsxParts } from '../sheet/XlsxWrite';

/**
 * Functions from files — Phase 33.
 *
 * A file's functions are off until somebody turns them on for that
 * workbook; turned on is kept in this browser's library as what the
 * code said, never in the file.
 */

function tab(library = new InMemorySheetLibrary()) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const service = new SheetService(new SheetDocument(), { library, schedule });
  let scripts: SheetScripts = { entries: [], running: '', refused: '', last: null };
  let transfer: SheetTransfer = { download: null, report: '' };
  let editor: SheetEditor | null = null;
  service.scripts.subscribe(next => (scripts = next));
  service.transfer.subscribe(next => (transfer = next));
  service.editor.subscribe(next => (editor = next));
  const settle = async () => {
    await service.settled;
    await service.functionsReady;
    for (let turns = 0; queue.length > 0 && turns < 100_000; turns++) {
      queue.shift()!();
    }
  };
  /** What a cell shows, read from what would be saved: the document changes as files open. */
  const shown = (row: number, column: number): unknown =>
    (service as unknown as { document: SheetDocument }).document.sheet.value(row, column);
  return { service, library, settle, shown, scripts: () => scripts, transfer: () => transfer, editor: () => editor };
}

/** A workbook saved to a file by somebody else: a function, and a column that calls it. */
async function aFile(source = 'function TAX(amount, rate) { return amount * rate; }'): Promise<string> {
  const author = tab();
  author.service.openDocument('');
  await author.settle();
  author.service.saveScript('', 'Taxes', source, 'functions');
  author.service.setCell(0, 0, '200');
  author.service.setCell(0, 1, '=TAX(A1, 0.2)');
  author.service.setCell(1, 1, 'kept');
  await author.settle();
  author.service.saveDocument(false);
  return author.transfer().download!.text;
}

describe('a file with functions', () => {
  it('opens with them off, and the sheet otherwise working', async () => {
    const text = await aFile();
    const reader = tab();
    reader.service.openDocument('');
    await reader.settle();
    reader.service.openFile('budget.gsheet', text, 7);
    await reader.settle();

    expect(reader.shown(0, 1)).toMatchObject({ code: '#NAME?' });
    expect(reader.shown(1, 1)).toBe('kept');
    expect(reader.scripts().entries[0]).toMatchObject({ kind: 'functions', from: 'budget.gsheet', on: false });
    // And the cell says why.
    reader.service.setSelection(0, 1, 0, 1);
    expect(reader.editor()?.explain?.meaning).toContain('TAX came with budget.gsheet');
  });

  it('turned on, gives the values it was saved with; reopened, is still on', async () => {
    const text = await aFile();
    const library = new InMemorySheetLibrary();
    const reader = tab(library);
    reader.service.openDocument('');
    await reader.settle();
    reader.service.openFile('budget.gsheet', text, 7);
    await reader.settle();

    reader.service.setFunctionsOn(true);
    await reader.settle();
    expect(reader.shown(0, 1)).toBe(40);
    expect(reader.scripts().entries[0].on).toBe(true);
    // Kept in the library, as fingerprints, and not in the file.
    const [entry] = (await library.entries()).filter(each => each.file?.handle === 7);
    expect(entry.trustedFunctions).toHaveLength(1);
    expect(entry.trustedFunctions![0]).toMatch(/^[0-9a-f]{64}$/);
    reader.service.saveDocument(false);
    expect(reader.transfer().download!.text).not.toContain(entry.trustedFunctions![0]);

    // The same file opened again in the same browser.
    reader.service.openDocument('new');
    await reader.settle();
    reader.service.openFile('budget.gsheet', text, 7);
    await reader.settle();
    expect(reader.shown(0, 1)).toBe(40);

    // And off again when asked.
    reader.service.setFunctionsOn(false);
    await reader.settle();
    expect(reader.shown(0, 1)).toMatchObject({ code: '#NAME?' });
  });

  it('is off again in another browser, whose library never said yes', async () => {
    const text = await aFile();
    const here = tab();
    here.service.openDocument('');
    await here.settle();
    here.service.openFile('budget.gsheet', text, 7);
    await here.settle();
    here.service.setFunctionsOn(true);
    await here.settle();
    expect(here.shown(0, 1)).toBe(40);

    const elsewhere = tab(new InMemorySheetLibrary());
    elsewhere.service.openDocument('');
    await elsewhere.settle();
    elsewhere.service.openFile('budget.gsheet', text, 7);
    await elsewhere.settle();
    expect(elsewhere.shown(0, 1)).toMatchObject({ code: '#NAME?' });
  });

  it('that comes back with different code is off again, though it was turned on', async () => {
    const library = new InMemorySheetLibrary();
    const reader = tab(library);
    reader.service.openDocument('');
    await reader.settle();
    reader.service.openFile('budget.gsheet', await aFile(), 7);
    await reader.settle();
    reader.service.setFunctionsOn(true);
    await reader.settle();
    expect(reader.shown(0, 1)).toBe(40);

    // The same file, from the same place, with TAX changed underneath.
    reader.service.openFile('budget.gsheet', await aFile('function TAX(amount, rate) { return 999; }'), 7);
    await reader.settle();
    expect(reader.shown(0, 1)).toMatchObject({ code: '#NAME?' });
    expect(reader.scripts().entries[0].on).toBe(false);
  });

  it('turned on and then edited here stays on: the edit is the person’s own', async () => {
    const reader = tab();
    reader.service.openDocument('');
    await reader.settle();
    reader.service.openFile('budget.gsheet', await aFile(), 7);
    await reader.settle();
    reader.service.setFunctionsOn(true);
    await reader.settle();
    reader.service.saveScript('Taxes', 'Taxes', 'function TAX(amount, rate) { return amount * rate * 2; }');
    await reader.settle();
    expect(reader.shown(0, 1)).toBe(80);
    expect(reader.scripts().entries[0]).toMatchObject({ from: 'budget.gsheet', on: true });
  });

  it('never loads the interpreter while every function is off', async () => {
    const reader = tab();
    reader.service.openDocument('');
    await reader.settle();
    reader.service.openFile('budget.gsheet', await aFile('this.ran = true; function TAX() { return 1; }'), 7);
    await reader.settle();
    // Nothing of a script that is off runs, so its top level had nowhere to run.
    expect((reader.service as unknown as { sheetFunctions: unknown }).sheetFunctions).toBeNull();
  });
});

describe('an .xlsx of a workbook with functions', () => {
  it('holds the values the functions computed, and the formulas as written', async () => {
    const author = tab();
    author.service.openDocument('');
    await author.settle();
    author.service.saveScript('', 'Taxes', 'function TAX(amount, rate) { return amount * rate; }', 'functions');
    author.service.setCell(0, 0, '200');
    author.service.setCell(0, 1, '=TAX(A1, 0.2)');
    await author.settle();
    const document = (author.service as unknown as { document: SheetDocument }).document;
    const sheet = xlsxParts(xlsxOfDocument(document, 200).book).find(part => part.name === 'xl/worksheets/sheet1.xml')!.text;
    // Excel cannot run it, so the value is what it shows until it
    // recalculates; the formula is kept, so nothing is lost coming back.
    expect(sheet).toMatch(/<c r="B1"[^>]*><f>TAX\(A1, ?0\.2\)<\/f><v>40<\/v><\/c>/);
    // And the script itself is not in the file: an Excel file's code is VBA.
    expect(xlsxParts(xlsxOfDocument(document, 200).book).some(part => part.text.includes('amount * rate'))).toBe(false);
  });
});
