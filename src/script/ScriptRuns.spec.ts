import { Worker } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';

import type { SheetScriptRun, SheetScripts } from '../app/SheetContract';
import { SheetDocument } from '../app/SheetDocument';
import { SheetService } from '../app/SheetService';
import type { ScriptWorker } from './ScriptHost';

/**
 * Scripts in the sheet — Phase 30's exit, through `SheetService` and the
 * real worker on a Node thread: a run that fills a column and formats
 * it, undone in one step; a run that never ends, stopped with the
 * workbook as it was; forbidden reaches that find nothing; and a file's
 * script that waits to be asked.
 */

function spawn(): ScriptWorker {
  const worker = new Worker(new URL('./nodeWorker.mjs', import.meta.url));
  return {
    post: message => worker.postMessage(message),
    onMessage: listener => worker.on('message', listener),
    terminate: () => void worker.terminate()
  };
}

function sheet(milliseconds = 3_000) {
  const document = new SheetDocument();
  const service = new SheetService(document, { scripts: spawn, scriptMilliseconds: milliseconds });
  let scripts: SheetScripts = { entries: [], running: '', refused: '', last: null };
  service.scripts.subscribe(next => (scripts = next));
  /** Runs a script and waits for how it ended. */
  const run = (name: string, confirmed = false): Promise<SheetScriptRun> => {
    const before = scripts.last?.serial ?? 0;
    service.runScript(name, confirmed);
    return new Promise(resolve => {
      const subscription = service.scripts.subscribe(next => {
        if (next.last !== null && next.last.serial > before) {
          queueMicrotask(() => subscription.unsubscribe());
          resolve(next.last);
        }
      });
    });
  };
  /** A cell's input on the sheet showing, read from what would be saved: the service may have opened another document. */
  const input = (row: number, column: number): string => {
    const snapshot = service.snapshot();
    return snapshot.sheets[snapshot.active].cells.find(cell => cell.row === row && cell.column === column)?.input ?? '';
  };
  return { service, document, run, scripts: () => scripts, input };
}

describe('a script in the sheet', () => {
  it('fills a column from another and formats it, and one undo takes all of it back', async () => {
    const { service, document, run, input } = sheet();
    service.setCell(0, 0, 'Price');
    service.setCell(1, 0, '10');
    service.setCell(2, 0, '=A2*3');
    service.saveScript(
      '',
      'With tax',
      `
        const prices = sheet.range("A2:A3").values;
        sheet.range("B1:B3").write([["With tax"], ...prices.map(([p]) => [p * 1.2])]);
        sheet.range("B1:B3").format({ bold: true, number: "currency" });
      `
    );

    const ran = await run('With tax');
    expect(ran).toMatchObject({ outcome: 'done', text: 'With tax changed 6 cells.' });
    expect([input(0, 1), input(1, 1), input(2, 1)]).toEqual(['With tax', '12', '36']);
    expect(document.formatAt(2, 1).paint.bold).toBe(true);
    expect(document.formatAt(2, 1).number.kind).toBe('currency');
    expect(document.undoLabel).toBe('script With tax');

    service.undo();
    expect([input(0, 1), input(1, 1), input(2, 1)]).toEqual(['', '', '']);
    expect(document.formatAt(2, 1).paint.bold).toBe(false);
    expect(document.formatAt(2, 1).number.kind).toBe('general');
    // What was there before the run is still there, and still undoable.
    expect(input(2, 0)).toBe('=A2*3');
    expect(document.canUndo).toBe(true);

    service.redo();
    expect(input(2, 1)).toBe('36');
  });

  it('that never ends is stopped by its time limit, with the workbook unchanged', async () => {
    const { service, document, run, input } = sheet(1_000);
    service.saveScript('', 'Forever', 'sheet.write("A1", "partial"); for (;;) {}');
    const undoable = document.canUndo;
    expect(await run('Forever')).toMatchObject({ outcome: 'timeout', text: 'Forever ran out of time, and nothing was changed.' });
    expect(input(0, 0)).toBe('');
    expect(document.canUndo).toBe(undoable);
  }, 10_000);

  it('can be stopped', async () => {
    const { service, run, scripts, input } = sheet();
    service.saveScript('', 'Forever', 'sheet.write("A1", 1); for (;;) {}');
    const ending = run('Forever');
    expect(scripts().running).toBe('Forever');
    service.stopScript();
    expect(await ending).toMatchObject({ outcome: 'stopped' });
    expect(scripts().running).toBe('');
    expect(input(0, 0)).toBe('');
  });

  it('that tries each forbidden reach gets nothing', async () => {
    const { service, run, input } = sheet();
    const probes = ['fetch', 'XMLHttpRequest', 'WebSocket', 'importScripts', 'postMessage', 'self', 'indexedDB', 'process', 'require', 'setTimeout'];
    service.saveScript(
      '',
      'Reach',
      `${JSON.stringify(probes)}.forEach((name, row) => {
        let found;
        try { found = typeof globalThis[name]; } catch (e) { found = "threw"; }
        sheet.range("A" + (row + 1)).write(found);
      });`
    );
    expect(await run('Reach')).toMatchObject({ outcome: 'done' });
    expect(probes.map((_, row) => input(row, 0))).toEqual(probes.map(() => 'undefined'));

    service.saveScript('', 'Import', 'import("node:fs")');
    expect(await run('Import')).toMatchObject({ outcome: 'failed' });
  });

  it('reads another sheet, and adds one that undo leaves in place', async () => {
    const { service, document, run } = sheet();
    service.setCell(0, 0, '7');
    service.saveScript('', 'Summary', 'workbook.addSheet("Summary").write("B2", sheet.read("A1") + 1)');
    expect(await run('Summary')).toMatchObject({
      outcome: 'done',
      text: 'Summary changed one cell. It added a sheet, which undo leaves in place.'
    });
    const added = document.pageAt(1)!;
    expect(added.sheet.name).toBe('Summary');
    expect(added.sheet.input(1, 1)).toBe('8');
    // The sheet that was showing still is.
    expect(document.active).toBe(0);
    service.undo();
    expect(document.sheetCount).toBe(2);
    expect(added.sheet.input(1, 1)).toBe('');
  });
});

describe('scripts are kept with the workbook', () => {
  it('saves, renames and removes them, and writes them down', () => {
    const { service, document, scripts } = sheet();
    service.saveScript('', 'One', 'sheet.write("A1", 1)');
    service.saveScript('', 'one', '');
    expect(scripts().refused).toBe('There is already a script called one.');
    service.saveScript('One', 'First', 'sheet.write("A1", 2)');
    expect(scripts().entries).toEqual([{ name: 'First', source: 'sheet.write("A1", 2)', from: '', kind: 'run', defines: [], problem: '', on: true }]);
    expect(service.snapshot().scripts).toEqual([
      { name: 'First', source: 'sheet.write("A1", 2)', origin: { kind: 'typed' } }
    ]);
    service.removeScript('First');
    expect(scripts().entries).toEqual([]);
    expect(service.snapshot().scripts).toBeUndefined();
  });

  it('marks every script in an opened file as the file’s, whatever the file says, and runs one only when asked', async () => {
    const { service, run, scripts, input } = sheet();
    const file = JSON.stringify({
      version: 3,
      active: 0,
      names: [],
      sheets: [{ name: 'Sheet1', cells: [] }],
      scripts: [{ name: 'Tidy', source: 'sheet.write("A1", "tidied")', origin: { kind: 'typed' } }]
    });
    service.openFile('budget.gsheet', file, null);
    await service.settled;
    expect(scripts().entries).toEqual([
      { name: 'Tidy', source: 'sheet.write("A1", "tidied")', from: 'budget.gsheet', kind: 'run', defines: [], problem: '', on: true }
    ]);

    const refused = await run('Tidy');
    expect(refused.outcome).toBe('refused');
    expect(refused.text).toContain('budget.gsheet');
    expect(input(0, 0)).toBe('');

    // Edited, it is still the file's.
    service.saveScript('Tidy', 'Tidy', 'sheet.write("A1", "edited")');
    expect(scripts().entries[0].from).toBe('budget.gsheet');

    expect(await run('Tidy', true)).toMatchObject({ outcome: 'done' });
    expect(input(0, 0)).toBe('edited');
  });
});
