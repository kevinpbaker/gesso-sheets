import { describe, expect, it } from 'vitest';

import type { SheetEditor, SheetScripts } from '../app/SheetContract';
import { SheetDocument } from '../app/SheetDocument';
import { SheetService, type Schedule } from '../app/SheetService';

/**
 * Formulas calling the workbook's script functions — Phase 32, through
 * `SheetService` and the real interpreter.
 */

function sheet(fill?: (document: SheetDocument) => void) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  fill?.(document);
  const service = new SheetService(document, { schedule });
  let scripts: SheetScripts = { entries: [], running: '', refused: '', last: null };
  let editor: SheetEditor | null = null;
  service.scripts.subscribe(next => (scripts = next));
  service.editor.subscribe(next => (editor = next));
  /** The recalculation, to the end. */
  const drain = () => {
    for (let turns = 0; queue.length > 0 && turns < 100_000; turns++) {
      queue.shift()!();
    }
  };
  const settle = async () => {
    await service.functionsReady;
    drain();
  };
  const value = (row: number, column: number) => document.sheet.value(row, column);
  return { service, document, settle, value, scripts: () => scripts, editor: () => editor };
}

const TAX = 'function TAX(amount, rate) { return Math.round(amount * rate * 100) / 100; }';

describe('a column that calls a script function', () => {
  it('is right, and recalculates when the function is edited', async () => {
    const { service, settle, value, scripts } = sheet();
    for (let row = 0; row < 50; row++) {
      service.setCell(row, 0, String((row + 1) * 10));
      service.setCell(row, 1, `=TAX(A${row + 1}, 0.2)`);
    }
    service.saveScript('', 'Taxes', TAX, 'functions');
    await settle();
    expect(scripts().entries[0]).toMatchObject({ kind: 'functions', defines: ['TAX'], problem: '' });
    expect(value(0, 1)).toBe(2);
    expect(value(49, 1)).toBe(100);

    service.saveScript('Taxes', 'Taxes', 'function TAX(amount, rate) { return amount * rate * 10; }');
    await settle();
    expect(value(0, 1)).toBe(20);
    expect(value(49, 1)).toBe(1000);
  });

  it('is #NAME? once the function is gone', async () => {
    const { service, settle, value } = sheet();
    service.setCell(0, 0, '=TAX(10, 0.5)');
    service.saveScript('', 'Taxes', TAX, 'functions');
    await settle();
    expect(value(0, 0)).toBe(5);
    service.removeScript('Taxes');
    await settle();
    expect(value(0, 0)).toMatchObject({ code: '#NAME?' });
  });

  it('spills rows a function returns, and takes a range as rows', async () => {
    const { service, settle, value } = sheet();
    service.setCell(0, 0, '3');
    service.setCell(1, 0, '4');
    service.setCell(0, 2, '=SQUARES(A1:A2)');
    service.saveScript('', 'Squares', 'function SQUARES(rows) { return rows.map(([n]) => [n, n * n]); }', 'functions');
    await settle();
    expect([value(0, 2), value(0, 3), value(1, 2), value(1, 3)]).toEqual([3, 9, 4, 16]);
  });
});

describe('a function that fails', () => {
  it('is an error in its cell, with its words in the formula bar, and the rest of the sheet goes on', async () => {
    const { service, settle, value, editor } = sheet();
    service.setCell(0, 0, '=CHECK(-1)');
    service.setCell(1, 0, '=CHECK(4)');
    service.setCell(2, 0, '=FOREVER()');
    service.setCell(3, 0, '=HOARD()');
    service.setCell(4, 0, '=1+1');
    service.saveScript(
      '',
      'Checks',
      `function CHECK(n) { if (n < 0) throw new Error("no negatives"); return n; }
       function FOREVER() { for (;;) {} }
       function HOARD() { const all = []; for (;;) all.push(new ArrayBuffer(1 << 20)); }`,
      'functions'
    );
    await settle();
    expect(value(0, 0)).toMatchObject({ code: '#VALUE!' });
    expect(value(1, 0)).toBe(4);
    expect(value(2, 0)).toMatchObject({ code: '#CALC!' });
    expect(value(3, 0)).toMatchObject({ code: '#CALC!' });
    expect(value(4, 0)).toBe(2);

    service.setSelection(0, 0, 0, 0);
    expect(editor()?.explain?.meaning).toBe('CHECK threw: no negatives');
    service.setSelection(2, 0, 2, 0);
    expect(editor()?.explain?.meaning).toBe('FOREVER ran past its 50ms.');
    service.setSelection(3, 0, 3, 0);
    expect(editor()?.explain?.meaning).toBe('HOARD ran out of memory.');
  }, 20_000);

  it('that is a mistake in one script leaves the other scripts’ functions working', async () => {
    const { service, settle, value, scripts } = sheet();
    service.setCell(0, 0, '=GOOD()');
    service.saveScript('', 'Broken', 'function BAD( {', 'functions');
    service.saveScript('', 'Fine', 'function GOOD() { return "ok"; }', 'functions');
    await settle();
    expect(value(0, 0)).toBe('ok');
    expect(scripts().entries.find(entry => entry.name === 'Broken')?.problem).toContain('stopped with an error');
  });

  it('named like a built-in cannot take the built-in’s place, and says so', async () => {
    const { service, settle, value, scripts } = sheet();
    service.setCell(0, 0, '=SUM(1, 2)');
    service.saveScript('', 'Mine', 'function SUM() { return "mine"; }', 'functions');
    await settle();
    expect(value(0, 0)).toBe(3);
    expect(scripts().entries[0].problem).toContain('the sheet’s own');
  });
});

describe('functions from a file', () => {
  it('are off, so a call to one is #NAME? and the script says why', async () => {
    const { service, settle, value, scripts } = sheet(document => {
      document.scripts = [{ name: 'Taxes', source: TAX, kind: 'functions', origin: { kind: 'file', file: 'budget.gsheet' } }];
      document.setCell(0, 0, '=TAX(10, 0.5)');
      document.sheet.recalculate();
    });
    await settle();
    expect(value(0, 0)).toMatchObject({ code: '#NAME?' });
    expect(scripts().entries[0].problem).toContain('Came with budget.gsheet');
  });

  it('are not run by Run', async () => {
    const { service, settle, scripts } = sheet();
    service.saveScript('', 'Taxes', TAX, 'functions');
    await settle();
    service.runScript('Taxes', false);
    expect(scripts().last).toMatchObject({ outcome: 'refused' });
  });
});
