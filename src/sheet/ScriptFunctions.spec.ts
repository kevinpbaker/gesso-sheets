import { describe, expect, it } from 'vitest';

import { isArray, type ArrayValue } from './FunctionKit';
import type { ScriptFunctions } from './ScriptFunctions';
import { NAME, VALUE, type CellValue } from './Values';
import { Workbook } from './Workbook';

/**
 * Formulas calling the script's functions — Phase 32, the sheet's half.
 *
 * Against a stand-in for the interpreter, because what is asked here is
 * the evaluator's: where a script function sits among the other names,
 * what it is handed, what it can answer, and what makes the sheet ask
 * again.
 */

function stand(defined: Record<string, (args: readonly (CellValue | ArrayValue)[]) => CellValue | ArrayValue>) {
  const calls: { name: string; args: readonly (CellValue | ArrayValue)[]; at: unknown }[] = [];
  let spent = 0;
  let limit = Number.POSITIVE_INFINITY;
  const scripts: ScriptFunctions = {
    has: name => Object.prototype.hasOwnProperty.call(defined, name),
    call: (name, args, at) => {
      calls.push({ name, args, at });
      spent++;
      return defined[name](args);
    },
    startSlice: () => (spent = 0),
    overBudget: () => spent >= limit
  };
  return { scripts, calls, limitTo: (calls: number) => (limit = calls) };
}

function book(scripts: ScriptFunctions) {
  const workbook = new Workbook();
  workbook.scripts = scripts;
  workbook.scriptsChanged();
  return { workbook, sheet: workbook.sheet(0) };
}

describe('a formula calling a script function', () => {
  it('is handed values, and ranges as arrays, and shows what comes back', () => {
    const { scripts, calls } = stand({
      TAX: ([amount, rate]) => (amount as number) * (rate as number),
      TOTAL: ([range]) => (isArray(range) ? range.values.reduce<number>((sum, v) => sum + (v as number), 0) : VALUE)
    });
    const { sheet } = book(scripts);
    sheet.setCell(0, 0, '200');
    sheet.setCell(1, 0, '50');
    sheet.setCell(0, 1, '=TAX(A1, 0.2)');
    sheet.setCell(1, 1, '=TOTAL(A1:A2)');
    sheet.recalculate();
    expect(sheet.value(0, 1)).toBe(40);
    expect(sheet.value(1, 1)).toBe(250);
    expect(calls.find(call => call.name === 'TOTAL')?.args[0]).toMatchObject({ kind: 'array', rows: 2, columns: 1, values: [200, 50] });
    expect(calls.find(call => call.name === 'TAX')?.at).toEqual({ sheet: 0, row: 0, column: 1 });
  });

  it('recalculates when what it reads changes, like any formula', () => {
    const { scripts } = stand({ DOUBLE: ([n]) => (n as number) * 2 });
    const { sheet } = book(scripts);
    sheet.setCell(0, 0, '4');
    sheet.setCell(0, 1, '=DOUBLE(A1)');
    sheet.recalculate();
    sheet.setCell(0, 0, '5');
    sheet.recalculate();
    expect(sheet.value(0, 1)).toBe(10);
  });

  it('spills an array it returns', () => {
    const { scripts } = stand({ PAIR: () => ({ kind: 'array', rows: 1, columns: 2, values: ['a', 'b'] }) });
    const { sheet } = book(scripts);
    sheet.setCell(0, 0, '=PAIR()');
    sheet.recalculate();
    expect([sheet.value(0, 0), sheet.value(0, 1)]).toEqual(['a', 'b']);
  });

  it('comes after every built-in and every name, so neither can be shadowed', () => {
    const { scripts, calls } = stand({ SUM: () => 'the script', RATE2: () => 'the script' });
    const { workbook, sheet } = book(scripts);
    workbook.names.defineFormula('RATE2', '=LAMBDA(x, "the name")');
    workbook.namesChanged();
    sheet.setCell(0, 0, '=SUM(1, 2)');
    sheet.setCell(0, 1, '=RATE2(1)');
    sheet.recalculate();
    expect(sheet.value(0, 0)).toBe(3);
    expect(sheet.value(0, 1)).toBe('the name');
    expect(calls).toEqual([]);
  });

  it('is #NAME? until the script defines it, and has an answer once it does', () => {
    const workbook = new Workbook();
    const sheet = workbook.sheet(0);
    sheet.setCell(0, 0, '=LATER(2)');
    sheet.recalculate();
    expect(sheet.value(0, 0)).toBe(NAME);
    workbook.scripts = stand({ LATER: ([n]) => (n as number) + 1 }).scripts;
    workbook.scriptsChanged();
    sheet.recalculate();
    expect(sheet.value(0, 0)).toBe(3);
  });

  it('hands the thread back when a slice has spent its share in scripts', () => {
    const { scripts, limitTo } = stand({ ONE: () => 1 });
    const { sheet } = book(scripts);
    for (let row = 0; row < 100; row++) {
      sheet.setCell(row, 0, '=ONE()');
    }
    limitTo(10);
    const slice = sheet.recalculate(2_000);
    expect(slice.evaluated).toBe(10);
    expect(slice.done).toBe(false);
    while (!sheet.recalculate(2_000).done) {
      // The rest, ten at a time.
    }
    expect(sheet.value(99, 0)).toBe(1);
  });
});
