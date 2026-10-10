import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import { Workbook } from './Workbook';

/**
 * A fork is a scenario: the workbook as it stands, with a few inputs
 * typed in differently, recalculated.
 *
 * The claims are the ones a copy could get wrong. The base must not
 * move when the fork does. The fork must recalculate what its
 * overrides reach and nothing else, which is the whole case for
 * forking over rebuilding. And it must not differ from its base
 * anywhere its overrides do not reach — not in a `RAND()`, not in a
 * spilled array, not across a sheet.
 */

/** A growth model: an input, a rate, a twelve-month spill, a total on a second sheet. */
function model(): Workbook {
  const w = new Workbook(['Inputs', 'Plan']);
  w.setCell(0, 0, 1, '100'); // Inputs!B1, the start
  w.setCell(0, 1, 1, '0.05'); // Inputs!B2, the growth
  w.setCell(1, 0, 0, '=ROUND(Inputs!B1*(1+Inputs!B2)^SEQUENCE(12),2)'); // Plan!A1:A12
  w.setCell(1, 0, 2, '=SUM(A1:A12)'); // Plan!C1
  w.setCell(1, 1, 2, 'Steady'); // Plan!C2, a label nothing reads
  for (let row = 0; row < 50; row++) {
    w.setCell(1, row, 4, String(row)); // Plan!E1:E50, data the growth does not touch
  }
  w.setCell(1, 0, 5, '=SUM(E1:E50)'); // Plan!F1
  w.recalculate();
  return w;
}

describe('a fork', () => {
  it('starts as its base, value for value', () => {
    const base = model();
    const fork = base.fork();
    expect(fork.pending).toBe(0);
    expect(fork.value(1, 0, 2)).toBe(base.value(1, 0, 2));
    expect(fork.value(1, 11, 0)).toBe(base.value(1, 11, 0));
    expect(fork.input(1, 0, 0)).toBe(base.input(1, 0, 0));
  });

  it('answers with its overrides, and its base does not move', () => {
    const base = model();
    const before = base.value(1, 0, 2);
    const fork = base.fork([{ sheet: 0, row: 1, column: 1, input: '0.1' }]);
    fork.recalculate();
    expect(fork.value(0, 1, 1)).toBe(0.1);
    expect(fork.value(1, 0, 0)).toBe(110);
    expect(fork.value(1, 0, 2)).not.toBe(before);
    expect(base.value(0, 1, 1)).toBe(0.05);
    expect(base.value(1, 0, 2)).toBe(before);
    expect(base.pending).toBe(0);
  });

  it('recalculates what its overrides reach, and nothing else', () => {
    const base = model();
    const fork = base.fork([{ sheet: 0, row: 1, column: 1, input: '0.1' }]);
    // The same work the same edit is in the base: the spill, the cells it
    // fills and their total, and not the fifty cells of data or their sum.
    const edited = model();
    edited.setCell(0, 1, 1, '0.1');
    expect(fork.pending).toBe(edited.pending);
    expect(fork.recalculate().evaluated).toBe(edited.recalculate().evaluated);
    expect(fork.pending).toBeLessThan(15);
  });

  it('does not touch its base when the base is edited after it', () => {
    const base = model();
    const fork = base.fork([{ sheet: 0, row: 1, column: 1, input: '0.1' }]);
    fork.recalculate();
    const forked = fork.value(1, 0, 2);
    base.setCell(0, 0, 1, '200');
    base.recalculate();
    expect(fork.value(1, 0, 2)).toBe(forked);
    expect(fork.pending).toBe(0);
  });

  it('keeps its base’s roll of the dice where nothing it changed reaches', () => {
    const base = model();
    base.setCell(1, 5, 2, '=RAND()');
    base.setCell(1, 6, 2, '=C6*2');
    base.recalculate();
    const fork = base.fork([{ sheet: 0, row: 1, column: 1, input: '0.1' }]);
    fork.recalculate();
    expect(fork.value(1, 5, 2)).toBe(base.value(1, 5, 2));
    expect(fork.value(1, 6, 2)).toBe(base.value(1, 6, 2));
  });

  it('can be forked again, and each fork is its own', () => {
    const base = model();
    const low = base.fork([{ sheet: 0, row: 1, column: 1, input: '0' }]);
    const high = base.fork([{ sheet: 0, row: 1, column: 1, input: '0.2' }]);
    low.recalculate();
    high.recalculate();
    expect(low.value(1, 0, 2)).toBe(1200);
    expect(high.value(1, 0, 0)).toBe(120);
    expect(base.value(1, 0, 0)).toBe(105);
  });

  it('takes a formula as an override, wired like any formula', () => {
    const base = model();
    const fork = base.fork([{ sheet: 0, row: 1, column: 1, input: '=0.02*5' }]);
    fork.recalculate();
    expect(fork.value(1, 0, 0)).toBe(110);
    fork.setCell(0, 0, 1, '200');
    fork.recalculate();
    expect(fork.value(1, 0, 0)).toBe(220);
    expect(base.value(1, 0, 0)).toBe(105);
  });

  it('knows its base’s names', () => {
    const w = new Workbook(['Sheet1']);
    w.setCell(0, 0, 0, '3');
    w.names.define('Rate', { start: relativeRef(0, 0), end: relativeRef(0, 0) });
    w.namesChanged();
    w.setCell(0, 1, 0, '=Rate*2');
    w.recalculate();
    const fork = w.fork([{ sheet: 0, row: 0, column: 0, input: '5' }]);
    fork.recalculate();
    expect(fork.value(0, 1, 0)).toBe(10);
    expect(w.value(0, 1, 0)).toBe(6);
  });
});
