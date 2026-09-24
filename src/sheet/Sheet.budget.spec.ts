import { describe, expect, it } from 'vitest';

import { Sheet } from './Sheet';

/**
 * Recalculation budgets — the exit criterion for Phase 1.
 *
 * The claim the whole project is evidence for is that a spreadsheet
 * can recalculate without the scroll ever noticing, and the first half
 * of that is doing no more arithmetic than the edit warrants. So these
 * are counts, not timings: editing one cell evaluates exactly the
 * cells that read it, transitively, and nothing else. A count fails
 * the build with a number the day a change makes recalc walk the sheet;
 * a timing would only say the machine was busy, and on a loaded CI box
 * it would say that whether or not anything regressed.
 *
 * The timings that are here are printed rather than asserted, so a
 * trend is visible without a slow machine failing the build. This is
 * the shape of Gesso's own `LayoutEngine.budget.spec.ts`, borrowed on
 * purpose: the two files are making the same kind of promise.
 */
describe('recalculation budgets', () => {
  /**
   * A column of N cells, each adding one to the last.
   *
   * The worst honest shape: every cell depends on the one above, so
   * editing the top is a transitive closure of everything below it and
   * there is exactly one valid order.
   */
  function chain(length: number): Sheet {
    const sheet = new Sheet();
    sheet.setCell(0, 0, '1');
    for (let row = 1; row < length; row++) {
      sheet.setCell(row, 0, `=A${row}+1`);
    }
    sheet.recalculate();
    return sheet;
  }

  function timed<T>(label: string, run: () => T): T {
    const start = performance.now();
    const result = run();
    console.info(`[recalc budget] ${label}: ${(performance.now() - start).toFixed(2)} ms`);
    return result;
  }

  it('evaluates exactly the transitive closure of the cell edited', () => {
    const DEPENDENTS = 50_000;
    const sheet = chain(DEPENDENTS + 1);
    const before = sheet.stats.evaluated;

    sheet.setCell(0, 0, '2');
    const result = timed(`one edit with ${DEPENDENTS} dependents`, () => sheet.recalculate());

    // The headline. Not "about N", not "under a limit": the closure is
    // 50,000 cells and 50,000 cells are evaluated. The edited cell is
    // a literal, so it is not one of them.
    expect(result.evaluated).toBe(DEPENDENTS);
    expect(sheet.stats.evaluated - before).toBe(DEPENDENTS);
    expect(sheet.value(DEPENDENTS, 0)).toBe(2 + DEPENDENTS);
  });

  it('evaluates nothing at all when the edit reaches nobody', () => {
    const sheet = chain(50_000);
    const before = sheet.stats.evaluated;

    // A cell off to the side that nothing reads.
    sheet.setCell(10, 5, '7');
    const result = sheet.recalculate();

    expect(result.evaluated).toBe(0);
    expect(sheet.stats.evaluated).toBe(before);
  });

  it('evaluates one cell when one cell reads the edit', () => {
    const sheet = new Sheet();
    sheet.setCell(0, 0, '1');
    sheet.setCell(0, 1, '=A1*2');
    sheet.recalculate();
    const before = sheet.stats.evaluated;

    sheet.setCell(0, 0, '5');
    sheet.recalculate();

    expect(sheet.stats.evaluated - before).toBe(1);
    expect(sheet.value(0, 1)).toBe(10);
  });

  /**
   * The diamond. Two formulas read one cell and a third reads both, so
   * the closure is three and the third must be evaluated once — after
   * both of its precedents, not once per path to it.
   */
  it('evaluates a shared dependent once, not once per path', () => {
    const sheet = new Sheet();
    sheet.setCell(0, 0, '1');
    sheet.setCell(1, 0, '=A1+1');
    sheet.setCell(2, 0, '=A1+2');
    sheet.setCell(3, 0, '=A2+A3');
    sheet.recalculate();
    const before = sheet.stats.evaluated;

    sheet.setCell(0, 0, '10');
    const result = sheet.recalculate();

    expect(result.evaluated).toBe(3);
    expect(sheet.stats.evaluated - before).toBe(3);
    expect(sheet.value(3, 0)).toBe(23);
  });

  it('does not re-evaluate a cell the edit did not reach, in a wide sheet', () => {
    const sheet = new Sheet();
    // Two hundred independent chains of ten. Editing the head of one
    // must cost nine, not one thousand nine hundred and ninety.
    for (let column = 0; column < 200; column++) {
      sheet.setCell(0, column, '1');
      for (let row = 1; row < 10; row++) {
        sheet.setCell(row, column, `=${columnName(column)}${row}+1`);
      }
    }
    sheet.recalculate();
    const before = sheet.stats.evaluated;

    sheet.setCell(0, 7, '100');
    sheet.recalculate();

    expect(sheet.stats.evaluated - before).toBe(9);
  });

  /**
   * A range is a hundred edges, and editing one cell in it is one
   * evaluation of the formula over it — not a hundred.
   */
  it('evaluates a formula over a range once per edit, not once per cell', () => {
    const sheet = new Sheet();
    for (let row = 0; row < 100; row++) {
      sheet.setCell(row, 0, String(row));
    }
    sheet.setCell(0, 1, '=SUM(A1:A100)');
    sheet.recalculate();
    const before = sheet.stats.evaluated;

    sheet.setCell(50, 0, '1000');
    const result = sheet.recalculate();

    expect(result.evaluated).toBe(1);
    expect(sheet.stats.evaluated - before).toBe(1);
    expect(sheet.value(0, 1)).toBe(4950 - 50 + 1000);
  });

  /**
   * The Phase 0 constraint, as a spec.
   *
   * A recalc that cannot be cut into slices blanks the sheet while it
   * runs: Phase 0 measured thirty milliseconds of uninterrupted
   * application thread leaving 89% of frames with a cell that had no
   * value in it. The budget is what lets the worker publish a window
   * between slices, so the engine has to honour one exactly and be
   * resumable across the gap.
   */
  it('stops at its budget and resumes where it left off', () => {
    const sheet = chain(1001);
    sheet.setCell(0, 0, '2');
    expect(sheet.pending).toBe(1000);

    let slices = 0;
    let total = 0;
    for (;;) {
      const result = sheet.recalculate(100);
      total += result.evaluated;
      slices++;
      if (result.done) {
        break;
      }
      expect(result.evaluated).toBe(100);
    }

    expect(slices).toBe(10);
    expect(total).toBe(1000);
    expect(sheet.pending).toBe(0);
    expect(sheet.value(1000, 0)).toBe(1002);
  });

  it('is correct when an edit lands in the middle of a slice', () => {
    const sheet = chain(1001);
    sheet.setCell(0, 0, '2');
    sheet.recalculate(100);

    // A second edit while a thousand cells are still queued. The plan
    // is dropped and rebuilt from what is still dirty; the hundred
    // already done are only redone if this edit reaches them.
    sheet.setCell(0, 0, '3');
    sheet.recalculate();

    expect(sheet.pending).toBe(0);
    expect(sheet.value(1000, 0)).toBe(1003);
    expect(sheet.value(1, 0)).toBe(4);
  });

  it('builds a full sheet of formulas within a generous ceiling', () => {
    const CELLS = 50_000;
    const sheet = timed(`building and recalculating ${CELLS} formulas`, () => chain(CELLS));
    expect(sheet.value(CELLS - 1, 0)).toBe(CELLS);
  });
});

function columnName(index: number): string {
  let name = '';
  let remaining = index;
  while (remaining >= 0) {
    name = String.fromCharCode(65 + (remaining % 26)) + name;
    remaining = Math.floor(remaining / 26) - 1;
  }
  return name;
}
