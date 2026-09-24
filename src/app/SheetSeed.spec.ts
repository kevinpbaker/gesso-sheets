import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { seed } from './SheetSeed';

/**
 * The seed, end to end through the engine.
 *
 * Not a fixture check. It is the one place the pieces of Phase 1 are
 * exercised together the way a person uses them — a formula reading a
 * formula, a range, an absolute reference, a rounding, an IF — and it
 * would be the first thing to break if the parser and the graph ever
 * disagreed about what a sheet means.
 */
describe('the seeded sheet', () => {
  function seeded(): SheetDocument {
    const document = new SheetDocument();
    seed(document);
    document.sheet.recalculate();
    return document;
  }

  it('settles with nothing left to do', () => {
    expect(seeded().sheet.pending).toBe(0);
  });

  it('multiplies units by price into revenue', () => {
    const document = seeded();
    expect(document.sheet.value(1, 3)).toBe(120 * 9.5);
  });

  it('totals a column with a range', () => {
    const document = seeded();
    let units = 0;
    for (let row = 1; row <= 5; row++) {
      units += document.sheet.value(row, 1) as number;
    }
    expect(document.sheet.value(6, 1)).toBe(units);
  });

  it('holds no errors anywhere', () => {
    const document = seeded();
    for (const { row, column } of document.sheet.entries()) {
      expect(document.sheet.display(row, column), `at ${row},${column}`).not.toMatch(/^#/);
    }
  });

  /**
   * The share column reads the total through an absolute reference,
   * so the five of them add up to a hundred — and go on doing so when
   * a unit count changes, which is the cascade this seed exists to
   * show.
   */
  it('shares add to a hundred, before and after an edit', () => {
    const document = seeded();
    const share = () => {
      let total = 0;
      for (let row = 1; row <= 5; row++) {
        total += document.sheet.value(row, 4) as number;
      }
      return Math.round(total);
    };
    expect(share()).toBe(100);

    document.setCell(1, 1, '900');
    document.sheet.recalculate();

    expect(share()).toBe(100);
    expect(document.sheet.value(1, 3)).toBe(900 * 9.5);
  });

  it('cascades one edit into the cells that read it, transitively', () => {
    const document = seeded();
    const before = document.sheet.value(6, 3) as number;

    document.setCell(1, 1, '240');
    document.sheet.recalculate();

    // Units -> revenue -> total revenue, and separately -> largest.
    expect(document.sheet.value(6, 3)).not.toBe(before);

    const revenues: number[] = [];
    for (let row = 1; row <= 5; row++) {
      revenues.push(document.sheet.value(row, 3) as number);
    }
    expect(document.sheet.value(8, 1)).toBe(Math.max(...revenues));
    expect(document.sheet.value(9, 1)).toBe(Math.min(...revenues));
  });
});
