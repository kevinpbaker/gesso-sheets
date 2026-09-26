import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { seed } from './SheetSeed';

/**
 * The seed, end to end through the engine.
 *
 * Not a fixture check. It is the one place the pieces of the sheet
 * are exercised together the way a person uses them — a formula
 * reading a formula, a lookup crossing a sheet, a range, an absolute
 * reference, a name, a rounding, an IF — and it would be the first
 * thing to break if the parser, the graph and the workbook ever
 * disagreed about what a sheet means.
 *
 * It is also the only spec that asserts the sheet somebody actually
 * sees first. A seed that quietly grew a `#REF!` would be wrong in
 * the most visible place there is, and nothing else here would catch
 * it.
 */
describe('the seeded workbook', () => {
  function seeded(): SheetDocument {
    const document = new SheetDocument();
    seed(document);
    document.sheet.recalculate();
    return document;
  }

  /** Every cell on every sheet, so nothing hides on a tab nobody opened. */
  function everywhere(document: SheetDocument, visit: (sheet: number, row: number, column: number) => void): void {
    const active = document.active;
    for (let sheet = 0; sheet < document.sheets().length; sheet++) {
      document.activate(sheet);
      for (const { row, column } of document.sheet.entries()) {
        visit(sheet, row, column);
      }
    }
    document.activate(active);
  }

  it('settles with nothing left to do', () => {
    expect(seeded().sheet.pending).toBe(0);
  });

  it('opens on the sheet with the orders on it', () => {
    const document = seeded();
    expect(document.active).toBe(0);
    expect(document.sheets().map(entry => entry.name)).toEqual(['Sales', 'Reference', 'Summary']);
  });

  it('holds no errors anywhere, on any sheet', () => {
    const document = seeded();
    everywhere(document, (sheet, row, column) => {
      expect(document.display(row, column), `at sheet ${sheet}, ${row},${column}`).not.toMatch(/^#/);
    });
  });

  /**
   * The lookup is the join between two sheets, and it is the thing
   * that breaks if a sheet is renamed, moved, or written before the
   * sheet it reads exists.
   */
  it('prices an order by looking its region up on another sheet', () => {
    const document = seeded();
    // North, at the list price Reference holds for it.
    expect(document.sheet.value(3, 5)).toBe(48);
    expect(document.sheet.value(3, 1)).toBe('Dana Okonjo');
    expect(document.sheet.value(3, 6)).toBe(180 * 48);
  });

  it('totals a column through a name rather than an address', () => {
    const document = seeded();
    let units = 0;
    for (let row = 3; row <= 12; row++) {
      units += document.sheet.value(row, 4) as number;
    }
    expect(document.sheet.value(13, 4)).toBe(units);
  });

  /**
   * The share column reads the total through an absolute reference,
   * so the ten of them add up to one — and go on doing so when a unit
   * count changes, which is the cascade this seed exists to show.
   */
  it('shares add to a whole, before and after an edit', () => {
    const document = seeded();
    const share = (): number => {
      let total = 0;
      for (let row = 3; row <= 12; row++) {
        total += document.sheet.value(row, 7) as number;
      }
      return Math.round(total * 1000) / 1000;
    };
    expect(share()).toBe(1);

    document.setCell(3, 4, '900');
    document.sheet.recalculate();

    expect(share()).toBe(1);
    expect(document.sheet.value(3, 6)).toBe(900 * 48);
  });

  it('cascades one edit into the cells that read it, transitively', () => {
    const document = seeded();
    const before = document.sheet.value(13, 6) as number;

    document.setCell(3, 4, '240');
    document.sheet.recalculate();

    expect(document.sheet.value(13, 6)).not.toBe(before);
  });

  /**
   * An edit on `Reference` has to reach `Sales`, which reads it, and
   * then `Summary`, which reads `Sales` — two hops across three
   * sheets, in an order nothing here declares.
   */
  it('carries an edit across two sheets in one recalculation', () => {
    const document = seeded();
    document.activate(2);
    const before = document.sheet.value(8, 3) as number;

    document.activate(1);
    // North's list price, doubled.
    document.setCell(1, 2, '96.00');
    document.sheet.recalculate();

    document.activate(2);
    expect(document.sheet.value(8, 3)).toBeGreaterThan(before);
  });

  it('answers a lookup that cannot succeed with words instead of an error', () => {
    const document = seeded();
    document.activate(2);
    expect(document.sheet.value(18, 1)).toBe('not on file');
  });

  it('names the region that took the largest order', () => {
    const document = seeded();
    document.activate(2);
    // East's 610 units at 44.75 is the largest order on the sheet.
    expect(document.sheet.value(14, 1)).toBe('East');
  });

  /**
   * The rules are per sheet, which is the claim `Page` makes and the
   * one a workbook-wide list would quietly break.
   */
  it('keeps each sheet’s rules on that sheet', () => {
    const document = seeded();
    expect(document.conditional).toHaveLength(3);
    expect(document.validations).toHaveLength(2);

    document.activate(1);
    expect(document.conditional).toHaveLength(0);

    document.activate(2);
    expect(document.conditional).toHaveLength(1);
  });

  it('refuses a review it does not know, and keeps the one that was there', () => {
    const document = seeded();
    const before = document.sheet.input(3, 9);

    expect(document.setCell(3, 9, 'Maybe')).toBe('A review is Approved, Pending or Held.');
    expect(document.sheet.input(3, 9)).toBe(before);
    expect(document.setCell(3, 9, 'Held')).toBeNull();
  });

  /**
   * The lenient rule, which is the other half of the pair: a unit
   * count outside the usual range is marked and allowed to stand.
   */
  it('lets an unusual unit count through and marks it', () => {
    const document = seeded();
    expect(document.setCell(4, 4, '9000')).toBeNull();
    expect(document.sheet.value(4, 4)).toBe(9000);
  });

  it('gives every sheet a width for every column it draws', () => {
    const document = seeded();
    for (let sheet = 0; sheet < document.sheets().length; sheet++) {
      document.activate(sheet);
      const widths = document.columnWidths;
      expect(widths.length, `sheet ${sheet}`).toBeGreaterThan(0);
      for (const width of widths) {
        expect(width, `sheet ${sheet}`).toBeGreaterThan(0);
      }
    }
  });
});
