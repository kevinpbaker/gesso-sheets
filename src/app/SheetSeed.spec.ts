import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { SHEET, seed } from './SheetSeed';

/** Where the seed puts things, as its own layout has them. */
const TOTAL_ROW = 27;
const ROLLUP_ROW = 8;
const GLANCE_COUNT = 14;

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

  it('opens on the dashboard, in front of four more sheets', () => {
    const document = seeded();
    expect(document.active).toBe(SHEET.dashboard);
    expect(document.sheets().map(entry => entry.name)).toEqual(['Dashboard', 'Sales', 'Forecast', 'Reference', 'Summary']);
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
    document.activate(SHEET.sales);
    // North, at the list price Reference holds for it.
    expect(document.sheet.value(3, 5)).toBe(48);
    expect(document.sheet.value(3, 1)).toBe('Dana Okonjo');
    expect(document.sheet.value(3, 6)).toBe(180 * 48);
  });

  it('pays commission by the tier an order reaches, found by an approximate lookup', () => {
    const document = seeded();
    document.activate(SHEET.sales);
    // 180 × 48 = 8,640, which is past 5,000 and short of 15,000: 3%.
    expect(document.sheet.value(3, 10)).toBe(259.2);
  });

  it('totals a column through a name rather than an address', () => {
    const document = seeded();
    document.activate(SHEET.sales);
    let units = 0;
    for (let row = 3; row < TOTAL_ROW; row++) {
      units += document.sheet.value(row, 4) as number;
    }
    expect(document.sheet.value(TOTAL_ROW, 4)).toBe(units);
  });

  /**
   * The share column reads the total through an absolute reference, so
   * the shares add up to one — and go on doing so when a unit count
   * changes, which is the cascade this seed exists to show.
   */
  it('shares add to a whole, before and after an edit', () => {
    const document = seeded();
    document.activate(SHEET.sales);
    const share = (): number => {
      let total = 0;
      for (let row = 3; row < TOTAL_ROW; row++) {
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

  /**
   * An edit on `Reference` has to reach `Sales`, which reads it, then
   * the dashboard, which reads `Sales`, and then the forecast, which
   * reads the dashboard — three hops across four sheets, in an order
   * nothing here declares.
   */
  it('carries a price edit across four sheets in one recalculation', () => {
    const document = seeded();
    document.activate(SHEET.forecast);
    const before = document.sheet.value(24, 1) as number;
    document.activate(SHEET.dashboard);
    const year = document.sheet.value(4, 0) as number;

    document.activate(SHEET.reference);
    document.setCell(1, 2, '96.00');
    document.sheet.recalculate();

    document.activate(SHEET.dashboard);
    expect(document.sheet.value(4, 0)).toBeGreaterThan(year);
    document.activate(SHEET.forecast);
    expect(document.sheet.value(24, 1)).toBeGreaterThan(before);
  });

  it('spills the months, the regions and the largest orders from one formula each', () => {
    const document = seeded();
    const sheet = document.sheet;
    expect(sheet.spillOf(8, 0)).toEqual({ rows: 12, columns: 1 });
    expect([sheet.value(8, 0), sheet.value(19, 0)]).toEqual(['Oct 25', 'Sep 26']);
    expect(sheet.spillOf(8, 4)).toEqual({ rows: 5, columns: 1 });
    expect(sheet.value(8, 4)).toBe('Central');
    expect(sheet.spillOf(24, 0)).toEqual({ rows: 5, columns: 7 });
    // The largest order of the year, at the top of the block.
    expect(sheet.value(24, 6)).toBe(720 * 61.25);
    // And the months' revenue adds up to the year's.
    let months = 0;
    for (let row = 8; row <= 19; row++) {
      months += sheet.value(row, 1) as number;
    }
    expect(Math.round(months * 100) / 100).toBe(Math.round((sheet.value(4, 0) as number) * 100) / 100);
  });

  it('projects twelve months from one formula, and follows the scenario', () => {
    const document = seeded();
    document.activate(SHEET.forecast);
    const sheet = document.sheet;
    expect(sheet.spillOf(11, 1)).toEqual({ rows: 12, columns: 1 });
    const base = sheet.value(24, 1) as number;
    document.setCell(2, 1, 'High');
    sheet.recalculate();
    expect(sheet.value(3, 1)).toBe(0.04);
    expect(sheet.value(24, 1)).toBeGreaterThan(base);
    expect(document.setCell(2, 1, 'Hopeful')).toBe('A scenario is Low, Base or High.');
  });

  it('calls a named LAMBDA down a column, and hands it to MAP', () => {
    const document = seeded();
    document.activate(SHEET.sales);
    expect(document.sheet.input(3, 10)).toBe('=Commission(G4)');
    const total = document.sheet.value(TOTAL_ROW, 10) as number;
    document.activate(SHEET.summary);
    // The same commission, worked out by MAP over the revenue column.
    expect(document.sheet.value(ROLLUP_ROW + 3 + GLANCE_COUNT - 1, 1)).toBeCloseTo(total, 6);
    // And a LET: the largest order over the rest of them.
    const share = document.sheet.value(ROLLUP_ROW + 3 + GLANCE_COUNT - 2, 1) as number;
    expect(share).toBeGreaterThan(0.1);
    expect(share).toBeLessThan(0.2);
  });

  it('answers a lookup that cannot succeed with words instead of an error', () => {
    const document = seeded();
    document.activate(SHEET.summary);
    const last = document.sheet.value(ROLLUP_ROW + 3 + GLANCE_COUNT, 1);
    expect(last).toBe('not on file');
  });

  it('draws three charts on the dashboard, one of them from another sheet', () => {
    const document = seeded();
    expect(document.charts.map(chart => chart.kind)).toEqual(['line', 'column', 'pie']);
    expect(document.charts[2].range.start.sheet).toBe('Summary');
    document.activate(SHEET.forecast);
    expect(document.charts).toHaveLength(1);
  });

  it('keeps notes on the cells worth one', () => {
    const document = seeded();
    expect(document.noteAt(4, 0)).toMatch(/all 24 orders/);
    document.activate(SHEET.forecast);
    expect(document.noteAt(2, 1)).toMatch(/Low, Base or High/);
  });

  /**
   * The rules are per sheet, which is the claim `Page` makes and the
   * one a workbook-wide list would quietly break.
   */
  it('keeps each sheet’s rules on that sheet', () => {
    const document = seeded();
    expect(document.conditional).toHaveLength(3);
    document.activate(SHEET.sales);
    expect(document.conditional).toHaveLength(4);
    expect(document.validations).toHaveLength(2);
    document.activate(SHEET.reference);
    expect(document.conditional).toHaveLength(0);
    document.activate(SHEET.summary);
    expect(document.conditional).toHaveLength(1);
  });

  it('refuses a review it does not know, and keeps the one that was there', () => {
    const document = seeded();
    document.activate(SHEET.sales);
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
    document.activate(SHEET.sales);
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

  it('gives every sheet as many columns as the workbook has', () => {
    // The window is as wide as the widths add up to, so a sheet with
    // six of them is a row six columns wide — and a chart placed past
    // the sixth is laid out in it at no width at all.
    const document = new SheetDocument();
    document.columnWidths = Array.from({ length: 40 }, () => 96);
    document.book.extent = { rows: 1000, columns: 40 };
    seed(document);
    for (let sheet = 0; sheet < document.sheets().length; sheet++) {
      document.activate(sheet);
      expect(document.columnWidths, `sheet ${sheet}`).toHaveLength(40);
    }
  });
});
