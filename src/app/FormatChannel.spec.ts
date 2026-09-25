import { beforeEach, describe, expect, it } from 'vitest';

import { attach, RecordingPort, type Harness } from './sheetHarness';

/**
 * Patch budgets for formatting — the exit criterion for Phase 9.
 *
 * "Formatting a column of 50,000 cells emits one palette patch and
 * one patch per visible cell, `toHaveLength` and not a ceiling."
 *
 * Same machinery as Phase 2's: `provide` over a recording port with
 * the framework's own differ in between, so these are the patches
 * that would cross a `postMessage`. A count rather than a timing,
 * for the reason every budget in this repository is a count — it
 * fails the build with a number when a change puts the sheet on the
 * wire instead of the window.
 */

function command(port: RecordingPort, name: string, ...args: unknown[]): void {
  port.send({ type: 'channel:command', command: name, payload: args[0], rest: args.slice(1) });
}

const CURRENCY = { kind: 'currency', places: 2, symbol: '$' } as const;

describe('the cost of formatting', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach();
  });

  /**
   * The claim, in two numbers.
   *
   * Fifty thousand cells formatted at once, thirty of them on screen:
   * one entry appears in the palette and thirty indices cross. The
   * palette is a table and the window is a window, which is the whole
   * reason they are two keys.
   */
  it('costs one palette entry and one index per visible cell', () => {
    const CELLS = 50_000;
    const VISIBLE = 30;
    for (let row = 0; row < CELLS; row++) {
      command(h.port, 'setCell', row, 0, String(row));
    }
    h.clock.drain();
    command(h.port, 'setViewport', 0, VISIBLE - 1, 0, 0);
    command(h.port, 'setSelection', 0, 0, CELLS - 1, 0);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'format', { number: CURRENCY });
    h.clock.drain();

    // One entry added to the palette, whoever else is using it.
    expect(h.port.patchesFor('palette')).toHaveLength(1);
    // One index per visible cell, and not one per formatted cell.
    expect(h.port.patchesFor('formats')).toHaveLength(VISIBLE);
  });

  /**
   * The same claim stated the other way round, which is the version
   * that would catch a regression the first one would not: format
   * fifty thousand cells that nobody is looking at and nothing about
   * *them* crosses at all.
   */
  it('sends no indices for cells that are not on screen', () => {
    const CELLS = 50_000;
    for (let row = 0; row < CELLS; row++) {
      command(h.port, 'setCell', row, 0, String(row));
    }
    h.clock.drain();
    // Looking at the far end of the sheet, nowhere near the column.
    command(h.port, 'setViewport', 0, 29, 5, 9);
    command(h.port, 'setSelection', 0, 0, CELLS - 1, 0);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'format', { bold: true });
    h.clock.drain();

    expect(h.port.patchesFor('formats')).toEqual([]);
    // The palette still grows by one: it is a table of what exists,
    // not of what is visible, and one entry is what it costs.
    expect(h.port.patchesFor('palette')).toHaveLength(1);
  });

  /**
   * A number format changes the *string* a cell shows, so the window
   * has to go out too — and it has to go out bounded by the window
   * rather than by the range that was formatted.
   */
  it('republishes only the visible strings when a number format changes', () => {
    const VISIBLE = 30;
    for (let row = 0; row < 5_000; row++) {
      command(h.port, 'setCell', row, 0, '1234.5');
    }
    h.clock.drain();
    command(h.port, 'setViewport', 0, VISIBLE - 1, 0, 0);
    command(h.port, 'setSelection', 0, 0, 4_999, 0);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'format', { number: CURRENCY });
    h.clock.drain();

    expect(h.port.patchesFor('window')).toHaveLength(VISIBLE);
    expect(h.window().cells[0][0]).toBe('$1,234.50');
  });

  /**
   * An unformatted sheet must cost nothing here, forever. A cell with
   * the default format is left out of the window rather than sent as
   * a zero, so after the first publish the differ has nothing to say.
   */
  it('sends nothing at all for a sheet nobody has formatted', () => {
    command(h.port, 'setCell', 0, 0, '1');
    h.clock.drain();
    command(h.port, 'setViewport', 0, 29, 0, 9);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, '2');
    h.clock.drain();

    expect(h.port.patchesFor('formats')).toEqual([]);
    expect(h.port.patchesFor('palette')).toEqual([]);
  });

  /**
   * A scroll over a formatted sheet costs what a scroll over an
   * unformatted one costs: the row that entered, the row that left,
   * and the two bounds. Phase 2 asserts this for `window`; without
   * the same assertion here a format window shaped as an array would
   * make every scroll cost a screenful.
   */
  it('costs a scroll the row that entered and the row that left', () => {
    for (let row = 0; row < 200; row++) {
      command(h.port, 'setCell', row, 0, String(row));
    }
    h.clock.drain();
    command(h.port, 'setViewport', 0, 29, 0, 0);
    command(h.port, 'setSelection', 0, 0, 199, 0);
    h.clock.drain();
    command(h.port, 'format', { bold: true });
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setViewport', 1, 30, 0, 0);

    // Four: the row that arrived, the row that left, and the two
    // bounds that moved.
    expect(h.port.patchesFor('formats')).toHaveLength(4);
  });

  /**
   * The palette only ever grows, and formatting the same way twice
   * must not grow it. Interning is what makes the first spec in this
   * file true a second time.
   */
  it('does not grow the palette for a format it already holds', () => {
    command(h.port, 'setCell', 0, 0, '1');
    command(h.port, 'setCell', 1, 0, '2');
    h.clock.drain();
    command(h.port, 'setViewport', 0, 29, 0, 0);
    command(h.port, 'setSelection', 0, 0, 0, 0);
    command(h.port, 'format', { bold: true });
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setSelection', 1, 0, 1, 0);
    command(h.port, 'format', { bold: true });
    h.clock.drain();

    expect(h.port.patchesFor('palette')).toEqual([]);
    // And the second cell points at the entry the first one made.
    expect(h.port.patchesFor('formats')).toHaveLength(1);
  });
});

/**
 * A region formatted as a region — the bug a browser found.
 *
 * Cell by cell, ctrl-A then ctrl-B wrote a million cell entries, a
 * million-entry undo step, and a thirty-megabyte file that was read
 * back on every load. Every spec passed: they all asserted what the
 * document *said*, and the document said the right thing at a cost
 * nobody had counted. These count it.
 */
describe('formatting a whole region', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach();
  });

  /** A sheet small enough to count, shaped like the real one. */
  function sheet(rows: number, columns: number): Harness {
    return attach({ rowCount: rows, columnCount: columns });
  }

  it('stores the whole sheet as one entry, not one per cell', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setCell', 0, 0, 'x');
    h.clock.drain();
    command(h.port, 'setSelection', 0, 0, 999, 49);
    command(h.port, 'format', { bold: true });
    h.clock.drain();

    // Fifty thousand cells are bold and nothing is stored per cell.
    expect(h.document.formatAt(0, 0).paint.bold).toBe(true);
    expect(h.document.formatAt(999, 49).paint.bold).toBe(true);
    expect(h.document.formats.size).toBe(0);
  });

  it('stores a whole column as one entry', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setSelection', 0, 3, 999, 3);
    command(h.port, 'format', { number: CURRENCY });
    h.clock.drain();

    expect(h.document.formatAt(500, 3).number.kind).toBe('currency');
    expect(h.document.formatAt(500, 4).number.kind).toBe('general');
    expect(h.document.formats.size).toBe(0);
  });

  it('stores a whole row as one entry', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setSelection', 7, 0, 7, 49);
    command(h.port, 'format', { italic: true });
    h.clock.drain();

    expect(h.document.formatAt(7, 20).paint.italic).toBe(true);
    expect(h.document.formatAt(8, 20).paint.italic).toBe(false);
    expect(h.document.formats.size).toBe(0);
  });

  /** A cell somebody formatted deliberately beats the column it is in. */
  it('lets a cell override the region under it', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setSelection', 0, 3, 999, 3);
    command(h.port, 'format', { number: CURRENCY });
    command(h.port, 'setSelection', 5, 3, 5, 3);
    command(h.port, 'format', { number: { kind: 'percent', places: 0 } });
    h.clock.drain();

    expect(h.document.formatAt(5, 3).number.kind).toBe('percent');
    expect(h.document.formatAt(6, 3).number.kind).toBe('currency');
    expect(h.document.formats.size).toBe(1);
  });

  /**
   * A region change reaches the cells inside it that had a format of
   * their own — it does not wipe them.
   *
   * The first version wiped them, and a browser showed what that
   * means in one keystroke: ctrl-A then ctrl-B made the sheet bold
   * and threw away the currency in column C and the fill on the
   * header row. A change is a change, and "make this bold" has
   * nothing to say about anybody's currency symbol.
   */
  it('applies the change to the overrides inside it, keeping the rest', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setSelection', 5, 3, 5, 3);
    command(h.port, 'format', { number: CURRENCY });
    command(h.port, 'setSelection', 0, 3, 999, 3);
    command(h.port, 'format', { italic: true });
    h.clock.drain();

    const special = h.document.formatAt(5, 3);
    expect(special.paint.italic).toBe(true);
    expect(special.number.kind).toBe('currency');
    // And its neighbour got the region's version.
    expect(h.document.formatAt(6, 3).paint.italic).toBe(true);
    expect(h.document.formatAt(6, 3).number.kind).toBe('general');
  });

  it('is one press of ctrl-Z however far it reached', () => {
    h = sheet(1_000, 50);
    command(h.port, 'setSelection', 5, 3, 5, 3);
    command(h.port, 'format', { number: CURRENCY });
    command(h.port, 'setSelection', 0, 3, 999, 3);
    command(h.port, 'format', { italic: true });
    h.clock.drain();

    command(h.port, 'undo');
    h.clock.drain();

    expect(h.document.formatAt(5, 3).number.kind).toBe('currency');
    expect(h.document.formatAt(5, 3).paint.italic).toBe(false);
    expect(h.document.formatAt(6, 3).paint.italic).toBe(false);
  });

  /**
   * The file is the harm the browser actually showed: thirty
   * megabytes written to somebody's disk, and parsed back on every
   * load for the rest of the sheet's life.
   */
  it('writes a sheet-wide format as a few hundred bytes', () => {
    h = sheet(10_000, 100);
    command(h.port, 'setCell', 0, 0, 'x');
    h.clock.drain();
    command(h.port, 'setSelection', 0, 0, 9_999, 99);
    command(h.port, 'format', { bold: true });
    h.clock.drain();

    const written = JSON.stringify(h.service.snapshot());
    // Two kilobytes against the thirty megabytes the per-cell version
    // wrote. The bound is loose on purpose: what it is guarding is
    // the *shape* — a file that grows with the number of distinct
    // formats and not with the number of cells — and a palette entry
    // carrying four border edges is a hundred bytes wider than one
    // that did not.
    expect(written.length).toBeLessThan(2_000);
  });

  /** And the wire is unchanged, which is what makes it all work. */
  it('still costs the viewport on the wire', () => {
    h = sheet(10_000, 100);
    command(h.port, 'setViewport', 0, 29, 0, 9);
    command(h.port, 'setSelection', 0, 0, 9_999, 99);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'format', { bold: true });
    h.clock.drain();

    expect(h.port.patchesFor('palette')).toHaveLength(1);
    // One per *visible cell* — thirty rows of ten — and not one per
    // formatted cell, of which there are a million. That ratio is the
    // whole claim: the wire carries the window and the store carries
    // one number.
    expect(h.port.patchesFor('formats')).toHaveLength(300);
  });
});

/**
 * A frozen pane is part of what is on screen, so its cells have to be
 * sent — and only its cells.
 *
 * The frozen columns were drawn empty until this: correctly placed,
 * correctly stuck, and holding nothing, because the window they would
 * have come from had scrolled past them. Found by freezing a column
 * in a browser and scrolling sideways.
 */
describe('the cells a frozen pane needs', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach({ rowCount: 1_000, columnCount: 40 });
  });

  it('sends the frozen column as well as the scrolled window', () => {
    command(h.port, 'setCell', 0, 0, 'frozen');
    command(h.port, 'setCell', 0, 20, 'scrolled');
    h.clock.drain();
    command(h.port, 'freeze', 0, 1);
    command(h.port, 'setViewport', 0, 9, 18, 25);
    h.clock.drain();

    const window = h.window();
    expect(window.cells[0][0]).toBe('frozen');
    expect(window.cells[0][20]).toBe('scrolled');
  });

  it('sends the frozen row as well', () => {
    command(h.port, 'setCell', 0, 5, 'heading');
    command(h.port, 'setCell', 500, 5, 'far down');
    h.clock.drain();
    command(h.port, 'freeze', 1, 0);
    command(h.port, 'setViewport', 495, 505, 0, 9);
    h.clock.drain();

    const window = h.window();
    expect(window.cells[0][5]).toBe('heading');
    expect(window.cells[500][5]).toBe('far down');
  });

  /**
   * The gap in the middle costs nothing, which is the whole reason
   * the rows are listed rather than bounded: asking for everything
   * from row 0 to the window would fetch five hundred rows to show
   * one.
   */
  it('does not send the rows between the pane and the window', () => {
    command(h.port, 'setCell', 250, 5, 'in between');
    h.clock.drain();
    command(h.port, 'freeze', 1, 0);
    command(h.port, 'setViewport', 495, 505, 0, 9);
    h.clock.drain();

    expect(h.window().cells[250]).toBeUndefined();
  });

  it('sends nothing extra when nothing is frozen', () => {
    command(h.port, 'setViewport', 495, 505, 0, 9);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setViewport', 496, 506, 0, 9);

    // The row that arrived, the row that left, and the two bounds —
    // the same four patches a scroll has always cost.
    expect(h.port.patchesFor('window')).toHaveLength(4);
  });
});
