import { beforeEach, describe, expect, it } from 'vitest';

import { type Patch } from 'gesso-framework';

import { cellIn, type SheetView, type SheetWindow } from './SheetContract';
import { attach, RecordingPort, type Harness } from './sheetHarness';

/**
 * Patch budgets — the exit criterion for Phase 2.
 *
 * "One keystroke in a cell with 50,000 dependents should emit patches
 * proportional to the visible window, not to the dependents. That
 * single assertion is the whole thesis." So this spec drives the real
 * thing: `provide` from the framework, over a port that records what
 * it is given, with the real differ in between. Nothing is mocked, and
 * the numbers below are the numbers that would cross a `postMessage`.
 *
 * Counts again, not timings, for the reason Phase 1's budgets are
 * counts: a count fails the build with a number when a change makes
 * the wire carry the sheet instead of the window.
 */

function command(port: RecordingPort, name: string, ...args: unknown[]): void {
  port.send({ type: 'channel:command', command: name, payload: args[0], rest: args.slice(1) });
}

describe('the sheet channel', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach();
  });

  it('sends nothing at all until somebody is looking', () => {
    // The viewport is empty until the render worker says otherwise, so
    // the window it publishes equals the token's initial value and the
    // differ has nothing to say about it.
    expect(h.port.patchesFor('window')).toEqual([]);
  });

  it('fills the window the render worker asks for', () => {
    command(h.port, 'setCell', 0, 0, '1');
    command(h.port, 'setCell', 0, 1, '=A1+1');
    h.clock.drain();
    command(h.port, 'setViewport', 0, 0, 1, 0, 1);

    expect(cellIn(h.window(), 0, 0)).toBe('1');
    expect(cellIn(h.window(), 0, 1)).toBe('2');
  });

  /**
   * The thesis, in one number.
   *
   * A column of 50,000 formulas, a viewport showing thirty rows of it,
   * and one keystroke at the top. Every one of the 50,000 is
   * recalculated — Phase 1's budget spec asserts exactly that — and
   * what crosses the barrier is the handful of cells somebody can see.
   */
  it('emits patches for what is visible, not for what was recalculated', () => {
    const DEPENDENTS = 50_000;
    h.document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= DEPENDENTS; row++) {
      h.document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    h.document.sheet.recalculate();

    // Thirty rows and five columns in view: the shape of a real screen.
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    const before = h.document.sheet.stats.evaluated;
    command(h.port, 'setCell', 0, 0, '2');
    h.clock.drain();

    // The whole column was recomputed.
    expect(h.document.sheet.stats.evaluated - before).toBe(DEPENDENTS);

    // And this is what it cost to say so. Thirty visible cells changed
    // — rows 0 to 29 of column A — so thirty patches, one per cell,
    // and not fifty thousand.
    const patches = h.port.patchesFor('window');
    expect(patches).toHaveLength(30);
    expect(patches.every(patch => patch.path[0] === 'cells')).toBe(true);
    expect(cellIn(h.window(), 29, 0)).toBe('31');
  });

  it('emits one patch when one visible cell changes', () => {
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 5, 2, 'hello');
    h.clock.drain();

    const patches = h.port.patchesFor('window');
    expect(patches).toHaveLength(1);
    expect(patches[0]).toMatchObject({ op: 'set', path: ['cells', '5', '2'], value: 'hello' });
  });

  /**
   * The same edit with the dependents off screen costs one patch, not
   * thirty — the wire is priced by what is visible and by nothing
   * else, which is the claim stated the other way round.
   */
  it('emits nothing for a recalculation nobody can see', () => {
    h.document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= 5_000; row++) {
      h.document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    h.document.sheet.recalculate();

    // A viewport somewhere else entirely.
    command(h.port, 'setViewport', 0, 2_000, 2_029, 10, 14);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, '2');
    h.clock.drain();

    expect(h.port.patchesFor('window')).toEqual([]);
  });

  /**
   * A scroll of one row is the operation a sheet does most, and the
   * shape of the projection is what decides its price. Row-major
   * arrays would resend the window cell by cell here; keyed by where a
   * cell is, the row that entered and the row that left are the only
   * things that moved.
   */
  it('prices a scroll of one row at the row that entered and the row that left', () => {
    for (let row = 0; row < 60; row++) {
      for (let column = 0; column < 5; column++) {
        h.document.sheet.setCell(row, column, String(row * 10 + column));
      }
    }
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setViewport', 0, 1, 30, 0, 4);
    h.clock.drain();

    const patches = h.port.patchesFor('window');
    // Two bounds moved, one row of five cells arrived, one was deleted.
    expect(patches.filter(patch => patch.op === 'set' && patch.path[0] === 'cells')).toHaveLength(1);
    expect(patches.filter(patch => patch.op === 'delete')).toHaveLength(1);
    expect(patches.filter(patch => patch.path[0] === 'firstRow' || patch.path[0] === 'lastRow')).toHaveLength(2);
    expect(patches).toHaveLength(4);
  });

  it('keeps the keys apart, so a keystroke does not walk the geometry', () => {
    command(h.port, 'setViewport', 0, 0, 9, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, '=1+1');
    h.clock.drain();

    expect(h.port.patchesFor('geometry')).toEqual([]);
    expect(h.port.patchesFor('selection')).toEqual([]);
    expect(h.port.patchesFor('window')).toHaveLength(1);
  });

  it('publishes only plain data, which the barrier checks on the first value', () => {
    // `provide` runs `requirePlainData` once per key and reports over
    // the port rather than throwing, so a rich value would surface as
    // a 'channel:error' — which `patches()` turns into a failure.
    command(h.port, 'setViewport', 0, 0, 4, 0, 4);
    command(h.port, 'setCell', 0, 0, '=1/0');
    h.clock.drain();
    expect(() => h.port.patches()).not.toThrow();
  });
});

describe('a recalculation while the viewport moves', () => {
  /**
   * The second spec Phase 0 asked this phase for.
   *
   * Phase 0 measured a recalc that holds the thread: thirty
   * milliseconds of it leaves the sheet blank in 89% of frames, while
   * the render worker goes on scrolling at 60fps and asking for
   * windows that nobody is free to serve. The engine is sliced and the
   * pump hands the thread back between slices, so the test is whether
   * a viewport that moves mid-recalc is answered *before* the
   * arithmetic finishes.
   */
  it('answers a scroll before the recalculation it interrupted', () => {
    const h = attach({ budget: 1_000 });
    h.document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= 20_000; row++) {
      h.document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, '2');
    // One slice done, nineteen thousand cells still to go.
    h.clock.tick();
    expect(h.service.recalculating).toBe(true);
    expect(h.document.sheet.pending).toBeGreaterThan(15_000);

    // The scroll arrives in the gap the pump left, and is served.
    h.port.clear();
    command(h.port, 'setViewport', 0, 5_000, 5_029, 0, 4);

    const answered = h.port.patchesFor('window');
    expect(answered.length).toBeGreaterThan(0);
    expect(h.service.recalculating).toBe(true);
    expect(h.document.sheet.pending).toBeGreaterThan(0);

    // And the values it served are the ones the recalc had reached,
    // not blanks: everything up to the first slice is current.
    expect(cellIn(h.window(), 5_000, 0)).not.toBeNull();
  });

  it('finishes the recalculation the scroll interrupted', () => {
    const h = attach({ budget: 1_000 });
    h.document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= 5_000; row++) {
      h.document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 4_990, 5_000, 0, 0);
    h.clock.drain();

    command(h.port, 'setCell', 0, 0, '2');
    h.clock.tick();
    command(h.port, 'setViewport', 0, 4_990, 5_000, 0, 0);
    h.clock.drain();

    expect(h.document.sheet.pending).toBe(0);
    expect(h.service.recalculating).toBe(false);
    expect(cellIn(h.window(), 5_000, 0)).toBe('5002');
  });

  it('takes more than one slice to do a large graph, so the thread is handed back', () => {
    const h = attach({ budget: 1_000 });
    h.document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= 5_000; row++) {
      h.document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 9, 0, 0);
    h.clock.drain();

    const before = h.service.stats.slices;
    command(h.port, 'setCell', 0, 0, '2');
    const slices = h.clock.drain() + 1;

    expect(slices).toBe(5);
    expect(h.service.stats.slices - before).toBe(5);
  });
});

/**
 * Phase 13's exit criterion, in Phase 2's shape.
 *
 * "A formula on Sheet 1 depending on 50,000 cells on Sheet 2
 * publishes **nothing** while Sheet 2 is not the sheet in view."
 * Cross-sheet references are exactly where a naive implementation
 * starts publishing the whole workbook, and the failure is invisible
 * until the workbook is large.
 *
 * The counts are what a `postMessage` would carry. The recalculation
 * is real and is asserted separately, so a spec that passed because
 * nothing happened would fail on the line above it.
 */
describe('a workbook of several sheets', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach();
  });

  const DEPENDENTS = 50_000;

  /**
   * A column of 50,000 chained formulas on Sheet 2, hanging from a
   * cell on Sheet 1 and read back by another one.
   *
   * Built this way round so the whole thing can be set off by a
   * keystroke through the real command path, on the sheet in view.
   * That is also the shape of the case worth worrying about: nobody
   * edits a sheet they are not looking at, but plenty of sheets feed
   * one.
   */
  function acrossTheBook(): void {
    command(h.port, 'addSheet');
    const second = h.document.sheet;
    second.setCell(0, 0, '=Sheet1!A1+1');
    for (let row = 1; row <= DEPENDENTS - 1; row++) {
      second.setCell(row, 0, `=A${row}+1`);
    }
    h.document.activate(0);
    h.document.sheet.setCell(0, 1, `=Sheet2!A${DEPENDENTS}`);
    h.document.book.recalculate();
  }

  it('publishes nothing for a recalculation on a sheet nobody is looking at', () => {
    acrossTheBook();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    const before = h.document.book.stats.evaluated;
    // One keystroke, on the sheet in view, with 50,000 cells on the
    // sheet that is not.
    command(h.port, 'setCell', 0, 0, '2');
    h.clock.drain();

    // The work really happened: the whole column on Sheet 2, and the
    // cell on Sheet 1 that reads the far end of it.
    expect(h.document.book.stats.evaluated - before).toBe(DEPENDENTS + 1);

    // And the wire carried two cells — the one typed into and the one
    // whose answer changed. Not fifty thousand, and nothing at all
    // for the sheet that did the work.
    const patches = h.port.patchesFor('window');
    expect(patches).toHaveLength(2);
    expect(patches.every(patch => patch.path[0] === 'cells')).toBe(true);
    expect(cellIn(h.window(), 0, 1)).toBe('50002');
    expect(h.port.patchesFor('geometry')).toEqual([]);
    expect(h.port.patchesFor('sheets')).toEqual([]);
  });

  it('sends the tab strip once, and not again for an edit', () => {
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    command(h.port, 'addSheet');
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, 'hello');
    h.clock.drain();
    expect(h.port.patchesFor('sheets')).toEqual([]);
  });

  it('prices a tab change at the window, not at the workbook', () => {
    command(h.port, 'addSheet');
    command(h.port, 'setViewport', 1, 0, 29, 0, 4);
    h.clock.drain();
    // Something to see on each sheet, so the change is not free by
    // being a change between two empty grids.
    h.document.sheet.setCell(0, 0, 'on two');
    h.document.activate(0);
    h.document.sheet.setCell(0, 0, 'on one');
    h.document.activate(1);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();

    expect(cellIn(h.window(), 0, 0)).toBe('on one');
    // One cell differed between the two sheets, so one patch.
    expect(h.port.patchesFor('window')).toHaveLength(1);
    expect(h.port.patchesFor('sheets')).toHaveLength(1);
  });
});
