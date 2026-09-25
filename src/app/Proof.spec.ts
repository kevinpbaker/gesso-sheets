import { describe, expect, it } from 'vitest';

import { cellIn } from './SheetContract';
import { attach } from './sheetHarness';

/**
 * Phase 7's claim, stated as assertions.
 *
 * "Recalculate 200,000 dependent cells while scrolling, and show that
 * the scroll re-measured nothing." The re-measuring half is the
 * engine's, and the heatmap is how a person checks it. This file is
 * the half the application owes: that two hundred thousand cells
 * really do go out of date, that they are really recalculated, that
 * the answers are right, and — the part that makes the scroll
 * possible at all — that the thread doing the recalculating is
 * answerable to the render worker the whole time it is doing it.
 *
 * The chain is deliberately the worst shape there is. Every cell reads
 * the cell before it, so nothing can be skipped, reordered or done
 * out of turn: it is a critical path two hundred thousand long, and
 * the evaluator has to walk all of it.
 */

/** The number on the button. */
const CELLS = 200_000;

/**
 * The sheet's own height, which is where the chain begins — one row
 * past the last row anybody can reach.
 */
const FIRST_ROW = 10_000;

/** Where the chain ends, which is where its answer can be read. */
function last(): [row: number, column: number] {
  return [FIRST_ROW + Math.floor(CELLS / 100), CELLS % 100];
}

describe('the proof surface', () => {
  it('makes two hundred thousand cells go out of date, and gets them all right', () => {
    const h = attach();
    const before = h.document.sheet.stats.evaluated;
    h.service.stress(CELLS);

    // The pump runs its first slice before returning, so what is owed
    // is what is still dirty plus what that slice already did. Stating
    // it as a sum rather than as `pending` is what keeps the assertion
    // about the chain instead of about the slice size.
    expect(h.document.sheet.pending + (h.document.sheet.stats.evaluated - before)).toBe(CELLS);

    h.clock.drain();

    expect(h.document.sheet.pending).toBe(0);
    expect(h.document.sheet.stats.evaluated - before).toBe(CELLS);
    // The chain runs a hundred columns at a time from its head, so its
    // last cell is the head's value plus the length of the chain.
    expect(h.document.sheet.value(...last())).toBe(1 + CELLS);
  });

  /**
   * The one that matters.
   *
   * A recalculation is only invisible if the thread can be interrupted
   * by the person scrolling. So: start the recalculation, let a single
   * slice run, and then scroll — while 200,000 cells are still owed.
   * The new window has to come back on the *next* slice, not after the
   * recalculation finishes.
   *
   * Run the other way, this is exactly the bug Phase 0 found and wrote
   * down: a recalculation that publishes only when it is done is a
   * recalculation that holds the viewport hostage, and no amount of
   * thread-count in the architecture diagram fixes it.
   */
  it('answers a scroll while it still owes two hundred thousand cells', () => {
    const h = attach();
    h.service.setViewport(0, 0, 30, 0, 10);
    h.document.sheet.setCell(5, 0, 'before');
    h.clock.drain();

    h.service.stress(CELLS);
    h.clock.tick();
    expect(h.document.sheet.pending).toBeGreaterThan(0);

    // Somebody scrolls, mid-recalculation.
    h.service.setViewport(0, 9_000, 9_030, 0, 10);
    h.clock.tick();

    expect(h.document.sheet.pending).toBeGreaterThan(0);
    expect(cellIn(h.window(), 9_000, 0)).not.toBe(undefined);
  });

  /**
   * And none of it is the person's file.
   *
   * Found by pressing the button in a browser and reloading, which is
   * the only place it could have been found: every spec passed, and
   * what the reload showed was a sheet that had quietly saved a
   * quarter of a million formulas and recalculated them on the way
   * back up. The chain lives past the end of the sheet so that a
   * snapshot, which is the sheet, does not contain it.
   */
  it('saves none of the chain, because none of it is in the sheet', () => {
    const h = attach();
    h.document.setCell(3, 1, 'kept');
    h.service.stress(CELLS);
    h.clock.drain();

    const snapshot = h.service.snapshot();

    expect(snapshot.sheets[0].cells).toEqual([{ row: 3, column: 1, input: 'kept' }]);
  });

  /**
   * And the wire stays small while it happens.
   *
   * The chain is two thousand rows below the viewport, so recalculating
   * it is not news: the window key gets nothing for cells nobody can
   * see, however many of them changed. Phase 2 proved this for 50,000
   * with a keystroke; this proves the button does not quietly undo it.
   */
  it('sends no window patches for cells nobody is looking at', () => {
    const h = attach();
    h.service.setViewport(0, 0, 30, 0, 10);
    h.clock.drain();
    h.port.clear();

    h.service.stress(CELLS);
    h.clock.drain();

    expect(h.port.patchesFor('window')).toEqual([]);
    expect(h.document.sheet.stats.evaluated).toBe(CELLS);
  });

  /**
   * Pressing it twice recalculates the chain again rather than building
   * a second one, which is what makes the button repeatable — and what
   * keeps the second press honest, since a press that built nothing new
   * and still recalculated everything is the demonstration.
   */
  it('rebuilds nothing on the second press, and still recalculates everything', () => {
    const h = attach();
    h.service.stress(CELLS);
    h.clock.drain();
    const afterFirst = h.document.sheet.stats.evaluated;

    h.service.stress(CELLS);
    h.clock.drain();

    expect(h.document.sheet.stats.evaluated - afterFirst).toBe(CELLS);
    expect(h.document.sheet.value(...last())).toBe(2 + CELLS);
  });
});
