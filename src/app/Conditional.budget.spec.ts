import { beforeEach, describe, expect, it } from 'vitest';


import { SCALE_STEPS } from '../sheet/Conditional';
import { cellIn } from './SheetContract';
import { attach, RecordingPort, type Harness } from './sheetHarness';

/**
 * Phase 14's exit criterion, in Phase 2's shape.
 *
 * "A rule over the whole sheet costs the viewport." The numbers below
 * are what a `postMessage` would carry, and they are exact: a count
 * fails the build with a number when a change makes the wire carry
 * the rule's range instead of the window.
 *
 * The roadmap asked for a *measurement* rather than a prediction
 * about whether re-resolving the window's rules on every publish is
 * affordable. The evaluation counts here are that measurement.
 */

function command(port: RecordingPort, name: string, ...args: unknown[]): void {
  port.send({ type: 'channel:command', command: name, payload: args[0], rest: args.slice(1) });
}

describe('a rule over a million cells', () => {
  let h: Harness;

  beforeEach(() => {
    h = attach();
  });

  /** A1:XFD1048576 — the whole sheet, which is what people actually select. */
  const WHOLE_SHEET = { firstRow: 0, firstColumn: 0, lastRow: 1_048_575, lastColumn: 999 };

  function scaleOverEverything(): void {
    h.service.setSelection(WHOLE_SHEET.firstRow, WHOLE_SHEET.firstColumn, WHOLE_SHEET.lastRow, WHOLE_SHEET.lastColumn);
    command(h.port, 'addConditional', { test: null, scale: { from: '#ffffff', to: '#4285f4' } });
    h.clock.drain();
  }

  it('costs the window and not the range', () => {
    // A thousand numbers, which is what a scale spreads between.
    for (let row = 0; row < 1_000; row++) {
      h.document.sheet.setCell(row, 0, String(row));
    }
    h.document.sheet.recalculate();

    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    scaleOverEverything();
    h.port.clear();

    // One row of scroll. Thirty rows in view, five columns; the rule
    // covers every cell of every one of them.
    command(h.port, 'setViewport', 0, 1, 30, 0, 4);
    h.clock.drain();

    const formats = h.port.patchesFor('formats');
    // The row that entered and the row that left, and the window's
    // own bounds. Not a thousand, and not a million.
    expect(formats.filter(patch => patch.path[0] === 'cells')).toHaveLength(2);
    expect(h.port.patchesFor('palette')).toEqual([]);
  });

  /**
   * The scale's colours are quantised, so scrolling asks for ones the
   * palette already holds. Without that the palette would grow by a
   * row of colours per scrolled row, for ever.
   */
  it('asks the palette for no more than the scale has steps', () => {
    for (let row = 0; row < 1_000; row++) {
      h.document.sheet.setCell(row, 0, String(row));
    }
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    scaleOverEverything();

    const sized = () => {
      let entries = 0;
      h.service.palette.subscribe(palette => (entries = palette.entries.length)).unsubscribe();
      return entries;
    };
    const plain = 1;

    // The whole thousand rows, which is every step of the scale
    // several times over.
    for (let top = 0; top < 1_000; top += 10) {
      command(h.port, 'setViewport', 0, top, top + 29, 0, 4);
      h.clock.drain();
    }
    expect(sized() - plain).toBeLessThanOrEqual(SCALE_STEPS);

    // And having seen them all, scrolling back over the same ground
    // asks for nothing new.
    const settled = sized();
    h.port.clear();
    for (let top = 0; top < 1_000; top += 10) {
      command(h.port, 'setViewport', 0, top, top + 29, 0, 4);
      h.clock.drain();
    }
    expect(sized()).toBe(settled);
    expect(h.port.patchesFor('palette')).toEqual([]);
  });

  /**
   * The extent is what a scale needs and a cell cannot answer. It is
   * cached, so scrolling never pays for it — which is the whole of
   * why a rule over a million cells can be scrolled at all.
   */
  it('works the extent out once and not again for a scroll', () => {
    for (let row = 0; row < 1_000; row++) {
      h.document.sheet.setCell(row, 0, String(row));
    }
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    scaleOverEverything();

    const scans = h.service.painterStats.scans;
    for (let top = 0; top < 200; top += 10) {
      command(h.port, 'setViewport', 0, top, top + 29, 0, 4);
      h.clock.drain();
    }
    expect(h.service.painterStats.scans).toBe(scans);
  });

  /** An edit can move a minimum, so an edit does pay for it — once. */
  it('works it out again after an edit, and once', () => {
    h.document.sheet.setCell(0, 0, '1');
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    scaleOverEverything();

    const scans = h.service.painterStats.scans;
    command(h.port, 'setCell', 5, 0, '99');
    h.clock.drain();
    expect(h.service.painterStats.scans).toBe(scans + 1);
  });

  /**
   * The measurement the roadmap asked for instead of a prediction.
   *
   * A formula rule is the expensive kind: it is evaluated per visible
   * cell per publish. Thirty rows by five columns is a hundred and
   * fifty, which is the same order as a slice of the recalc pump the
   * worker already runs between publishes — and it is a *number*
   * here rather than an expectation.
   */
  it('evaluates a formula rule once per visible cell per publish', () => {
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    h.service.setSelection(0, 0, 99, 4);
    command(h.port, 'addConditional', {
      test: { kind: 'formula', input: '=A1<0' },
      paint: { fill: '#fce8e6' }
    });
    h.clock.drain();

    const before = h.service.painterStats.evaluations;
    command(h.port, 'setViewport', 0, 1, 30, 0, 4);
    h.clock.drain();
    const perPublish = h.service.painterStats.evaluations - before;

    // Thirty rows and five columns, once each, for the one publish
    // the scroll caused.
    expect(perPublish).toBe(30 * 5);
  });

  /** A sheet with no rules pays for none of this. */
  it('costs a sheet with no rules nothing at all', () => {
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    for (let top = 0; top < 200; top += 10) {
      command(h.port, 'setViewport', 0, top, top + 29, 0, 4);
      h.clock.drain();
    }
    expect(h.service.painterStats).toMatchObject({ scans: 0, scanned: 0, evaluations: 0 });
  });

  it('still publishes the cells themselves as it always did', () => {
    h.document.sheet.setCell(3, 1, '42');
    h.document.sheet.recalculate();
    command(h.port, 'setViewport', 0, 0, 29, 0, 4);
    h.clock.drain();
    scaleOverEverything();
    expect(cellIn(h.window(), 3, 1)).toBe('42');
  });
});
