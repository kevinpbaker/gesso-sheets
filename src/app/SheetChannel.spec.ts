import { beforeEach, describe, expect, it } from 'vitest';

import { applyPatches, provide, type Patch } from 'gesso-framework';

import { cellIn, Sheet, type SheetView, type SheetWindow } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService, type Schedule } from './SheetService';

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

/** A `ChannelPort` that keeps what was posted. */
class RecordingPort {
  onmessage: ((event: { data: unknown }) => void) | null = null;
  /** Since the last `clear`, which is what a budget is measured over. */
  readonly posted: unknown[] = [];
  /** Everything, ever, which is what replaying the view needs. */
  private readonly history: unknown[] = [];

  postMessage(message: unknown): void {
    this.posted.push(message);
    this.history.push(message);
  }

  /** The render thread's end, sending a message in. */
  send(message: unknown): void {
    this.onmessage?.({ data: message });
  }

  patches(): Patch[] {
    return this.collect(this.posted);
  }

  /** Every patch for a key since the channel opened. */
  allPatchesFor(projection: keyof SheetView): Patch[] {
    return this.collect(this.history).filter(patch => patch.projection === projection);
  }

  private collect(messages: readonly unknown[]): Patch[] {
    const all: Patch[] = [];
    for (const message of messages) {
      const typed = message as { type?: string; patches?: Patch[] };
      if (typed.type === 'channel:patch') {
        all.push(...(typed.patches ?? []));
      }
      if (typed.type === 'channel:error') {
        throw new Error(`the channel reported: ${String((message as { message: string }).message)}`);
      }
    }
    return all;
  }

  /** Patches for one view key, which is how a budget is stated. */
  patchesFor(projection: keyof SheetView): Patch[] {
    return this.patches().filter(patch => patch.projection === projection);
  }

  clear(): void {
    this.posted.length = 0;
  }
}

/** A scheduler the spec advances by hand, so slices are deterministic. */
class ManualSchedule {
  private queue: (() => void)[] = [];

  readonly schedule: Schedule = run => {
    this.queue.push(run);
  };

  get pending(): number {
    return this.queue.length;
  }

  /** Runs one slice. */
  tick(): boolean {
    const next = this.queue.shift();
    next?.();
    return next !== undefined;
  }

  /** Runs slices until the pump stops. */
  drain(): number {
    let ran = 0;
    while (this.tick()) {
      ran++;
      if (ran > 10_000) {
        throw new Error('the pump never finished');
      }
    }
    return ran;
  }
}

interface Harness {
  readonly port: RecordingPort;
  readonly service: SheetService;
  readonly document: SheetDocument;
  readonly clock: ManualSchedule;
  window(): SheetWindow;
}

function attach(options: { budget?: number } = {}): Harness {
  const document = new SheetDocument();
  const clock = new ManualSchedule();
  const service = new SheetService(document, { schedule: clock.schedule, budget: options.budget ?? 2_000 });
  const port = new RecordingPort();
  provide(Sheet, sheetChannel(service).source as never, port);
  // The replica asks rather than waiting to be pushed to; this is that.
  port.send({ type: 'channel:sync' });
  return {
    port,
    service,
    document,
    clock,
    window: () => rebuild(port)
  };
}

/**
 * The window as the render thread would hold it.
 *
 * Replayed with the framework's own `applyPatches` rather than
 * something written for the spec: the claim being tested is that these
 * patches reconstruct the window, and a hand-rolled applier would only
 * test that they reconstruct it *the way the spec imagined*. It also
 * has to start from the token's initial value, which is where a
 * replica starts.
 */
function rebuild(port: RecordingPort): SheetWindow {
  return applyPatches(Sheet.initial.window, port.allPatchesFor('window')) as SheetWindow;
}

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
    command(h.port, 'setViewport', 0, 1, 0, 1);

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
    command(h.port, 'setViewport', 0, 29, 0, 4);
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
    command(h.port, 'setViewport', 0, 29, 0, 4);
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
    command(h.port, 'setViewport', 2_000, 2_029, 10, 14);
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
    command(h.port, 'setViewport', 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setViewport', 1, 30, 0, 4);
    h.clock.drain();

    const patches = h.port.patchesFor('window');
    // Two bounds moved, one row of five cells arrived, one was deleted.
    expect(patches.filter(patch => patch.op === 'set' && patch.path[0] === 'cells')).toHaveLength(1);
    expect(patches.filter(patch => patch.op === 'delete')).toHaveLength(1);
    expect(patches.filter(patch => patch.path[0] === 'firstRow' || patch.path[0] === 'lastRow')).toHaveLength(2);
    expect(patches).toHaveLength(4);
  });

  it('keeps the keys apart, so a keystroke does not walk the geometry', () => {
    command(h.port, 'setViewport', 0, 9, 0, 4);
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
    command(h.port, 'setViewport', 0, 4, 0, 4);
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
    command(h.port, 'setViewport', 0, 29, 0, 4);
    h.clock.drain();
    h.port.clear();

    command(h.port, 'setCell', 0, 0, '2');
    // One slice done, nineteen thousand cells still to go.
    h.clock.tick();
    expect(h.service.recalculating).toBe(true);
    expect(h.document.sheet.pending).toBeGreaterThan(15_000);

    // The scroll arrives in the gap the pump left, and is served.
    h.port.clear();
    command(h.port, 'setViewport', 5_000, 5_029, 0, 4);

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
    command(h.port, 'setViewport', 4_990, 5_000, 0, 0);
    h.clock.drain();

    command(h.port, 'setCell', 0, 0, '2');
    h.clock.tick();
    command(h.port, 'setViewport', 4_990, 5_000, 0, 0);
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
    command(h.port, 'setViewport', 0, 9, 0, 0);
    h.clock.drain();

    const before = h.service.stats.slices;
    command(h.port, 'setCell', 0, 0, '2');
    const slices = h.clock.drain() + 1;

    expect(slices).toBe(5);
    expect(h.service.stats.slices - before).toBe(5);
  });
});
