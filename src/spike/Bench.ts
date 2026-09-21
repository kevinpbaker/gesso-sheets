import type { WireShape } from './SheetContract';

/**
 * The Phase 0 measurement, as something a machine runs.
 *
 * Scrolling by hand answers "does it feel smooth", which is the wrong
 * kind of answer for an exit criterion: the velocity is not repeatable,
 * so neither is the frame gap or the share of cells that were not
 * there yet. Each run here sets the knobs, jumps to the middle of the
 * sheet, waits for the window to settle, then scrolls at a fixed
 * velocity for a fixed time and prints one line of JSON.
 *
 * It runs in the render worker, which is the thread whose frame gap the
 * exit criterion is about, and reports over `console.log` because
 * `captureConsole` already forwards that to the page for a harness to
 * read.
 */

export interface BenchRun {
  readonly label: string;
  readonly axis: 'x' | 'y';
  /** Logical pixels per second. */
  readonly speed: number;
  readonly mountRows: number;
  readonly mountColumns: number;
  readonly fetchRows: number;
  readonly fetchColumns: number;
  readonly shape: WireShape;
  /** Milliseconds the application worker burns before each publish. */
  readonly busy: number;
}

export interface BenchKnobs {
  setMountBand(rows: number, columns: number): void;
  setFetchBand(rows: number, columns: number): void;
  setShape(shape: WireShape): void;
  setBusy(ms: number): void;
  scrollTo(x: number, y: number): void;
  report(line: string): void;
  done(): void;
}

/** What the driver reads off each frame. */
export interface BenchSample {
  /** Gap since the previous frame, on the render worker's own clock. */
  readonly gapMs: number;
  /**
   * What the frame cost the render worker.
   *
   * Reported alongside the gap because the two answer different
   * questions, and in headless Chrome only one of them is trustworthy:
   * the gap is whatever the compositor chose to schedule, while this
   * is work, and work under 16.7 ms is what "60fps is available"
   * actually means.
   */
  readonly durationMs: number;
  /**
   * The frame's phases, summed over the run.
   *
   * Reported because the horizontal axis came out slower than the
   * vertical with a *lower* frame cost and a *smaller* wire, which
   * rules out both of the obvious answers and leaves the question of
   * where the time goes. A phase breakdown is the next place to look
   * before guessing.
   */
  readonly phases: Readonly<Record<string, number>>;
  /** Share of visible cells with no value yet, 0..1. */
  readonly missed: number;
  readonly nodes: number;
  readonly measured: number;
  /** Cumulative counters from the application worker. */
  readonly publishes: number;
  readonly patches: number;
  readonly bytes: number;
  readonly cells: number;
}

/**
 * The settle is a condition, not a delay.
 *
 * A run that started measuring while the previous run's publishes were
 * still draining reported the tail of that run as this one's misses —
 * which is how an identical configuration came out at 0% in one slot
 * and 67% in another. Waiting for the window to be fully covered, and
 * for a floor under it, makes each run measure scrolling rather than
 * arrival.
 */
const SETTLE_FLOOR_MS = 400;
const SETTLE_CEILING_MS = 4000;
const SWEEP_MS = 2500;

/**
 * Twenty runs, in four groups, each answering one question.
 *
 * The bands are swept one at a time with the other at zero, because
 * the point is to price them against each other: a mount band costs a
 * node on every frame it is mounted for, a fetch band costs only cells
 * on the wire, and if they buy the same coverage the cheap one wins.
 */
export function benchMatrix(): BenchRun[] {
  const base = {
    mountRows: 2,
    mountColumns: 1,
    fetchRows: 0,
    fetchColumns: 0,
    shape: 'keyed' as WireShape,
    busy: 0
  };
  const runs: BenchRun[] = [];

  // 1. The headline: both axes, an ordinary scroll and a fling.
  for (const axis of ['y', 'x'] as const) {
    for (const speed of [3000, 9000]) {
      runs.push({ ...base, label: `axis-${axis}-${speed}`, axis, speed });
    }
  }

  // 2. The band, one axis at a time, at the velocity that hurts.
  for (const band of [0, 2, 4, 8, 16, 24]) {
    runs.push({
      ...base,
      label: `mount-band-${band}`,
      axis: 'y',
      speed: 9000,
      mountRows: band,
      mountColumns: Math.max(1, band >> 2)
    });
  }
  for (const band of [0, 2, 4, 8, 16, 24]) {
    runs.push({
      ...base,
      label: `fetch-band-${band}`,
      axis: 'y',
      speed: 9000,
      mountRows: 0,
      mountColumns: 0,
      fetchRows: band,
      fetchColumns: Math.max(1, band >> 2)
    });
  }

  // 3. The wire shape, at the same band on both.
  for (const shape of ['keyed', 'rows'] as WireShape[]) {
    runs.push({ ...base, label: `wire-${shape}`, axis: 'y', speed: 3000, fetchRows: 16, fetchColumns: 4, shape });
  }

  // 4. An application thread with something to do.
  for (const busy of [8, 30]) {
    runs.push({ ...base, label: `recalc-${busy}ms`, axis: 'y', speed: 3000, fetchRows: 16, fetchColumns: 4, busy });
  }

  // 5. Which axis the wire should be keyed by. `keyed` is row-major,
  //    so a horizontal scroll touches every mounted row; `columnKeyed`
  //    is the same map transposed. Both axes, so neither shape can win
  //    by being measured only on the scroll it suits.
  for (const shape of ['keyed', 'columnKeyed'] as WireShape[]) {
    for (const axis of ['y', 'x'] as const) {
      runs.push({
        ...base,
        label: `major-${shape === 'keyed' ? 'row' : 'column'}-${axis}`,
        axis,
        speed: 9000,
        fetchRows: 16,
        fetchColumns: 4,
        shape
      });
    }
  }

  // 6. The headline again, with the band the runs above chose. This is
  //    the group the exit criterion is read off: group 1 measured both
  //    axes with a band that turned out to be the wrong one.
  for (const axis of ['y', 'x'] as const) {
    for (const speed of [3000, 9000]) {
      runs.push({ ...base, label: `confirm-${axis}-${speed}`, axis, speed, fetchRows: 16, fetchColumns: 4 });
    }
  }

  return runs;
}

type Phase = 'apply' | 'settle' | 'sweep';

export class BenchDriver {
  private index = 0;
  private phase: Phase = 'apply';
  private until = 0;
  private ceiling = 0;
  private direction = 1;
  private gaps: number[] = [];
  private durations: number[] = [];
  private phaseTotals = new Map<string, number>();
  private missedFrames = 0;
  private worstMiss = 0;
  private peakNodes = 0;
  private peakMeasured = 0;
  private start: BenchSample | null = null;
  private finished = false;
  /** Where the driver has scrolled the sheet to; it owns this, not the sheet. */
  private readonly position = { x: 0, y: 0 };

  constructor(
    private readonly runs: readonly BenchRun[],
    private readonly knobs: BenchKnobs,
    private readonly geometry: { rowHeight: number; columnWidth: number; rowCount: number; columnCount: number }
  ) {}

  frame(at: number, sample: BenchSample): void {
    if (this.finished) {
      return;
    }
    const run = this.runs[this.index];
    if (run === undefined) {
      this.finished = true;
      this.knobs.done();
      return;
    }

    if (this.phase === 'apply') {
      this.knobs.setMountBand(run.mountRows, run.mountColumns);
      this.knobs.setFetchBand(run.fetchRows, run.fetchColumns);
      this.knobs.setShape(run.shape);
      this.knobs.setBusy(run.busy);
      // The middle of the sheet, so a band applies on both sides of
      // both axes and no run is measuring a clamped edge.
      this.position.x = Math.floor(this.geometry.columnCount / 2) * this.geometry.columnWidth;
      this.position.y = Math.floor(this.geometry.rowCount / 2) * this.geometry.rowHeight;
      this.knobs.scrollTo(this.position.x, this.position.y);
      this.phase = 'settle';
      this.until = at + SETTLE_FLOOR_MS;
      this.ceiling = at + SETTLE_CEILING_MS;
      return;
    }

    if (this.phase === 'settle') {
      if (at >= this.until && (sample.missed === 0 || at >= this.ceiling)) {
        this.phase = 'sweep';
        this.until = at + SWEEP_MS;
        this.direction = 1;
        this.gaps = [];
        this.durations = [];
        this.phaseTotals = new Map();
        this.missedFrames = 0;
        this.worstMiss = 0;
        this.peakNodes = 0;
        this.peakMeasured = 0;
        this.start = sample;
      }
      return;
    }

    // Sweeping. The step is velocity × the gap this frame actually
    // took, so a dropped frame moves further rather than slowing the
    // sweep down and flattering the result.
    const step = (run.speed * sample.gapMs) / 1000;
    const extent =
      run.axis === 'y'
        ? this.geometry.rowCount * this.geometry.rowHeight
        : this.geometry.columnCount * this.geometry.columnWidth;
    const viewport = run.axis === 'y' ? 800 : 1200;
    const limit = Math.max(0, extent - viewport);
    let next = this.position[run.axis] + this.direction * step;
    // The horizontal axis is only a hundred columns wide, so a fling
    // crosses the whole sheet in about a second; bouncing keeps both
    // axes' runs the same length, and a reversal is a harder case for
    // a window than a steady scroll anyway.
    if (next <= 0 || next >= limit) {
      next = Math.min(limit, Math.max(0, next));
      this.direction = -this.direction;
    }
    this.position[run.axis] = next;
    this.knobs.scrollTo(this.position.x, this.position.y);

    this.gaps.push(sample.gapMs);
    this.durations.push(sample.durationMs);
    for (const [phase, ms] of Object.entries(sample.phases)) {
      this.phaseTotals.set(phase, (this.phaseTotals.get(phase) ?? 0) + ms);
    }
    if (sample.missed > 0) {
      this.missedFrames++;
      this.worstMiss = Math.max(this.worstMiss, sample.missed);
    }
    this.peakNodes = Math.max(this.peakNodes, sample.nodes);
    this.peakMeasured = Math.max(this.peakMeasured, sample.measured);

    if (at >= this.until) {
      this.finish(run, sample);
    }
  }

  private finish(run: BenchRun, sample: BenchSample): void {
    const gaps = [...this.gaps].sort((a, b) => a - b);
    const durations = [...this.durations].sort((a, b) => a - b);
    const at = (fraction: number) => gaps[Math.min(gaps.length - 1, Math.floor(gaps.length * fraction))] ?? 0;
    const cost = (fraction: number) =>
      durations[Math.min(durations.length - 1, Math.floor(durations.length * fraction))] ?? 0;
    const started = this.start;
    const publishes = started === null ? 0 : sample.publishes - started.publishes;
    this.knobs.report(
      JSON.stringify({
        run: run.label,
        axis: run.axis,
        speed: run.speed,
        mountBand: [run.mountRows, run.mountColumns],
        fetchBand: [run.fetchRows, run.fetchColumns],
        shape: run.shape,
        busy: run.busy,
        frames: gaps.length,
        fps: gaps.length === 0 ? 0 : round(1000 / mean(gaps)),
        gapMs: round(mean(gaps)),
        gapP95Ms: round(at(0.95)),
        costMs: round(mean(durations)),
        costP95Ms: round(cost(0.95)),
        costWorstMs: round(durations[durations.length - 1] ?? 0),
        missFramesPct: gaps.length === 0 ? 0 : round((100 * this.missedFrames) / gaps.length),
        worstMissPct: round(100 * this.worstMiss),
        phases: Object.fromEntries(
          [...this.phaseTotals].map(([phase, total]) => [phase, round(total / Math.max(1, this.durations.length))])
        ),
        peakNodes: this.peakNodes,
        peakMeasured: this.peakMeasured,
        publishes,
        cells: sample.cells,
        patchesPerPublish: publishes === 0 || started === null ? 0 : round((sample.patches - started.patches) / publishes),
        bytesPerPublish: publishes === 0 || started === null ? 0 : Math.round((sample.bytes - started.bytes) / publishes)
      })
    );
    this.index++;
    this.phase = 'apply';
  }
}

function mean(values: readonly number[]): number {
  if (values.length === 0) {
    return 0;
  }
  let total = 0;
  for (const value of values) {
    total += value;
  }
  return total / values.length;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/** The bench runs when the page url carries `?bench`; see `still.ts` in Gesso. */
export function isBench(): boolean {
  if (typeof document === 'undefined') {
    return typeof self !== 'undefined' && (self as { name?: string }).name === 'bench';
  }
  try {
    return new URLSearchParams(location.search).has('bench');
  } catch {
    return false;
  }
}

export const BENCH_PREFIX = 'PHASE0 ';
export const BENCH_DONE = 'PHASE0-DONE';
