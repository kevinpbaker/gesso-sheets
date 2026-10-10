import type { CellValue } from './Values';
import type { Workbook } from './Workbook';

/**
 * What a cell came to across a simulation: where its values fell.
 *
 * The three percentiles are what a person reads ("one year in ten it is
 * below this"); the bins are what the grid draws, a histogram small
 * enough to sit under the figure in its cell.
 */
export interface Spread {
  /** Trials in which the cell was a number. */
  readonly count: number;
  readonly mean: number;
  readonly p10: number;
  readonly p50: number;
  readonly p90: number;
  readonly min: number;
  readonly max: number;
  /** How many trials fell in each of `BINS` equal steps from `min` to `max`. */
  readonly bins: readonly number[];
}

/** How many bars a cell's histogram has. */
export const BINS = 16;

/** The most numbers a simulation will hold, all cells and trials together: 32 MB of doubles. */
export const SAMPLE_LIMIT = 4_000_000;

/**
 * A small, fast generator with a seed, so a simulation run twice with
 * the same seed draws the same numbers and says the same thing.
 * Mulberry32: 32 bits of state, a full period, and good enough for
 * drawing guesses, which is all it is for.
 */
export function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/**
 * A Monte Carlo run: the workbook recalculated `trials` times, with
 * every `NORMAL`, `UNIFORM` and `TRIANGULAR` in it drawn afresh each
 * time, and what each cell they reach came to, kept.
 *
 * On a fork of the workbook, so the one being edited never moves; and
 * a fork's `RAND()` keeps its base's value, so the only thing that
 * varies between trials is the guesses. A trial recalculates what the
 * guesses reach and nothing else, as any edit would.
 *
 * Run in pieces by `step`, each bounded by time, so it can share a
 * thread with everything else and be stopped between any two trials.
 * Only numbers are kept — a cell that is text or an error in a trial
 * is left out of that trial's count — in one array per cell, so a
 * percentile is a sort of one array.
 */
export class Simulation {
  readonly trials: number;
  readonly seed: number;
  private readonly fork: Workbook;
  private readonly keys: readonly number[];
  private readonly samples: Float64Array[];
  private readonly at = new Map<number, number>();
  private readonly spreads = new Map<number, Spread | null>();
  private ran = 0;

  /**
   * `base` should be settled; the fork is taken from it as it stands.
   * The trials asked for are cut down, if they have to be, so that the
   * cells reached times the trials stays inside `SAMPLE_LIMIT`.
   */
  constructor(base: Workbook, trials: number, seed: number) {
    this.seed = seed;
    this.fork = base.fork();
    this.fork.sampler = seeded(seed);
    this.keys = [...this.fork.uncertainReach()].sort((a, b) => a - b);
    this.trials = Math.max(1, Math.min(trials, Math.floor(SAMPLE_LIMIT / Math.max(1, this.keys.length))));
    this.samples = this.keys.map(() => new Float64Array(this.trials).fill(Number.NaN));
    this.keys.forEach((key, index) => this.at.set(key, index));
  }

  /** How many trials are done. */
  get done(): number {
    return this.ran;
  }

  get finished(): boolean {
    return this.ran >= this.trials;
  }

  /** How many cells the guesses reach: the cells there is a spread for. */
  get cells(): number {
    return this.keys.length;
  }

  /**
   * Runs trials until `milliseconds` have gone or the run is finished,
   * and says how many it ran. At least one, so a slow trial still moves.
   */
  step(milliseconds: number, now: () => number = () => performance.now()): number {
    const until = now() + milliseconds;
    let ran = 0;
    while (this.ran < this.trials && (ran === 0 || now() < until)) {
      this.fork.resample();
      this.fork.recalculate();
      for (let index = 0; index < this.keys.length; index++) {
        const value: CellValue = this.fork.valueAt(this.keys[index]);
        if (typeof value === 'number' && Number.isFinite(value)) {
          this.samples[index][this.ran] = value;
        }
      }
      this.ran++;
      ran++;
    }
    if (ran > 0) {
      this.spreads.clear();
    }
    return ran;
  }

  /** Whether a cell is one the guesses reach. */
  covers(key: number): boolean {
    return this.at.has(key);
  }

  /**
   * Where a cell's values fell over the trials run so far, or null for
   * a cell the guesses do not reach, or one that was never a number.
   * Worked out once per cell between steps.
   */
  spreadOf(key: number): Spread | null {
    if (this.spreads.has(key)) {
      return this.spreads.get(key)!;
    }
    const index = this.at.get(key);
    const spread = index === undefined ? null : spreadOf(this.samples[index], this.ran);
    this.spreads.set(key, spread);
    return spread;
  }
}

/** The spread of the first `count` samples, leaving out the trials that were not a number. */
export function spreadOf(samples: Float64Array, count: number): Spread | null {
  const values: number[] = [];
  for (let index = 0; index < count; index++) {
    const value = samples[index];
    if (!Number.isNaN(value)) {
      values.push(value);
    }
  }
  if (values.length === 0) {
    return null;
  }
  const sorted = Float64Array.from(values).sort();
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  let sum = 0;
  for (const value of sorted) {
    sum += value;
  }
  const bins = new Array<number>(BINS).fill(0);
  const width = max - min;
  for (const value of sorted) {
    const bin = width === 0 ? Math.floor(BINS / 2) : Math.min(BINS - 1, Math.floor(((value - min) / width) * BINS));
    bins[bin]++;
  }
  return {
    count: sorted.length,
    mean: sum / sorted.length,
    p10: quantile(sorted, 0.1),
    p50: quantile(sorted, 0.5),
    p90: quantile(sorted, 0.9),
    min,
    max,
    bins
  };
}

/** A quantile of sorted values, interpolated between the two either side, as `PERCENTILE` is. */
function quantile(sorted: Float64Array, fraction: number): number {
  const at = (sorted.length - 1) * fraction;
  const below = Math.floor(at);
  const above = Math.min(sorted.length - 1, below + 1);
  return sorted[below] + (sorted[above] - sorted[below]) * (at - below);
}
