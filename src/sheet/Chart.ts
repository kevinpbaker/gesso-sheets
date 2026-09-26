import type { RangeRef } from './A1';
import type { Series } from './Series';

/**
 * What a chart is, and the arithmetic of drawing one.
 *
 * Model and geometry in one file because they are one idea: a chart
 * kind is not a name, it is a decision about what the axes mean, and
 * keeping the two apart would put `kind === 'bar'` in two places that
 * could disagree.
 *
 * Nothing here draws anything. It is the same rule every file in this
 * directory is under — no framework, no viewport, no canvas — and it
 * is what lets the tick selection and the stacking be asserted as
 * numbers instead of inspected as pictures.
 */

/**
 * The seven kinds, and what separates them.
 *
 * - `line`, `area` and `scatter` plot a value against a position.
 * - `column` and `bar` are the same chart on its side, which is why
 *   they are two names and one code path with an axis swap.
 * - `stacked` is `column` with the series piled up rather than side
 *   by side, so its value axis runs to the total and not to the
 *   largest series.
 * - `pie` has no axes at all, reads one series, and is the only kind
 *   whose geometry is angles.
 */
export type ChartKind = 'line' | 'area' | 'column' | 'bar' | 'stacked' | 'pie' | 'scatter';

/**
 * Where a chart sits, in the sheet's own pixels.
 *
 * **Pixels from the grid's origin, not an anchor cell**, and that is
 * a choice with a cost worth writing down. A chart anchored to `D4`
 * would follow the cell when a column above it is widened or a row is
 * inserted; this one stays where it was put and the cells move under
 * it. Anchoring is what a spreadsheet does in the end, and it is a
 * mapping this phase does not need to get charts on the screen: the
 * coordinate space is the thing that has to exist first, and an
 * anchor is a translation into it.
 */
export interface ChartPlacement {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface Chart {
  /**
   * Stable for the life of the chart, and never reused.
   *
   * The render worker holds selection and a drag against this, and a
   * position in a list would move under both when a chart in front of
   * it is deleted.
   */
  readonly id: number;
  readonly kind: ChartKind;
  readonly title: string;
  readonly range: RangeRef;
  readonly place: ChartPlacement;
  readonly legend: boolean;
}

/** The smallest a chart is allowed to be dragged. */
export const MIN_CHART_WIDTH = 160;
export const MIN_CHART_HEIGHT = 120;

/** What a chart is given when it is inserted and nobody has said. */
export const DEFAULT_CHART_WIDTH = 480;
export const DEFAULT_CHART_HEIGHT = 300;

/**
 * The colours a series takes, in order.
 *
 * Eight, because a ninth series on one chart is a chart nobody can
 * read and cycling is a better answer than inventing a colour nobody
 * chose. Picked to stay apart in the common kinds of colour blindness
 * rather than to look like a palette.
 */
export const SERIES_COLOURS: readonly string[] = [
  '#3b6fd4',
  '#e2703a',
  '#3f9e5c',
  '#b4499a',
  '#c9a227',
  '#2f9fb0',
  '#9a5bd4',
  '#a0522d'
];

export function colourOf(index: number): string {
  return SERIES_COLOURS[index % SERIES_COLOURS.length];
}

/** Whether a kind has axes at all. */
export function hasAxes(kind: ChartKind): boolean {
  return kind !== 'pie';
}

/** Whether the value axis is the horizontal one. */
export function isSideways(kind: ChartKind): boolean {
  return kind === 'bar';
}

/** Whether the series are piled up rather than drawn beside each other. */
export function isStacked(kind: ChartKind): boolean {
  return kind === 'stacked';
}

/**
 * The span a chart's value axis has to cover.
 *
 * **Zero is included for anything drawn as an area or a bar**, and
 * left out for a line or a scatter. A bar chart whose axis starts at
 * ninety-eight makes a two-percent difference look like a tenfold
 * one, which is the most common way a chart lies; a line of
 * temperatures forced down to zero is a flat line that says nothing,
 * which is the second most common. The kind decides, because the kind
 * is what the reader is comparing: lengths in one case and positions
 * in the other.
 *
 * A stacked chart spans the totals, not the series, or the pile runs
 * off the top of its own axis.
 */
export function valueSpan(series: readonly Series[], kind: ChartKind): { low: number; high: number } {
  const values: number[] = [];
  if (isStacked(kind)) {
    for (const total of stackedTotals(series)) {
      values.push(total);
    }
  } else {
    for (const one of series) {
      for (const point of one.points) {
        values.push(point.y);
      }
    }
  }
  if (values.length === 0) {
    return { low: 0, high: 1 };
  }
  let low = Math.min(...values);
  let high = Math.max(...values);
  if (kind === 'area' || kind === 'column' || kind === 'bar' || kind === 'stacked') {
    low = Math.min(low, 0);
    high = Math.max(high, 0);
  }
  if (low === high) {
    // A flat series still needs a span, or every scale divides by
    // zero. One unit either side puts the line through the middle.
    return low === 0 ? { low: 0, high: 1 } : { low: low - Math.abs(low) / 2, high: high + Math.abs(high) / 2 };
  }
  return { low, high };
}

/** The pile height at each position, for a stacked chart. */
export function stackedTotals(series: readonly Series[]): readonly number[] {
  const totals: number[] = [];
  for (const one of series) {
    for (const point of one.points) {
      totals[point.x] = (totals[point.x] ?? 0) + point.y;
    }
  }
  // A position nothing reached is a zero rather than a hole, because
  // it is being summed rather than plotted.
  return Array.from(totals, total => total ?? 0);
}

export interface Ticks {
  /** The axis span, widened to land on round numbers. */
  readonly low: number;
  readonly high: number;
  readonly step: number;
  readonly values: readonly number[];
}

/**
 * Round numbers to label an axis with.
 *
 * The step is the smallest of 1, 2, 5 or 10 times a power of ten that
 * gives no more than `want` gaps, and the span is then widened out to
 * whole steps. The 1-2-5 ladder is the one every plotting library
 * converges on for the same reason: those are the numbers people
 * divide by in their heads, and an axis labelled 0, 3.7, 7.4 is an
 * axis nobody reads.
 *
 * `want` is a *want* and not a promise. Widening the span to whole
 * steps can add one gap at each end, so an axis asked for five ticks
 * can come back with seven, and pretending otherwise would mean
 * either uneven steps or labels that do not reach the data.
 */
export function niceTicks(low: number, high: number, want: number): Ticks {
  if (!Number.isFinite(low) || !Number.isFinite(high) || want < 1) {
    return { low: 0, high: 1, step: 1, values: [0, 1] };
  }
  if (high <= low) {
    return { low, high: low + 1, step: 1, values: [low, low + 1] };
  }
  const rough = (high - low) / want;
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  const step = [1, 2, 5, 10].map(one => one * magnitude).find(one => one >= rough) ?? 10 * magnitude;
  const first = Math.floor(low / step) * step;
  const last = Math.ceil(high / step) * step;
  const values: number[] = [];
  // Counted rather than accumulated: adding a step repeatedly gathers
  // float error and puts 0.30000000000000004 on an axis.
  const count = Math.round((last - first) / step);
  for (let at = 0; at <= count; at++) {
    values.push(round(first + at * step, step));
  }
  return { low: first, high: last, step, values };
}

/**
 * A tick, printed at the precision its own step deserves.
 *
 * A step of 0.25 wants two decimals and a step of 2,000 wants none,
 * and working it out from the step rather than from the value is what
 * keeps an axis from mixing `1` with `1.50`.
 */
export function tickLabel(value: number, step: number): string {
  if (!Number.isFinite(value)) {
    return '';
  }
  const places = decimalsIn(step);
  const text = value.toFixed(places);
  // `-0.00` is arithmetic showing through, and nobody means it.
  return text === `-${(0).toFixed(places)}` ? (0).toFixed(places) : text;
}

export interface Slice {
  readonly name: string;
  readonly value: number;
  /** Radians, clockwise from twelve o'clock. */
  readonly from: number;
  readonly to: number;
  readonly colour: string;
}

/**
 * One series as a ring of angles, starting at the top.
 *
 * **Negative values are left out rather than made positive.** A pie
 * is parts of a whole, and a negative part is not one; taking its
 * size would draw a slice that claims a share of a total it is
 * subtracting from. Dropping it is the only reading that does not
 * invent a number.
 */
export function pieSlices(series: Series | undefined, categories: readonly string[]): readonly Slice[] {
  if (series === undefined) {
    return [];
  }
  const parts = series.points.filter(point => point.y > 0);
  const total = parts.reduce((sum, point) => sum + point.y, 0);
  if (total <= 0) {
    return [];
  }
  const slices: Slice[] = [];
  let angle = -Math.PI / 2;
  parts.forEach((point, at) => {
    const sweep = (point.y / total) * Math.PI * 2;
    slices.push({
      name: categories[point.x] ?? `${point.x + 1}`,
      value: point.y,
      from: angle,
      to: angle + sweep,
      colour: colourOf(at)
    });
    angle += sweep;
  });
  return slices;
}

/** Rounded to the decimals its step implies, to keep float noise off an axis. */
function round(value: number, step: number): number {
  return Number(value.toFixed(Math.min(10, decimalsIn(step) + 1)));
}

/**
 * How many decimals a step actually has.
 *
 * From the step's own text rather than from `log10`, which gets 0.25
 * wrong: its logarithm says one decimal and writing it needs two, so
 * an axis stepping by a quarter came out labelled 1.0, 1.3, 1.5. Every
 * step this file produces is 1, 2, 5 or 10 times a power of ten, so
 * its text is short and exact and there is no float noise to strip.
 */
function decimalsIn(step: number): number {
  const size = Math.abs(step);
  if (!Number.isFinite(size) || size === 0) {
    return 0;
  }
  const text = String(size);
  const dot = text.indexOf('.');
  if (dot >= 0) {
    return Math.min(6, text.length - dot - 1);
  }
  // `1e-7` and smaller, where a step is past anything an axis labels.
  return text.includes('e-') ? 6 : 0;
}
