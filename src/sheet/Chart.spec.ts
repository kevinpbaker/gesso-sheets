import { describe, expect, it } from 'vitest';

import { colourOf, niceTicks, pieSlices, stackedTotals, tickLabel, valueSpan, type ChartKind } from './Chart';
import type { Series } from './Series';

/**
 * The arithmetic of drawing a chart, without drawing one.
 *
 * Every claim here is a number, which is the point of keeping the
 * geometry out of the renderer: an axis that labels itself 0, 3.7,
 * 7.4 is a bug somebody has to notice in a picture, and an axis whose
 * ticks are asserted is a bug that fails a run.
 */
describe('the span a value axis covers', () => {
  const series = (...values: number[]): Series => ({
    name: 'one',
    points: values.map((y, x) => ({ x, y }))
  });

  /**
   * The most common way a chart lies, and the second most common,
   * are the two halves of this decision.
   */
  it('includes zero for a chart read as lengths', () => {
    for (const kind of ['column', 'bar', 'stacked', 'area'] as ChartKind[]) {
      expect(valueSpan([series(98, 99, 100)], kind).low, kind).toBe(0);
    }
  });

  it('leaves zero out for a chart read as positions', () => {
    for (const kind of ['line', 'scatter'] as ChartKind[]) {
      expect(valueSpan([series(98, 99, 100)], kind).low, kind).toBe(98);
    }
  });

  it('reaches below zero when the data does', () => {
    expect(valueSpan([series(-40, 10)], 'column')).toEqual({ low: -40, high: 10 });
  });

  it('spans the totals of a stacked chart, not its largest series', () => {
    const stacked = [series(10, 10), { name: 'two', points: [{ x: 0, y: 30 }, { x: 1, y: 5 }] }];
    expect(valueSpan(stacked, 'stacked').high).toBe(40);
  });

  /** A flat line still has to divide by something. */
  it('gives a flat series a span rather than a zero one', () => {
    const span = valueSpan([series(7, 7, 7)], 'line');
    expect(span.high).toBeGreaterThan(span.low);
  });

  it('gives an empty chart a span rather than a NaN', () => {
    expect(valueSpan([], 'line')).toEqual({ low: 0, high: 1 });
  });
});

describe('piling series up', () => {
  it('adds each position across the series', () => {
    const totals = stackedTotals([
      { name: 'a', points: [{ x: 0, y: 1 }, { x: 1, y: 2 }] },
      { name: 'b', points: [{ x: 0, y: 10 }, { x: 1, y: 20 }] }
    ]);
    expect(totals).toEqual([11, 22]);
  });

  /** A gap in one series is a zero in the pile, not a hole in it. */
  it('treats a position one series skipped as nothing added', () => {
    const totals = stackedTotals([
      { name: 'a', points: [{ x: 0, y: 1 }, { x: 2, y: 3 }] },
      { name: 'b', points: [{ x: 0, y: 10 }] }
    ]);
    expect(totals).toEqual([11, 0, 3]);
  });
});

describe('choosing the numbers on an axis', () => {
  it('lands on a round step', () => {
    expect(niceTicks(0, 97, 5).step).toBe(20);
    expect(niceTicks(0, 1, 5).step).toBe(0.2);
    // Not 1,000: that would be five and a half gaps, which is six,
    // which is more than was asked for.
    expect(niceTicks(0, 5_500, 5).step).toBe(2_000);
  });

  it('widens the span out to whole steps', () => {
    const ticks = niceTicks(3, 97, 5);
    expect(ticks.low).toBe(0);
    expect(ticks.high).toBe(100);
    expect(ticks.values).toEqual([0, 20, 40, 60, 80, 100]);
  });

  it('reaches the data at both ends', () => {
    for (const [low, high] of [[3, 97], [-17, 4], [0.003, 0.017], [-500, -20]]) {
      const ticks = niceTicks(low, high, 5);
      expect(ticks.low, `${low}..${high}`).toBeLessThanOrEqual(low);
      expect(ticks.high, `${low}..${high}`).toBeGreaterThanOrEqual(high);
    }
  });

  /**
   * Counted rather than accumulated: adding a step ten times gathers
   * float error and puts 0.30000000000000004 on an axis.
   */
  it('does not print float noise', () => {
    for (const value of niceTicks(0, 1, 5).values) {
      expect(String(value).length).toBeLessThan(6);
    }
  });

  it('gives a span of nothing something to draw', () => {
    expect(niceTicks(5, 5, 5).values.length).toBeGreaterThan(1);
    expect(niceTicks(Number.NaN, 10, 5).values).toEqual([0, 1]);
  });

  /**
   * `want` is a want. Widening to whole steps can add a gap at each
   * end, and the alternative is uneven steps or labels that stop
   * short of the data.
   */
  it('may come back with more gaps than were asked for', () => {
    const ticks = niceTicks(1, 99, 5);
    expect(ticks.values.length).toBeGreaterThanOrEqual(5);
    expect(ticks.values.length).toBeLessThanOrEqual(9);
  });
});

describe('labelling a tick', () => {
  it('takes its precision from the step, not the value', () => {
    expect(tickLabel(1, 0.25)).toBe('1.00');
    expect(tickLabel(1, 20)).toBe('1');
    expect(tickLabel(2_000, 1_000)).toBe('2000');
  });

  it('does not print a negative zero', () => {
    expect(tickLabel(-0, 0.5)).toBe('0.0');
  });
});

describe('a pie', () => {
  const series: Series = {
    name: 'share',
    points: [
      { x: 0, y: 25 },
      { x: 1, y: 25 },
      { x: 2, y: 50 }
    ]
  };

  it('starts at the top and goes all the way round', () => {
    const slices = pieSlices(series, ['a', 'b', 'c']);
    expect(slices[0].from).toBeCloseTo(-Math.PI / 2);
    expect(slices[slices.length - 1].to).toBeCloseTo(-Math.PI / 2 + Math.PI * 2);
  });

  it('gives each slice its share of the turn', () => {
    const slices = pieSlices(series, ['a', 'b', 'c']);
    expect(slices[0].to - slices[0].from).toBeCloseTo(Math.PI / 2);
    expect(slices[2].to - slices[2].from).toBeCloseTo(Math.PI);
  });

  it('names its slices from the categories', () => {
    expect(pieSlices(series, ['North', 'South', 'East']).map(slice => slice.name)).toEqual([
      'North',
      'South',
      'East'
    ]);
  });

  /**
   * A pie is parts of a whole and a negative part is not one. Taking
   * its size would draw a slice claiming a share of a total it is
   * subtracting from.
   */
  it('leaves a negative value out rather than taking its size', () => {
    const mixed: Series = { name: 'x', points: [{ x: 0, y: 30 }, { x: 1, y: -10 }, { x: 2, y: 70 }] };
    const slices = pieSlices(mixed, ['a', 'b', 'c']);
    expect(slices.map(slice => slice.value)).toEqual([30, 70]);
    expect(slices[1].to - slices[1].from).toBeCloseTo(Math.PI * 2 * 0.7);
  });

  it('draws nothing at all when there is nothing positive', () => {
    expect(pieSlices({ name: 'x', points: [{ x: 0, y: -1 }] }, [])).toEqual([]);
    expect(pieSlices(undefined, [])).toEqual([]);
  });
});

describe('series colours', () => {
  it('cycles rather than inventing a colour nobody chose', () => {
    expect(colourOf(8)).toBe(colourOf(0));
    expect(colourOf(0)).not.toBe(colourOf(1));
  });
});
