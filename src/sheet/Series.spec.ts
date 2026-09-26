import { describe, expect, it } from 'vitest';

import { downsample, headersIn, orientationOf, seriesFrom, type SeriesPoint } from './Series';
import { DIV0, type CellValue } from './Values';

/**
 * A range, read as a chart — Phase 15's first half.
 *
 * The exit criterion this file owes the phase is one number: **a
 * series published for a chart `n` pixels wide holds at most `n`
 * points**, whatever it was built from. It is asserted here against
 * fifty thousand of them, and against every degenerate input that
 * could make the arithmetic fall off one end.
 *
 * The rest is about what a cell *means* to a chart, which is a
 * different question from what it means to a formula: a blank is a
 * gap rather than a zero, an error is a gap rather than an argument,
 * and a spike one sample wide has to survive being thinned or the
 * thinning is a lie.
 */
describe('reading a range as series', () => {
  const grid = (...rows: CellValue[][]): CellValue[][] => rows;

  it('takes a series per column, named by the heading over it', () => {
    const read = seriesFrom(
      grid(['Month', 'Sales', 'Costs'], ['Jan', 10, 4], ['Feb', 20, 6]),
      { byColumn: true, headers: true, labels: true, limit: 0 }
    );

    expect(read.series.map(one => one.name)).toEqual(['Sales', 'Costs']);
    expect(read.series[0].points).toEqual([
      { x: 0, y: 10 },
      { x: 1, y: 20 }
    ]);
    expect(read.categories).toEqual(['Jan', 'Feb']);
  });

  it('reads a wide range across its rows instead', () => {
    const read = seriesFrom(
      grid(['Month', 'Jan', 'Feb', 'Mar'], ['Sales', 10, 20, 30]),
      { byColumn: false, headers: true, labels: true, limit: 0 }
    );

    expect(read.series.map(one => one.name)).toEqual(['Sales']);
    expect(read.series[0].points.map(point => point.y)).toEqual([10, 20, 30]);
    expect(read.categories).toEqual(['Jan', 'Feb', 'Mar']);
  });

  it('names a series that has no heading after its place in the range', () => {
    const read = seriesFrom(grid([1, 5], [2, 6]), { byColumn: true, headers: false, labels: false, limit: 0 });
    expect(read.series.map(one => one.name)).toEqual(['Series 1', 'Series 2']);
  });

  /**
   * A gap is a gap. Plotting a blank at zero draws a cliff that is
   * not in the data, and plotting an error there draws a reading
   * somebody might act on.
   */
  it('steps over a blank rather than plotting it at zero', () => {
    const read = seriesFrom(grid([10], [null], [30]), { byColumn: true, headers: false, labels: false, limit: 0 });

    expect(read.series[0].points).toEqual([
      { x: 0, y: 10 },
      { x: 2, y: 30 }
    ]);
  });

  it('steps over an error the same way', () => {
    const read = seriesFrom(grid([10], [DIV0], [30]), { byColumn: true, headers: false, labels: false, limit: 0 });
    expect(read.series[0].points.map(point => point.y)).toEqual([10, 30]);
  });

  it('leaves text out of a series rather than reading it as a number', () => {
    const read = seriesFrom(grid([10], ['n/a'], [30]), { byColumn: true, headers: false, labels: false, limit: 0 });
    expect(read.series[0].points.map(point => point.y)).toEqual([10, 30]);
  });

  it('has nothing to say about an empty range', () => {
    expect(seriesFrom([], { byColumn: true, headers: false, labels: false, limit: 0 }).series).toEqual([]);
    expect(seriesFrom([[]], { byColumn: true, headers: false, labels: false, limit: 0 }).series).toEqual([]);
  });

  it('has nothing to say about a range that is only its own headings', () => {
    const read = seriesFrom(grid(['Month', 'Sales']), { byColumn: true, headers: true, labels: true, limit: 0 });
    expect(read.series).toEqual([]);
  });
});

/**
 * The exit criterion, and the shape it has to keep while it meets it.
 */
describe('thinning a series to what can be drawn', () => {
  const ramp = (count: number): SeriesPoint[] =>
    Array.from({ length: count }, (_, at) => ({ x: at, y: at }));

  it('holds at most as many points as the chart is wide', () => {
    for (const width of [1, 2, 3, 7, 100, 399, 400, 401]) {
      expect(downsample(ramp(50_000), width).length, `at ${width} wide`).toBeLessThanOrEqual(width);
    }
  });

  /** The phase's own number, written out rather than derived. */
  it('holds at most four hundred points for a chart four hundred pixels wide', () => {
    expect(downsample(ramp(50_000), 400).length).toBeLessThanOrEqual(400);
  });

  it('leaves a series that already fits exactly as it was', () => {
    const points = ramp(40);
    expect(downsample(points, 400)).toBe(points);
    expect(downsample(points, 40)).toBe(points);
  });

  it('keeps every point when no limit is asked for', () => {
    const points = ramp(50_000);
    expect(downsample(points, 0)).toBe(points);
  });

  /**
   * The reason it is not every nth point.
   *
   * One sample out of fifty thousand, three times the height of
   * everything around it. Taking every 125th sample loses it unless
   * it happens to land on one, which is the same as losing it.
   */
  it('keeps a spike one sample wide', () => {
    const points = ramp(50_000).map(point => ({ ...point, y: 10 }));
    const spiked = [...points];
    spiked[31_337] = { x: 31_337, y: 9_999 };

    const thinned = downsample(spiked, 400);

    expect(thinned.some(point => point.y === 9_999)).toBe(true);
  });

  it('keeps a trough the same way', () => {
    const points = ramp(50_000).map(point => ({ ...point, y: 10 }));
    points[12_345] = { x: 12_345, y: -9_999 };

    expect(downsample(points, 400).some(point => point.y === -9_999)).toBe(true);
  });

  it('keeps the points in the order they were read', () => {
    const thinned = downsample(ramp(50_000), 400);
    for (let at = 1; at < thinned.length; at++) {
      expect(thinned[at].x).toBeGreaterThan(thinned[at - 1].x);
    }
  });

  /**
   * A flat line thinned to a bucket's worth of duplicates would be
   * twice the points for none of the shape.
   */
  it('contributes one point for a bucket that is flat', () => {
    const flat = Array.from({ length: 10_000 }, (_, at) => ({ x: at, y: 5 }));
    expect(downsample(flat, 400).length).toBe(200);
  });

  it('survives being asked for one point', () => {
    expect(downsample(ramp(50_000), 1)).toEqual([{ x: 0, y: 0 }]);
  });

  it('survives a series of one point', () => {
    expect(downsample([{ x: 0, y: 3 }], 400)).toEqual([{ x: 0, y: 3 }]);
  });
});

describe('what the range is read as when nobody has said', () => {
  it('reads a tall range down its columns and a wide one across its rows', () => {
    expect(orientationOf(50, 3).byColumn).toBe(true);
    expect(orientationOf(3, 50).byColumn).toBe(false);
  });

  /** The common case wins the tie: a square range is a table of columns. */
  it('reads a square range down its columns', () => {
    expect(orientationOf(4, 4).byColumn).toBe(true);
  });

  it('sees words over numbers as headings', () => {
    expect(headersIn(['Month', 'Sales'], ['Jan', 10])).toBe(true);
  });

  it('does not see numbers over numbers as headings', () => {
    expect(headersIn([1, 2], [3, 4])).toBe(false);
  });

  it('does not see headings in a range with only one line', () => {
    expect(headersIn(['Month', 'Sales'], undefined)).toBe(false);
  });
});

/**
 * The two halves together, which is how the service will call it: a
 * range far too big to draw, asked for at a chart's width.
 */
describe('a range too big to draw', () => {
  it('comes back at the width it was asked for, and says what it read', () => {
    const grid: CellValue[][] = [['Reading']];
    for (let row = 0; row < 50_000; row++) {
      grid.push([Math.sin(row / 400) * 100]);
    }

    const read = seriesFrom(grid, { byColumn: true, headers: true, labels: false, limit: 400 });

    expect(read.series[0].points.length).toBeLessThanOrEqual(400);
    expect(read.read).toBe(50_000);
    expect(read.series[0].name).toBe('Reading');
    // The peaks of the sine are still there, which is the whole
    // reason the thinning keeps extremes rather than samples.
    const highest = Math.max(...read.series[0].points.map(point => point.y));
    expect(highest).toBeGreaterThan(99);
  });

  it('drops the category labels, because they no longer describe a point', () => {
    const grid: CellValue[][] = [];
    for (let row = 0; row < 5_000; row++) {
      grid.push([`row ${row}`, row]);
    }

    const read = seriesFrom(grid, { byColumn: true, headers: false, labels: true, limit: 400 });

    expect(read.series[0].points.length).toBeLessThanOrEqual(400);
    expect(read.categories).toEqual([]);
  });
});
