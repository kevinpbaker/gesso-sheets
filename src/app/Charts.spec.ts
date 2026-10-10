import { describe, expect, it } from 'vitest';

import { relativeRef } from '../sheet/A1';
import type { SheetCharts, SheetSeriesView } from './SheetContract';
import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';

/**
 * Charts, through the service that publishes them — Phase 15.
 *
 * **The exit criterion lives here.** `Series.spec.ts` asserts that the
 * thinning never returns more points than it was asked for; this
 * asserts that the number it is asked for is the chart's own width,
 * end to end from a range of fifty thousand cells to what crosses the
 * barrier. A chart four hundred pixels wide publishes at most four
 * hundred points.
 *
 * The other half is the two keys. What a chart *is* and what a chart
 * *shows* change at different rates, and the reason they are separate
 * is that either one sharing the other's key would put the whole of
 * itself on the wire every time the other moved.
 */
function harness(rows = 0) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  for (let row = 0; row < rows; row++) {
    document.setCell(row, 0, String(Math.sin(row / 400) * 100));
  }
  document.sheet.recalculate();
  const service = new SheetService(document, { schedule, rowCount: 60_000, columnCount: 20 });
  const drain = (): void => {
    while (queue.length > 0) {
      queue.shift()?.();
    }
  };
  const charts = (): SheetCharts => {
    let held: SheetCharts = { entries: [], selected: 0 };
    service.charts.subscribe(value => (held = value)).unsubscribe();
    return held;
  };
  const series = (): SheetSeriesView => {
    let held: SheetSeriesView = { charts: {} };
    service.chartSeries.subscribe(value => (held = value)).unsubscribe();
    return held;
  };
  const only = () => series().charts[String(charts().entries[0].id)];
  return { document, service, drain, charts, series, only };
}

describe('putting a chart on a sheet', () => {
  it('reads the selection and lands beside it, not over it', () => {
    const h = harness(20);
    h.service.setSelection(0, 0, 5, 1);

    h.service.insertChart('line');

    const [chart] = h.charts().entries;
    expect(chart.range).toBe('A1:B6');
    expect(chart.kind).toBe('line');
    // Past the right-hand edge of column B, so the numbers stay
    // readable next to the picture of them.
    expect(chart.x).toBeGreaterThan(0);
  });

  it('selects what it just inserted', () => {
    const h = harness(20);
    h.service.insertChart('column');
    expect(h.charts().selected).toBe(h.charts().entries[0].id);
  });

  it('gives each chart an id nothing else has', () => {
    const h = harness(20);
    h.service.setSelection(0, 0, 3, 0);
    h.service.insertChart('line');
    h.service.insertChart('pie');

    const ids = h.charts().entries.map(entry => entry.id);
    expect(new Set(ids).size).toBe(2);
  });

  it('takes one back with undo, and puts it back with redo', () => {
    const h = harness(20);
    h.service.insertChart('line');
    expect(h.charts().entries).toHaveLength(1);

    h.service.undo();
    expect(h.charts().entries).toHaveLength(0);

    h.service.redo();
    expect(h.charts().entries).toHaveLength(1);
  });

  it('forgets the selected chart when another sheet is shown', () => {
    const h = harness(20);
    h.service.insertChart('line');
    h.service.addSheet();
    expect(h.charts().selected).toBe(0);
    expect(h.charts().entries).toHaveLength(0);
  });
});

describe('what a chart is sent to draw', () => {
  /** The phase's exit criterion, end to end. */
  it('sends at most as many points as the chart is wide', () => {
    const h = harness(50_000);
    h.service.setSelection(0, 0, 49_999, 0);
    h.service.insertChart('line');
    h.service.placeChart(h.charts().entries[0].id, 200, 40, 400, 300);

    const drawn = h.only();

    expect(drawn.series[0].points.length).toBeLessThanOrEqual(400);
    expect(drawn.read).toBe(50_000);
  });

  it('sends more points to a wider chart and fewer to a narrower one', () => {
    const h = harness(50_000);
    h.service.setSelection(0, 0, 49_999, 0);
    h.service.insertChart('line');
    const id = h.charts().entries[0].id;

    h.service.placeChart(id, 0, 0, 800, 300);
    const wide = h.only().series[0].points.length;
    h.service.placeChart(id, 0, 0, 200, 300);
    const narrow = h.only().series[0].points.length;

    expect(wide).toBeLessThanOrEqual(800);
    expect(narrow).toBeLessThanOrEqual(200);
    expect(narrow).toBeLessThan(wide);
  });

  /**
   * The thinning is not sampling, and the difference is visible from
   * out here: the peaks of the sine still reach the top.
   */
  it('keeps the shape of what it thinned', () => {
    const h = harness(50_000);
    h.service.setSelection(0, 0, 49_999, 0);
    h.service.insertChart('line');

    const highest = Math.max(...h.only().series[0].points.map(point => point.y));
    expect(highest).toBeGreaterThan(99);
  });

  it('redraws when a cell the chart reads is edited', () => {
    const h = harness(20);
    h.service.setSelection(0, 0, 19, 0);
    h.service.insertChart('line');
    const before = h.only().series[0].points[0].y;

    h.service.setCell(0, 0, '4321');
    h.drain();

    expect(h.only().series[0].points[0].y).not.toBe(before);
    expect(h.only().series[0].points[0].y).toBe(4321);
  });

  it('redraws when a formula the chart reads settles', () => {
    const h = harness(20);
    h.service.setCell(0, 1, '=A1*10');
    h.drain();
    h.service.setSelection(0, 1, 19, 1);
    h.service.insertChart('line');

    h.service.setCell(0, 0, '5');
    h.drain();

    expect(h.only().series[0].points[0].y).toBe(50);
  });

  /** So its value axis says $60K under a column of currency, and not 60000. */
  it('sends the number format of the cells it reads, and sends it again when it changes', () => {
    const h = harness(20);
    h.service.setSelection(0, 0, 19, 0);
    h.service.insertChart('line');
    expect(h.only().format).toEqual({ kind: 'general' });

    h.service.setSelection(0, 0, 19, 0);
    h.service.format({ number: { kind: 'currency', places: 2, symbol: '$' } });
    h.drain();

    expect(h.only().format).toEqual({ kind: 'currency', places: 2, symbol: '$' });
  });

  it('has nothing to send once the chart is gone', () => {
    const h = harness(20);
    h.service.insertChart('line');
    h.service.removeChart(h.charts().entries[0].id);

    expect(h.series().charts).toEqual({});
    expect(h.charts().selected).toBe(0);
  });
});

/**
 * The cost of the feature to a sheet that does not use it.
 *
 * A chart's range is read cell by cell on every publish, and the
 * guard that stops a sheet with no charts paying for that is the only
 * reason `pnpm proof` can still measure what it measures.
 */
describe('a sheet with no charts on it', () => {
  it('does not publish a series while two hundred formulas settle', () => {
    const h = harness(0);
    let published = 0;
    const watching = h.service.chartSeries.subscribe(() => published++);

    for (let row = 0; row < 200; row++) {
      h.service.setCell(row, 0, `=${row}+1`);
    }
    h.drain();
    watching.unsubscribe();

    // One, which is the value a `BehaviorSubject` hands over on
    // subscription, and not one more. The guard in `redrawCharts` is
    // what does it: without it every edit and every recalculation
    // slice would call `next` with a fresh empty object, and a
    // `BehaviorSubject` emits whatever it is given — the differ is
    // further down the wire than this.
    expect(published).toBe(1);
  });
});

/**
 * A chart is part of the document, so it is part of the file.
 *
 * The reading is deliberately suspicious: a file is outside the
 * program, and a chart with a kind this build does not know would be
 * drawn as something else if it were let through.
 */
describe('a chart that is written down', () => {
  it('comes back as the chart it was', () => {
    const before = new SheetDocument();
    before.setCell(0, 0, '5');
    before.addChart({
      kind: 'column',
      title: 'Takings',
      range: { start: relativeRef(0, 0), end: relativeRef(9, 1) },
      place: { x: 300, y: 40, width: 520, height: 260 },
      legend: false
    });

    const after = new SheetDocument();
    applySnapshot(after, snapshotOf(before));

    expect(after.charts).toHaveLength(1);
    expect(after.charts[0].kind).toBe('column');
    expect(after.charts[0].title).toBe('Takings');
    expect(after.charts[0].place).toEqual({ x: 300, y: 40, width: 520, height: 260 });
    expect(after.charts[0].legend).toBe(false);
    expect(after.charts[0].range.end.row).toBe(9);
  });

  /** Found in a browser: a chart of another sheet read its own sheet's cells after a reload. */
  it('keeps the sheet a chart reads through a save and a load', () => {
    const document = new SheetDocument();
    document.addSheet('Summary');
    document.activate(0);
    document.addChart({
      kind: 'pie',
      title: '',
      range: { start: { ...relativeRef(2, 0), sheet: 'Summary' }, end: { ...relativeRef(7, 3), sheet: 'Summary' } },
      place: { x: 0, y: 0, width: 300, height: 200 },
      legend: true
    });
    const back = new SheetDocument();
    applySnapshot(back, parseSnapshot(JSON.stringify(snapshotOf(document)), 20)!);
    back.activate(0);
    expect(back.charts[0].range.start.sheet).toBe('Summary');
    expect(back.charts[0].range.end.sheet).toBe('Summary');
  });

  it('reads a file written before charts existed as a sheet with none', () => {
    const read = parseSnapshot(
      JSON.stringify({ version: 2, cells: [{ row: 0, column: 0, input: '1' }] }),
      10
    );
    expect(read?.sheets[0].charts).toEqual([]);
  });

  it('drops a chart whose kind this build does not know', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        charts: [
          { kind: 'sunburst', range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, place: { x: 0, y: 0, width: 200, height: 200 } },
          { kind: 'pie', range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, place: { x: 0, y: 0, width: 200, height: 200 } }
        ]
      }),
      10
    );
    expect(read?.sheets[0].charts.map(chart => chart.kind)).toEqual(['pie']);
  });

  it('drops a chart with no range or no size rather than drawing nothing', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        charts: [
          { kind: 'pie', place: { x: 0, y: 0, width: 200, height: 200 } },
          { kind: 'pie', range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } } }
        ]
      }),
      10
    );
    expect(read?.sheets[0].charts).toEqual([]);
  });

  /**
   * Two charts sharing an id would be one chart as far as selecting
   * and dragging are concerned, and a file can repeat a number.
   */
  it('renumbers the ids it reads rather than trusting them', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        charts: [
          { id: 7, kind: 'pie', range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, place: { x: 0, y: 0, width: 200, height: 200 } },
          { id: 7, kind: 'line', range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, place: { x: 0, y: 0, width: 200, height: 200 } }
        ]
      }),
      10
    );
    const ids = read?.sheets[0].charts.map(chart => chart.id) ?? [];
    expect(new Set(ids).size).toBe(2);
  });
});
