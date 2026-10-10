import type { PaintBox, PaintSurface } from 'gesso-core';

import {
  colourOf,
  hasAxes,
  isSideways,
  isStacked,
  niceTicks,
  pieSlices,
  stackedTotals,
  tickLabel,
  valueSpan,
  type ChartKind
} from '../sheet/Chart';
import type { Series } from '../sheet/Series';
import type { SheetChart, SheetChartBand, SheetChartSeries } from './SheetContract';

/**
 * A chart, drawn.
 *
 * One function and a `PaintSurface`, which is the whole of what the
 * engine's `Paint` element hands over: a 2D surface with a transform
 * stack, paths, and one call that puts a line of text at a baseline —
 * whose own documentation says it is there for the labels a chart
 * puts on its axes.
 *
 * **Nothing here decides anything.** Where the axis starts, what the
 * ticks are, how a pie divides and whether zero is on the scale are
 * all `Chart.ts`, where they are numbers a spec can assert. This file
 * is the part that cannot be asserted — it turns those numbers into
 * strokes — and keeping it that thin is what stops the interesting
 * decisions from hiding inside a paint callback nobody can test.
 *
 * It runs when the painter's `inputs` change and not once per frame,
 * which is the engine's arrangement and the reason a chart standing
 * still over a scrolling sheet costs one image draw.
 */

/** Room for the value labels down the side, and the categories under. */
const VALUE_GUTTER = 52;
const CATEGORY_GUTTER = 22;
const PADDING = 12;
const TITLE_HEIGHT = 22;
const LEGEND_HEIGHT = 22;
const LABEL_SIZE = 10;

interface Plot {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export function drawChart(
  surface: PaintSurface,
  box: PaintBox,
  chart: SheetChart,
  data: SheetChartSeries | null
): void {
  const hairline = 1 / box.scale;

  surface.save();
  surface.beginPath();
  surface.roundRect(0, 0, box.width, box.height, 4);
  surface.fillColor('surface');
  surface.fill();
  surface.strokeColor('border');
  surface.lineWidth(hairline);
  surface.stroke();
  surface.restore();

  const series = data?.series ?? [];
  const categories = data?.categories ?? [];
  const bands = data?.bands ?? [];
  const banded = bands.some(band => band !== null);

  let top = PADDING;
  if (chart.title !== '') {
    surface.fillColor('text');
    surface.text(chart.title, box.width / 2, top + 14, { fontSize: 13, fontWeight: 600, align: 'center' });
    top += TITLE_HEIGHT;
  }

  let bottom = box.height - PADDING;
  const showLegend = chart.legend && series.length > 1;
  if (showLegend) {
    drawLegend(surface, box, series, bottom - 6);
    bottom -= LEGEND_HEIGHT;
  }

  /**
   * The honest label on a thinned line.
   *
   * A chart of fifty thousand readings drawn through four hundred
   * points is not a lie, but a reader who is not told is entitled to
   * think every point is a row. One line in the corner is the whole
   * cost of saying so.
   */
  if (data !== null && data.read > longestOf(series)) {
    surface.fillColor('textMuted');
    surface.text(`${data.read.toLocaleString()} points`, box.width - PADDING, bottom, {
      fontSize: 9,
      align: 'right'
    });
    bottom -= 12;
  }

  // What the shading is, in the same corner and the same voice.
  if (banded) {
    surface.fillColor('textMuted');
    surface.text('Shaded: P10 to P90', box.width - PADDING, bottom, { fontSize: 9, align: 'right' });
    bottom -= 12;
  }

  if (series.length === 0) {
    surface.fillColor('textMuted');
    surface.text('Nothing to draw', box.width / 2, (top + bottom) / 2, { fontSize: 11, align: 'center' });
    return;
  }

  if (chart.kind === 'pie') {
    drawPie(surface, { x: PADDING, y: top, width: box.width - PADDING * 2, height: bottom - top }, series, categories);
    return;
  }

  const sideways = isSideways(chart.kind);
  const plot: Plot = {
    x: PADDING + (sideways ? VALUE_GUTTER : VALUE_GUTTER),
    y: top,
    width: box.width - PADDING * 2 - VALUE_GUTTER,
    height: bottom - top - CATEGORY_GUTTER
  };
  if (plot.width < 8 || plot.height < 8) {
    return;
  }

  // The band can reach past the series, and is drawn inside the axes.
  const span = valueSpan(banded ? [...series, ...bandSeries(bands)] : series, chart.kind);
  const ticks = niceTicks(span.low, span.high, sideways ? 4 : 5);
  drawAxes(surface, box, plot, ticks, categories, series, chart.kind);

  switch (chart.kind) {
    case 'line':
    case 'scatter':
      drawBands(surface, plot, series, bands, ticks);
      drawPoints(surface, plot, series, ticks, chart.kind);
      break;
    case 'area':
      drawArea(surface, plot, series, ticks, bands);
      break;
    case 'stacked':
      drawStacked(surface, plot, series, ticks);
      break;
    default:
      drawBars(surface, plot, series, ticks, sideways);
  }
}

/**
 * The axes, their ticks and the lines across the plot.
 *
 * Gridlines rather than tick marks, because a value is read by
 * following it across to a bar and a mark at the edge does not help
 * with that. Drawn first, so everything else is over them.
 */
function drawAxes(
  surface: PaintSurface,
  box: PaintBox,
  plot: Plot,
  ticks: ReturnType<typeof niceTicks>,
  categories: readonly string[],
  series: readonly Series[],
  kind: ChartKind
): void {
  const hairline = 1 / box.scale;
  const sideways = isSideways(kind);

  surface.save();
  surface.lineWidth(hairline);
  for (const value of ticks.values) {
    const at = alongValue(value, ticks, plot, sideways);
    surface.beginPath();
    if (sideways) {
      surface.moveTo(at, plot.y);
      surface.lineTo(at, plot.y + plot.height);
    } else {
      surface.moveTo(plot.x, at);
      surface.lineTo(plot.x + plot.width, at);
    }
    // The zero line is the one a reader measures from, so it is drawn
    // as an axis and not as a gridline.
    surface.strokeColor(value === 0 ? 'border' : 'controlBorder');
    surface.stroke();

    surface.fillColor('textMuted');
    if (sideways) {
      surface.text(tickLabel(value, ticks.step), at, plot.y + plot.height + 13, {
        fontSize: LABEL_SIZE,
        align: 'center'
      });
    } else {
      surface.text(tickLabel(value, ticks.step), plot.x - 6, at + 3, { fontSize: LABEL_SIZE, align: 'right' });
    }
  }
  surface.restore();

  drawCategories(surface, plot, categories, series, kind);
}

/**
 * The labels along the category axis, thinned to what fits.
 *
 * Every nth label rather than every one, because a hundred categories
 * in four hundred pixels is a smear. Which nth comes from the width,
 * so the same chart wider shows more of them.
 */
function drawCategories(
  surface: PaintSurface,
  plot: Plot,
  categories: readonly string[],
  series: readonly Series[],
  kind: ChartKind
): void {
  if (categories.length === 0) {
    return;
  }
  const sideways = isSideways(kind);
  const room = sideways ? plot.height : plot.width;
  const step = Math.max(1, Math.ceil(categories.length / Math.floor(room / 48)));
  const count = countOf(series);
  surface.fillColor('textMuted');
  for (let at = 0; at < categories.length; at += step) {
    const middle = alongCategory(at, count, plot, sideways) + bandOf(count, plot, sideways) / 2;
    if (sideways) {
      surface.text(categories[at], plot.x - 6, middle + 3, { fontSize: LABEL_SIZE, align: 'right' });
    } else {
      surface.text(categories[at], middle, plot.y + plot.height + 14, { fontSize: LABEL_SIZE, align: 'center' });
    }
  }
}

function drawPoints(surface: PaintSurface, plot: Plot, series: readonly Series[], ticks: ReturnType<typeof niceTicks>, kind: ChartKind): void {
  const count = countOf(series);
  series.forEach((one, index) => {
    const colour = colourOf(index);
    if (kind === 'scatter') {
      surface.fillColor(colour);
      for (const point of one.points) {
        surface.beginPath();
        surface.arc(alongCategory(point.x, count, plot, false) + bandOf(count, plot, false) / 2, alongValue(point.y, ticks, plot, false), 2.5, 0, Math.PI * 2);
        surface.fill();
      }
      return;
    }
    surface.save();
    surface.strokeColor(colour);
    surface.lineWidth(1.5);
    surface.lineJoin('round');
    surface.lineCap('round');
    surface.beginPath();
    one.points.forEach((point, at) => {
      const x = alongCategory(point.x, count, plot, false) + bandOf(count, plot, false) / 2;
      const y = alongValue(point.y, ticks, plot, false);
      if (at === 0) {
        surface.moveTo(x, y);
      } else {
        surface.lineTo(x, y);
      }
    });
    surface.stroke();
    surface.restore();
  });
}

function drawArea(
  surface: PaintSurface,
  plot: Plot,
  series: readonly Series[],
  ticks: ReturnType<typeof niceTicks>,
  bands: readonly (SheetChartBand | null)[]
): void {
  const count = countOf(series);
  const base = alongValue(Math.max(ticks.low, 0), ticks, plot, false);
  series.forEach((one, index) => {
    if (one.points.length === 0) {
      return;
    }
    surface.save();
    surface.fillColor(colourOf(index));
    // Stacked areas would hide each other outright; overlaid ones at
    // least show both, and the translucency says they are overlaid.
    surface.alpha(0.35);
    surface.beginPath();
    const middle = bandOf(count, plot, false) / 2;
    surface.moveTo(alongCategory(one.points[0].x, count, plot, false) + middle, base);
    for (const point of one.points) {
      surface.lineTo(alongCategory(point.x, count, plot, false) + middle, alongValue(point.y, ticks, plot, false));
    }
    surface.lineTo(alongCategory(one.points[one.points.length - 1].x, count, plot, false) + middle, base);
    surface.closePath();
    surface.fill();
    surface.restore();
  });
  // Over the fill and under the line, so the band reads as a halo round
  // the forecast rather than as more of the area.
  drawBands(surface, plot, series, bands, ticks);
  drawPoints(surface, plot, series, ticks, 'line');
}

/**
 * Each series' P10–P90 band: the shape between its two percentile
 * lines, in its own colour, faint, with the two edges drawn as hairlines
 * so the band has a top and a bottom where it is thin.
 */
function drawBands(
  surface: PaintSurface,
  plot: Plot,
  series: readonly Series[],
  bands: readonly (SheetChartBand | null)[],
  ticks: ReturnType<typeof niceTicks>
): void {
  const count = countOf(series);
  const middle = bandOf(count, plot, false) / 2;
  const x = (at: number): number => alongCategory(at, count, plot, false) + middle;
  const y = (value: number): number => alongValue(value, ticks, plot, false);
  bands.forEach((band, index) => {
    if (band === null || band.high.length === 0) {
      return;
    }
    surface.save();
    surface.fillColor(colourOf(index));
    surface.alpha(0.22);
    surface.beginPath();
    band.high.forEach((point, at) => (at === 0 ? surface.moveTo(x(point.x), y(point.y)) : surface.lineTo(x(point.x), y(point.y))));
    for (let at = band.low.length - 1; at >= 0; at--) {
      surface.lineTo(x(band.low[at].x), y(band.low[at].y));
    }
    surface.closePath();
    surface.fill();
    surface.alpha(0.55);
    surface.strokeColor(colourOf(index));
    surface.lineWidth(0.75);
    for (const edge of [band.high, band.low]) {
      surface.beginPath();
      edge.forEach((point, at) => (at === 0 ? surface.moveTo(x(point.x), y(point.y)) : surface.lineTo(x(point.x), y(point.y))));
      surface.stroke();
    }
    surface.restore();
  });
}

/** The bands' edges as series, so the value axis can be made to hold them. */
function bandSeries(bands: readonly (SheetChartBand | null)[]): Series[] {
  return bands.flatMap(band => (band === null ? [] : [{ name: '', points: band.low }, { name: '', points: band.high }]));
}

function drawBars(
  surface: PaintSurface,
  plot: Plot,
  series: readonly Series[],
  ticks: ReturnType<typeof niceTicks>,
  sideways: boolean
): void {
  const count = countOf(series);
  const band = bandOf(count, plot, sideways);
  // A gap of a fifth of the band, split between the two sides, and
  // the rest shared by however many series there are.
  const gap = band * 0.2;
  const each = (band - gap) / Math.max(1, series.length);
  const base = alongValue(Math.min(Math.max(0, ticks.low), ticks.high), ticks, plot, sideways);
  series.forEach((one, index) => {
    surface.fillColor(colourOf(index));
    for (const point of one.points) {
      const start = alongCategory(point.x, count, plot, sideways) + gap / 2 + index * each;
      const at = alongValue(point.y, ticks, plot, sideways);
      surface.beginPath();
      if (sideways) {
        surface.rect(Math.min(base, at), start, Math.abs(at - base), each);
      } else {
        surface.rect(start, Math.min(base, at), each, Math.abs(at - base));
      }
      surface.fill();
    }
  });
}

function drawStacked(surface: PaintSurface, plot: Plot, series: readonly Series[], ticks: ReturnType<typeof niceTicks>): void {
  const count = stackedTotals(series).length;
  const band = bandOf(count, plot, false);
  const gap = band * 0.2;
  const running: number[] = [];
  series.forEach((one, index) => {
    surface.fillColor(colourOf(index));
    for (const point of one.points) {
      const from = running[point.x] ?? 0;
      const to = from + point.y;
      running[point.x] = to;
      const top = alongValue(to, ticks, plot, false);
      const bottom = alongValue(from, ticks, plot, false);
      surface.beginPath();
      surface.rect(alongCategory(point.x, count, plot, false) + gap / 2, Math.min(top, bottom), band - gap, Math.abs(bottom - top));
      surface.fill();
    }
  });
}

function drawPie(surface: PaintSurface, plot: Plot, series: readonly Series[], categories: readonly string[]): void {
  const slices = pieSlices(series[0], categories);
  if (slices.length === 0) {
    return;
  }
  const middleX = plot.x + plot.width / 2;
  const middleY = plot.y + plot.height / 2;
  const radius = Math.max(4, Math.min(plot.width, plot.height) / 2 - 18);
  for (const slice of slices) {
    surface.fillColor(slice.colour);
    surface.beginPath();
    surface.moveTo(middleX, middleY);
    surface.arc(middleX, middleY, radius, slice.from, slice.to);
    surface.closePath();
    surface.fill();
  }
  // Labels outside the ring rather than inside it: a thin slice has
  // no room for its own name and a leader line is a different chart.
  surface.fillColor('text');
  for (const slice of slices) {
    if (slice.to - slice.from < 0.2) {
      continue;
    }
    const middle = (slice.from + slice.to) / 2;
    const x = middleX + Math.cos(middle) * (radius + 12);
    const y = middleY + Math.sin(middle) * (radius + 12);
    surface.text(slice.name, x, y + 3, {
      fontSize: LABEL_SIZE,
      align: Math.cos(middle) < -0.2 ? 'right' : Math.cos(middle) > 0.2 ? 'left' : 'center'
    });
  }
}

function drawLegend(surface: PaintSurface, box: PaintBox, series: readonly Series[], baseline: number): void {
  const names = series.map(one => one.name);
  // Laid out from the middle outwards, measured by an estimate rather
  // than by the text service: a painter has no measurer, and six
  // pixels a character is close enough for a row of short names.
  const widths = names.map(name => name.length * 5.5 + 20);
  const total = widths.reduce((sum, width) => sum + width, 0);
  let x = Math.max(PADDING, (box.width - total) / 2);
  names.forEach((name, index) => {
    surface.fillColor(colourOf(index));
    surface.beginPath();
    surface.roundRect(x, baseline - 7, 8, 8, 2);
    surface.fill();
    surface.fillColor('textMuted');
    surface.text(name, x + 12, baseline, { fontSize: LABEL_SIZE });
    x += widths[index];
  });
}

/** Where a value sits along the value axis, in the plot's pixels. */
function alongValue(value: number, ticks: ReturnType<typeof niceTicks>, plot: Plot, sideways: boolean): number {
  const span = ticks.high - ticks.low || 1;
  const share = (value - ticks.low) / span;
  return sideways ? plot.x + share * plot.width : plot.y + plot.height - share * plot.height;
}

/** Where a category's band starts, in the plot's pixels. */
function alongCategory(index: number, count: number, plot: Plot, sideways: boolean): number {
  const band = bandOf(count, plot, sideways);
  return (sideways ? plot.y : plot.x) + index * band;
}

function bandOf(count: number, plot: Plot, sideways: boolean): number {
  return (sideways ? plot.height : plot.width) / Math.max(1, count);
}

/** How many positions the category axis has to hold. */
function countOf(series: readonly Series[]): number {
  let most = 0;
  for (const one of series) {
    for (const point of one.points) {
      most = Math.max(most, point.x + 1);
    }
  }
  return Math.max(1, most);
}

function longestOf(series: readonly Series[]): number {
  return series.reduce((most, one) => Math.max(most, one.points.length), 0);
}
