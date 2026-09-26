import { isError, type CellValue } from './Values';

/**
 * A range of cells, read as something a chart can draw.
 *
 * This is the file where Phase 15 makes its admission, and making it
 * here rather than in a comment on the renderer is the point.
 *
 * **A chart is the first thing that lets the render worker know about
 * cells that are not on screen.** Every phase before it has been able
 * to say that the render worker learns the window and nothing else. A
 * chart over `A1:B50000` is drawn from all fifty thousand of them, and
 * no amount of arranging changes that.
 *
 * What it does *not* learn is the cells. It learns a **series**: a
 * name and a list of numbers, built here from whatever the cells
 * happened to hold, with the errors, the blanks and the text already
 * resolved. And it learns the series at the resolution it can draw —
 * see `downsample`, which is the other half of the claim and the one
 * with a number attached to it.
 *
 * So the sentence the project has been writing since Phase 3 becomes:
 * *the render worker learns a series, at the resolution it can draw*.
 * That is weaker than the one it replaces, and it is true, which the
 * old one would not have been.
 *
 * Nothing here imports the framework or knows what a viewport is. It
 * is handed a rectangle of values and told how many points the thing
 * drawing them can show.
 */

/** One point. `x` is a category index for every kind but scatter. */
export interface SeriesPoint {
  readonly x: number;
  readonly y: number;
}

export interface Series {
  readonly name: string;
  readonly points: readonly SeriesPoint[];
}

export interface SeriesSet {
  /**
   * The labels down the first column, when there were any.
   *
   * Empty once the points have been thinned, and that is honest
   * rather than lazy: a downsampled line has no one cell behind each
   * point, so a label under it would name a row that is not what is
   * being drawn. A chart with fifty thousand points has no room for
   * category labels anyway.
   */
  readonly categories: readonly string[];
  readonly series: readonly Series[];
  /** How many points each series had before thinning. */
  readonly read: number;
}

export interface SeriesOptions {
  /**
   * Whether each series runs down a column or across a row.
   *
   * Columns is the common shape — a header row, then a row per
   * observation — and is what `orientationOf` guesses when nobody has
   * said.
   */
  readonly byColumn: boolean;
  /** The first row (or column) holds the series names. */
  readonly headers: boolean;
  /** The first column (or row) holds the category labels. */
  readonly labels: boolean;
  /**
   * The most points any series may carry, or zero for all of them.
   *
   * The chart's own width in pixels, which is the only number that
   * makes sense here: a point per pixel is the most a line can show,
   * and everything past that is bytes across a `postMessage` to draw
   * over something already drawn.
   */
  readonly limit: number;
}

/**
 * The values of a rectangle, as a chart.
 *
 * `grid` is row-major and rectangular; a short row is read as blanks,
 * because a rectangle read out of a sheet always is one and a caller
 * building one by hand should not have to pad it.
 */
export function seriesFrom(grid: readonly (readonly CellValue[])[], options: SeriesOptions): SeriesSet {
  const rows = grid.length;
  const columns = rows === 0 ? 0 : Math.max(...grid.map(row => row.length));
  if (rows === 0 || columns === 0) {
    return { categories: [], series: [], read: 0 };
  }

  // Everything below is written for series running down columns. A
  // row-wise range is transposed on the way in rather than doubling
  // every loop, which is the same trick and half the code.
  const cells = options.byColumn ? grid : transpose(grid, rows, columns);
  const height = cells.length;
  const width = Math.max(...cells.map(row => row.length));

  const firstRow = options.headers ? 1 : 0;
  const firstColumn = options.labels ? 1 : 0;
  if (height <= firstRow || width <= firstColumn) {
    return { categories: [], series: [], read: 0 };
  }

  const categories: string[] = [];
  if (options.labels) {
    for (let row = firstRow; row < height; row++) {
      categories.push(labelOf(cells[row]?.[0]));
    }
  }

  const series: Series[] = [];
  for (let column = firstColumn; column < width; column++) {
    const points: SeriesPoint[] = [];
    for (let row = firstRow; row < height; row++) {
      const value = numberOf(cells[row]?.[column]);
      if (value !== null) {
        points.push({ x: row - firstRow, y: value });
      }
    }
    series.push({
      name: options.headers ? labelOf(cells[0]?.[column]) || `Series ${column - firstColumn + 1}` : `Series ${column - firstColumn + 1}`,
      points
    });
  }

  const read = series.reduce((most, one) => Math.max(most, one.points.length), 0);
  if (options.limit <= 0 || read <= options.limit) {
    return { categories, series, read };
  }
  return {
    // Thinned, so the labels no longer describe what is drawn.
    categories: [],
    series: series.map(one => ({ name: one.name, points: downsample(one.points, options.limit) })),
    read
  };
}

/**
 * At most `limit` points, keeping the shape of the line.
 *
 * **Not every nth point**, which is the obvious answer and is wrong in
 * the way that matters: a spike one sample wide disappears, and a
 * spike is the thing somebody drew the chart to find. This buckets the
 * points and keeps the **smallest and the largest of each bucket**, in
 * the order they occur, so the envelope of the line survives and a
 * single-sample peak still reaches the top of the chart.
 *
 * Two points per bucket means `limit / 2` buckets, which is why the
 * arithmetic below halves before it divides. The guarantee this file
 * owes the rest of the phase is the simple one — *never more than
 * `limit`* — and it holds for every input including the degenerate
 * ones, which is what `Series.spec.ts` spends most of its length on.
 *
 * A bucket whose smallest and largest are the same point contributes
 * one rather than two, so a flat stretch does not pad the result out
 * with duplicates.
 */
export function downsample(points: readonly SeriesPoint[], limit: number): readonly SeriesPoint[] {
  if (limit <= 0 || points.length <= limit) {
    return points;
  }
  if (limit === 1) {
    return [points[0]];
  }
  const buckets = Math.floor(limit / 2);
  const kept: SeriesPoint[] = [];
  for (let bucket = 0; bucket < buckets; bucket++) {
    // Boundaries from the bucket index rather than by accumulating a
    // step, so the last bucket ends exactly at the end and no point
    // is dropped or counted twice by rounding.
    const from = Math.floor((bucket * points.length) / buckets);
    const to = Math.floor(((bucket + 1) * points.length) / buckets);
    if (to <= from) {
      continue;
    }
    let lowest = from;
    let highest = from;
    for (let at = from + 1; at < to; at++) {
      if (points[at].y < points[lowest].y) {
        lowest = at;
      }
      if (points[at].y > points[highest].y) {
        highest = at;
      }
    }
    if (lowest === highest) {
      kept.push(points[lowest]);
      continue;
    }
    const first = Math.min(lowest, highest);
    const second = Math.max(lowest, highest);
    kept.push(points[first], points[second]);
  }
  return kept;
}

/**
 * Which way the series run, guessed from the shape of the range.
 *
 * A range taller than it is wide holds a series per column, because
 * that is how a person types a table: a heading, then a row each. A
 * wide, short range holds a series per row. A square one is read by
 * column, because the common case has to win the tie.
 */
export function orientationOf(rows: number, columns: number): { byColumn: boolean } {
  return { byColumn: rows >= columns };
}

/**
 * How a chart reads a range: which way its series run, and whether its
 * first line names them and its first column labels them.
 *
 * One answer, shared by the chart drawn on screen and the chart written
 * into an `.xlsx`, whose series have to be spelled out cell by cell —
 * a file that guessed differently from the screen would chart other
 * numbers than the ones somebody looked at.
 */
export function layoutOf(grid: readonly (readonly CellValue[])[]): { byColumn: boolean; headers: boolean; labels: boolean } {
  const rows = grid.length;
  const columns = rows === 0 ? 0 : Math.max(...grid.map(row => row.length));
  const { byColumn } = orientationOf(rows, columns);
  const lines = byColumn ? grid : transpose(grid, rows, columns);
  const headers = headersIn(lines[0] ?? [], lines[1]);
  const labels = (lines[headers ? 1 : 0] ?? []).some(value => typeof value === 'string');
  return { byColumn, headers, labels };
}

/**
 * Whether the first line of a range looks like headings.
 *
 * Text over numbers, which is the same test `looksLikeHeader` makes
 * for a sort and is the only one that works without asking: a row of
 * words above rows of numbers is a heading in every spreadsheet
 * anybody has used.
 */
export function headersIn(line: readonly CellValue[], next: readonly CellValue[] | undefined): boolean {
  if (next === undefined) {
    return false;
  }
  const words = line.filter(value => typeof value === 'string' && value.trim() !== '').length;
  const numbers = next.filter(value => typeof value === 'number').length;
  return words > 0 && numbers > 0;
}

function transpose(
  grid: readonly (readonly CellValue[])[],
  rows: number,
  columns: number
): readonly (readonly CellValue[])[] {
  const out: CellValue[][] = [];
  for (let column = 0; column < columns; column++) {
    const line: CellValue[] = [];
    for (let row = 0; row < rows; row++) {
      line.push(grid[row]?.[column] ?? null);
    }
    out.push(line);
  }
  return out;
}

/**
 * A cell as a number, or null when it is not one.
 *
 * **A blank is not a zero and an error is not a zero**, and both would
 * be if this went through `toNumber`. A gap in a column of readings is
 * a gap in the line, and a `#DIV/0!` plotted at the axis is a reading
 * somebody would act on. Both are left out, and the point after keeps
 * its own index, so the line steps over the hole rather than closing
 * it up.
 */
function numberOf(value: CellValue | undefined): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return null;
  }
  return value;
}

function labelOf(value: CellValue | undefined): string {
  if (value === undefined || value === null || isError(value)) {
    return '';
  }
  return typeof value === 'string' ? value : String(value);
}
