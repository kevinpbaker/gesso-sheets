/**
 * What the status bar says about the selection, and how it says it.
 *
 * Separate from `Aggregate.ts`, which computes these, because this
 * half crosses the barrier and that half does not. The render worker
 * draws the statistics and must not import a line of the engine to do
 * it; `aggregateOf` reads the store, `Values` and `Sheet`, and an
 * import of it from the status bar would pull all three into the
 * render worker's bundle to get at one formatting function.
 */

export interface SheetStats {
  /** Cells holding anything at all, text included. */
  readonly count: number;
  /** Of those, the ones that are numbers. Errors are not. */
  readonly numeric: number;
  readonly sum: number;
  /** Over the numeric cells. Zero when there are none. */
  readonly average: number;
  readonly min: number;
  readonly max: number;
}

export const NO_STATS: SheetStats = { count: 0, numeric: 0, sum: 0, average: 0, min: 0, max: 0 };

/**
 * The statistics as the status bar prints them.
 *
 * Nothing at all when the selection holds no numbers — a status bar
 * that said `Sum 0` over a column of names would be stating a fact
 * about the empty set that reads as a fact about the names.
 */
export function describeStats(stats: SheetStats): string {
  if (stats.numeric === 0) {
    return stats.count === 0 ? '' : `Count ${stats.count.toLocaleString('en-US')}`;
  }
  return [
    `Sum ${number(stats.sum)}`,
    `Average ${number(stats.average)}`,
    `Count ${stats.count.toLocaleString('en-US')}`
  ].join('  ·  ');
}

/**
 * A number short enough for a status bar.
 *
 * Long decimals are the common case once an average is involved, and
 * a status bar is not where somebody reads a full-precision answer —
 * that is what a cell is for.
 */
function number(value: number): string {
  if (!Number.isFinite(value)) {
    return String(value);
  }
  return Number.isInteger(value)
    ? value.toLocaleString('en-US')
    : value.toLocaleString('en-US', { maximumFractionDigits: 4 });
}
