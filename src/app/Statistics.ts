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

/** The figures the status bar can show, in the order it shows them. */
export type StatFigure = 'sum' | 'average' | 'count' | 'numeric' | 'min' | 'max';

export const FIGURES: readonly { readonly id: StatFigure; readonly label: string }[] = [
  { id: 'average', label: 'Average' },
  { id: 'count', label: 'Count' },
  { id: 'numeric', label: 'Numerical count' },
  { id: 'min', label: 'Min' },
  { id: 'max', label: 'Max' },
  { id: 'sum', label: 'Sum' }
];

/** Whether a stored value is one of the figures, for reading a preference back. */
export function isFigure(value: unknown): value is StatFigure {
  return typeof value === 'string' && FIGURES.some(figure => figure.id === value);
}

/** What the status bar shows until somebody chooses otherwise: Excel's three. */
export const DEFAULT_FIGURES: readonly StatFigure[] = ['sum', 'average', 'count'];

/**
 * The chosen figures for a selection, each as the status bar prints it
 * and as a click copies it.
 *
 * `describeStats`'s rule holds: a selection with no numbers shows its
 * count and nothing that would be a fact about the empty set, and one
 * with nothing in it shows nothing at all. Sum first, as in the default
 * readout, then the rest in `FIGURES` order.
 */
export function figuresOf(
  stats: SheetStats,
  chosen: readonly StatFigure[]
): readonly { readonly id: StatFigure; readonly label: string; readonly value: string; readonly copy: string }[] {
  if (stats.count === 0) {
    return [];
  }
  const order: StatFigure[] = ['sum', 'average', 'count', 'numeric', 'min', 'max'];
  const wanted = order.filter(id => chosen.includes(id) && (stats.numeric > 0 || id === 'count'));
  if (wanted.length === 0 && stats.numeric === 0) {
    wanted.push('count');
  }
  return wanted.map(id => {
    const label = FIGURES.find(figure => figure.id === id)?.label ?? id;
    const raw = id === 'numeric' ? stats.numeric : stats[id];
    const value = id === 'count' || id === 'numeric' ? raw.toLocaleString('en-US') : number(raw);
    // Copied whole: the status bar rounds for reading, and a figure
    // pasted into a cell should be the number and not its summary.
    return { id, label, value, copy: String(raw) };
  });
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
