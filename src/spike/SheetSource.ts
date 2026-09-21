import { BehaviorSubject, combineLatest, map, Observable } from 'rxjs';

import { diffProjection } from 'gesso-framework';
import {
  COLUMN_COUNT,
  EMPTY_BLOCK,
  ROW_COUNT,
  type SheetBlock,
  type SheetStats,
  type WireShape
} from './SheetContract';

/**
 * A million cells with no formulas behind them.
 *
 * Phase 0 is about the wire and the window, so the values are a pure
 * function of the coordinates: nothing is stored, nothing is
 * recalculated, and the cost of producing a window is exactly the cost
 * of formatting it. The `busy` knob puts a recalc-shaped delay back in
 * when we want to see what one does to the round trip.
 */
function cellText(row: number, column: number): string {
  if (column === 0) {
    return `R${row + 1}`;
  }
  if (column === 1) {
    return LABELS[(row * 7 + column) % LABELS.length];
  }
  // A cheap hash, so the text differs per cell and a scroll cannot be
  // made to look fast by every cell holding the same string.
  const hashed = (row * 2_654_435_761 + column * 40_503) >>> 0;
  return ((hashed % 1_000_000) / 100).toFixed(2);
}

const LABELS = ['North', 'South', 'East', 'West', 'Central', 'Coastal', 'Inland', 'Northeast'];

interface Viewport {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

/**
 * The application thread's half of the spike.
 *
 * Plain RxJS with one framework import, and that one is only for the
 * instrumentation: `diffProjection` is the same differ `provide` will
 * run over what this publishes, so the patch counts in the readout are
 * the real ones rather than an estimate. It is run a second time here,
 * which overstates this thread's cost and leaves the render thread's —
 * the one the 60fps claim is about — untouched.
 */
export class SheetSource {
  readonly block: Observable<SheetBlock>;
  readonly stats: Observable<SheetStats>;

  private readonly viewport = new BehaviorSubject<Viewport>({
    firstRow: 0,
    lastRow: -1,
    firstColumn: 0,
    lastColumn: -1
  });
  private readonly fetchBand = new BehaviorSubject<{ rows: number; columns: number }>({ rows: 0, columns: 0 });
  private readonly shape = new BehaviorSubject<WireShape>('keyed');
  private readonly busy = new BehaviorSubject<number>(0);
  private readonly statsSubject: BehaviorSubject<SheetStats>;

  private previous: SheetBlock = EMPTY_BLOCK;
  private publishes = 0;
  private patches = 0;
  private bytes = 0;

  constructor() {
    this.statsSubject = new BehaviorSubject<SheetStats>({
      publishes: 0,
      patches: 0,
      bytes: 0,
      lastPatches: 0,
      lastBytes: 0,
      lastCells: 0,
      fetchBandRows: 0,
      fetchBandColumns: 0,
      shape: 'keyed',
      busyMs: 0
    });
    this.block = combineLatest([this.viewport, this.fetchBand, this.shape, this.busy]).pipe(
      map(([viewport, band, shape, busy]) => this.build(viewport, band, shape, busy))
    );
    this.stats = this.statsSubject;
  }

  setViewport(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void {
    this.viewport.next({ firstRow, lastRow, firstColumn, lastColumn });
  }

  setFetchBand(rows: number, columns: number): void {
    this.fetchBand.next({ rows: Math.max(0, rows), columns: Math.max(0, columns) });
  }

  setShape(shape: WireShape): void {
    this.shape.next(shape);
  }

  setBusy(ms: number): void {
    this.busy.next(Math.max(0, ms));
  }

  private build(viewport: Viewport, band: { rows: number; columns: number }, shape: WireShape, busy: number): SheetBlock {
    if (busy > 0) {
      burn(busy);
    }
    if (viewport.lastRow < viewport.firstRow || viewport.lastColumn < viewport.firstColumn) {
      return this.record({ ...EMPTY_BLOCK, shape }, band, busy);
    }
    const firstRow = Math.max(0, viewport.firstRow - band.rows);
    const lastRow = Math.min(ROW_COUNT - 1, viewport.lastRow + band.rows);
    const firstColumn = Math.max(0, viewport.firstColumn - band.columns);
    const lastColumn = Math.min(COLUMN_COUNT - 1, viewport.lastColumn + band.columns);

    if (shape === 'rows') {
      const rows: string[][] = [];
      for (let row = firstRow; row <= lastRow; row++) {
        const line: string[] = [];
        for (let column = firstColumn; column <= lastColumn; column++) {
          line.push(cellText(row, column));
        }
        rows.push(line);
      }
      return this.record({ shape, firstRow, lastRow, firstColumn, lastColumn, rows, cells: {} }, band, busy);
    }

    const cells: Record<string, Record<string, string>> = {};
    if (shape === 'columnKeyed') {
      for (let column = firstColumn; column <= lastColumn; column++) {
        const line: Record<string, string> = {};
        for (let row = firstRow; row <= lastRow; row++) {
          line[row] = cellText(row, column);
        }
        cells[column] = line;
      }
    } else {
      for (let row = firstRow; row <= lastRow; row++) {
        const line: Record<string, string> = {};
        for (let column = firstColumn; column <= lastColumn; column++) {
          line[column] = cellText(row, column);
        }
        cells[row] = line;
      }
    }
    return this.record({ shape, firstRow, lastRow, firstColumn, lastColumn, rows: [], cells }, band, busy);
  }

  /**
   * Counts what this block will cost on the wire, then hands it on.
   *
   * The diff is the one `provide` is about to run; the byte count is
   * `JSON.stringify` of the patches, which is not what structured
   * clone sends but is proportional to it and is the only number
   * available on this side of the boundary.
   */
  private record(block: SheetBlock, band: { rows: number; columns: number }, busy: number): SheetBlock {
    const patches = diffProjection('block', this.previous, block);
    const bytes = patches.length === 0 ? 0 : JSON.stringify(patches).length;
    this.previous = block;
    this.publishes++;
    this.patches += patches.length;
    this.bytes += bytes;
    this.statsSubject.next({
      publishes: this.publishes,
      patches: this.patches,
      bytes: this.bytes,
      lastPatches: patches.length,
      lastBytes: bytes,
      lastCells:
        block.lastRow < block.firstRow
          ? 0
          : (block.lastRow - block.firstRow + 1) * (block.lastColumn - block.firstColumn + 1),
      fetchBandRows: band.rows,
      fetchBandColumns: band.columns,
      shape: block.shape,
      busyMs: busy
    });
    return block;
  }
}

/** Occupies the thread the way a recalc would, without being one. */
function burn(ms: number): void {
  const until = performance.now() + ms;
  while (performance.now() < until) {
    /* spin */
  }
}
