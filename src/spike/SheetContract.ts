import { channel } from 'gesso-framework';

/**
 * The barrier for the Phase 0 spike.
 *
 * Throwaway, and narrower than `SheetContract.ts` will be in Phase 2:
 * there are no edits, no selection and no formulas here, only the one
 * question this phase exists to answer — can the render worker ask for
 * a window of cells every frame and have them arrive in time.
 *
 * Both wire shapes are declared, because which one to keep is the
 * other thing the spike is for. See `WireShape`.
 */

/**
 * How a window of cells is laid out for the differ.
 *
 * `diffProjection` walks a projection structurally and
 * `diffArray` trims a common prefix and suffix, so the shape of the
 * value decides what a scroll of one row costs on the wire.
 *
 *  - `rows` is the shape the roadmap assumed: row-major arrays of
 *    display strings, relative to the block's origin. A scroll shifts
 *    every element, so no prefix or suffix matches and the whole
 *    window is re-sent as one splice.
 *  - `keyed` indexes by absolute row and column. A scroll of one row
 *    leaves every other key structurally equal, so the patch is the
 *    one row that entered and the one that left.
 *  - `columnKeyed` is the same thing with the axes swapped, and it is
 *    here because `keyed` is not symmetric: it makes a vertical scroll
 *    two patches and a horizontal one two patches *per mounted row*.
 *    Which major axis a sheet should use is a real question and this
 *    is what answers it.
 */
export type WireShape = 'rows' | 'keyed' | 'columnKeyed';

/** One published window of already-formatted display strings. */
export interface SheetBlock {
  readonly shape: WireShape;
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  /** `rows` shape: row-major, relative to `firstRow` / `firstColumn`. */
  readonly rows: readonly (readonly string[])[];
  /** `keyed`: absolute row then column. `columnKeyed`: column then row. */
  readonly cells: Readonly<Record<string, Readonly<Record<string, string>>>>;
}

/**
 * What the application thread is doing, for the readout.
 *
 * A separate view key from `block` so that publishing it does not
 * appear in the patch count it reports — the two are diffed
 * independently, which is the same reason `NotesContract` splits
 * `rows` from `open`.
 */
export interface SheetStats {
  /** Blocks published since the start. */
  readonly publishes: number;
  /** Patches the differ produced for `block`, cumulative. */
  readonly patches: number;
  /** Serialized patch bytes for `block`, cumulative. */
  readonly bytes: number;
  /** Patches for the block published last. */
  readonly lastPatches: number;
  /** Serialized size of the last block's patches, as a stand-in for the wire. */
  readonly lastBytes: number;
  /** Cells in the block published last. */
  readonly lastCells: number;
  readonly fetchBandRows: number;
  readonly fetchBandColumns: number;
  readonly shape: WireShape;
  /** Milliseconds the application thread burns before each publish. */
  readonly busyMs: number;
}

export interface SheetCommands {
  /** The range the render worker has mounted, sent when it changes. */
  setViewport(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void;
  /** Cells published beyond the asked range, on each side of each axis. */
  setFetchBand(rows: number, columns: number): void;
  setShape(shape: WireShape): void;
  /** Simulates a recalc: burn this many milliseconds before publishing. */
  setBusy(ms: number): void;
}

export interface SheetView {
  readonly block: SheetBlock;
  readonly stats: SheetStats;
}

export const EMPTY_BLOCK: SheetBlock = {
  shape: 'keyed',
  firstRow: 0,
  lastRow: -1,
  firstColumn: 0,
  lastColumn: -1,
  rows: [],
  cells: {}
};

/**
 * Reads one cell out of a block, or null when the block does not cover
 * it.
 *
 * Null is the interesting answer: it is a cell the render worker has
 * mounted and the application worker has not sent yet, which is the
 * round-trip risk made countable.
 */
export function valueAt(block: SheetBlock, row: number, column: number): string | null {
  if (block.shape === 'keyed') {
    return block.cells[row]?.[column] ?? null;
  }
  if (block.shape === 'columnKeyed') {
    return block.cells[column]?.[row] ?? null;
  }
  if (row < block.firstRow || row > block.lastRow || column < block.firstColumn || column > block.lastColumn) {
    return null;
  }
  return block.rows[row - block.firstRow]?.[column - block.firstColumn] ?? null;
}

/** The sheet the spike scrolls: exactly a million cells. */
export const ROW_COUNT = 10_000;
export const COLUMN_COUNT = 100;

export const Sheet = channel<SheetView, SheetCommands>('sheet', {
  block: EMPTY_BLOCK,
  stats: {
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
  }
});
