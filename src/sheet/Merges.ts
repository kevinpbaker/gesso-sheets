import { shiftIndex, type Shift } from './Shift';

/**
 * Cells that span other cells.
 *
 * A merge is a rectangle drawn by its **anchor** — its top-left cell
 * — while the rest hold nothing and are not drawn at all. That is the
 * model every spreadsheet uses and the reason merging is destructive:
 * the cells it covers lose what was in them, because there is nowhere
 * left to show it.
 *
 * Kept as a list rather than a map per cell. Merges are few — tens on
 * a busy sheet — and every question asked of them is "what covers
 * this cell", which a short list answers by scanning. A map keyed per
 * covered cell would be the same information spread over a thousand
 * entries and would have to be rebuilt whenever a row moved.
 */

export interface MergeRect {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

export class Merges {
  private rects: MergeRect[] = [];

  get all(): readonly MergeRect[] {
    return this.rects;
  }

  get size(): number {
    return this.rects.length;
  }

  /** The merge covering a cell, or null. */
  at(row: number, column: number): MergeRect | null {
    for (const rect of this.rects) {
      if (covers(rect, row, column)) {
        return rect;
      }
    }
    return null;
  }

  /** Whether a cell is covered by a merge but is not its anchor. */
  isHidden(row: number, column: number): boolean {
    const rect = this.at(row, column);
    return rect !== null && !(rect.firstRow === row && rect.firstColumn === column);
  }

  /**
   * Adds a merge, dropping anything it overlaps.
   *
   * Overlapping merges have no meaning — two rectangles cannot both
   * own a cell — and the alternative to dropping them is refusing,
   * which leaves somebody looking at a selection that will not merge
   * and no explanation of which cell is the problem.
   *
   * A rectangle of one cell is not a merge and is ignored, so
   * "merge" on a single cell does nothing rather than creating an
   * entry that has to be special-cased everywhere after.
   */
  add(rect: MergeRect): void {
    if (rect.lastRow <= rect.firstRow && rect.lastColumn <= rect.firstColumn) {
      return;
    }
    this.rects = this.rects.filter(held => !overlaps(held, rect));
    this.rects.push(rect);
  }

  /** Removes every merge that intersects a rectangle. */
  remove(rect: MergeRect): boolean {
    const before = this.rects.length;
    this.rects = this.rects.filter(held => !overlaps(held, rect));
    return this.rects.length !== before;
  }

  /**
   * Moves the merges because the sheet changed shape.
   *
   * A merge grows when a row is inserted inside it and shrinks when
   * one is deleted — the same three behaviours a range in a formula
   * has, for the same reason: it names a run of cells and the run is
   * what changed. A merge left with one cell is no longer a merge and
   * goes.
   */
  shift(shift: Shift): void {
    const moved: MergeRect[] = [];
    for (const rect of this.rects) {
      const next = shiftRect(rect, shift);
      if (next !== null) {
        moved.push(next);
      }
    }
    this.rects = moved;
  }

  /** Replaces everything, for a load. */
  restore(rects: readonly MergeRect[]): void {
    this.rects = [];
    for (const rect of rects) {
      this.add(rect);
    }
  }
}

export function covers(rect: MergeRect, row: number, column: number): boolean {
  return row >= rect.firstRow && row <= rect.lastRow && column >= rect.firstColumn && column <= rect.lastColumn;
}

export function overlaps(a: MergeRect, b: MergeRect): boolean {
  return (
    a.firstRow <= b.lastRow && a.lastRow >= b.firstRow && a.firstColumn <= b.lastColumn && a.lastColumn >= b.firstColumn
  );
}

/** One merge under a shift, or null when there is nothing left of it. */
function shiftRect(rect: MergeRect, shift: Shift): MergeRect | null {
  const first = shift.axis === 'row' ? rect.firstRow : rect.firstColumn;
  const last = shift.axis === 'row' ? rect.lastRow : rect.lastColumn;

  let movedFirst: number;
  let movedLast: number;
  if (shift.by > 0) {
    movedFirst = first >= shift.at ? first + shift.by : first;
    movedLast = last >= shift.at ? last + shift.by : last;
  } else {
    const removed = -shift.by;
    const lastRemoved = shift.at + removed - 1;
    if (first >= shift.at && last <= lastRemoved) {
      return null;
    }
    movedFirst = first < shift.at ? first : Math.max(shift.at, first - removed);
    movedLast = last <= lastRemoved ? shift.at - 1 : last - removed;
  }

  const next =
    shift.axis === 'row'
      ? { ...rect, firstRow: movedFirst, lastRow: movedLast }
      : { ...rect, firstColumn: movedFirst, lastColumn: movedLast };
  // A merge shrunk to a single cell is not a merge any more.
  return next.lastRow > next.firstRow || next.lastColumn > next.firstColumn ? next : null;
}
