import { describe, expect, it } from 'vitest';

import { Merges, type MergeRect } from './Merges';

const rect = (firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): MergeRect => ({
  firstRow,
  lastRow,
  firstColumn,
  lastColumn
});

describe('merged cells', () => {
  it('reports what covers a cell', () => {
    const merges = new Merges();
    merges.add(rect(1, 2, 1, 3));

    expect(merges.at(1, 1)).toEqual(rect(1, 2, 1, 3));
    expect(merges.at(2, 3)).toEqual(rect(1, 2, 1, 3));
    expect(merges.at(0, 1)).toBeNull();
    expect(merges.at(1, 4)).toBeNull();
  });

  /** Everything but the anchor is covered, and the anchor draws. */
  it('knows which cells are hidden and which is the anchor', () => {
    const merges = new Merges();
    merges.add(rect(1, 2, 1, 3));

    expect(merges.isHidden(1, 1)).toBe(false);
    expect(merges.isHidden(1, 2)).toBe(true);
    expect(merges.isHidden(2, 1)).toBe(true);
    expect(merges.isHidden(5, 5)).toBe(false);
  });

  /**
   * Two rectangles cannot both own a cell. Refusing instead would
   * leave somebody with a selection that will not merge and no
   * explanation of which cell is the problem.
   */
  it('drops what a new merge overlaps', () => {
    const merges = new Merges();
    merges.add(rect(0, 1, 0, 1));
    merges.add(rect(1, 2, 1, 2));

    expect(merges.size).toBe(1);
    expect(merges.at(0, 0)).toBeNull();
    expect(merges.at(2, 2)).not.toBeNull();
  });

  it('keeps merges that do not touch', () => {
    const merges = new Merges();
    merges.add(rect(0, 1, 0, 1));
    merges.add(rect(5, 6, 5, 6));
    expect(merges.size).toBe(2);
  });

  /** A rectangle of one cell is not a merge. */
  it('ignores a merge of one cell', () => {
    const merges = new Merges();
    merges.add(rect(3, 3, 3, 3));
    expect(merges.size).toBe(0);
  });

  it('removes every merge a rectangle touches', () => {
    const merges = new Merges();
    merges.add(rect(0, 1, 0, 1));
    merges.add(rect(5, 6, 5, 6));

    expect(merges.remove(rect(1, 1, 1, 1))).toBe(true);
    expect(merges.size).toBe(1);
    expect(merges.remove(rect(9, 9, 9, 9))).toBe(false);
  });
});

describe('merges under a shift', () => {
  const merges = () => {
    const held = new Merges();
    held.add(rect(2, 3, 1, 2));
    return held;
  };

  it('moves down when a row is inserted above it', () => {
    const held = merges();
    held.shift({ axis: 'row', at: 0, by: 1 });
    expect(held.all[0]).toEqual(rect(3, 4, 1, 2));
  });

  /** It names a run of cells, and the run is what changed. */
  it('grows when a row is inserted inside it', () => {
    const held = merges();
    held.shift({ axis: 'row', at: 3, by: 1 });
    expect(held.all[0]).toEqual(rect(2, 4, 1, 2));
  });

  it('is untouched by a row inserted below it', () => {
    const held = merges();
    held.shift({ axis: 'row', at: 50, by: 1 });
    expect(held.all[0]).toEqual(rect(2, 3, 1, 2));
  });

  /** Shrunk to a single cell, it is not a merge any more. */
  it('goes when a delete leaves it one cell wide', () => {
    const held = new Merges();
    held.add(rect(2, 2, 1, 2));
    held.shift({ axis: 'column', at: 2, by: -1 });
    expect(held.size).toBe(0);
  });

  it('goes when everything it covered was deleted', () => {
    const held = merges();
    held.shift({ axis: 'row', at: 2, by: -2 });
    expect(held.size).toBe(0);
  });

  it('shrinks when part of it was deleted', () => {
    const held = new Merges();
    held.add(rect(2, 5, 1, 2));
    held.shift({ axis: 'row', at: 4, by: -2 });
    expect(held.all[0]).toEqual(rect(2, 3, 1, 2));
  });

  it('moves sideways on a column shift', () => {
    const held = merges();
    held.shift({ axis: 'column', at: 0, by: 2 });
    expect(held.all[0]).toEqual(rect(2, 3, 3, 4));
  });
});

describe('reading merges back', () => {
  it('round-trips', () => {
    const held = new Merges();
    held.add(rect(0, 1, 0, 2));
    const loaded = new Merges();
    loaded.restore(held.all);
    expect(loaded.all).toEqual(held.all);
  });

  it('drops an overlap a file happens to contain', () => {
    const loaded = new Merges();
    loaded.restore([rect(0, 1, 0, 1), rect(1, 2, 1, 2)]);
    expect(loaded.size).toBe(1);
  });

  it('forgets what was there before', () => {
    const held = new Merges();
    held.add(rect(0, 1, 0, 1));
    held.restore([]);
    expect(held.size).toBe(0);
  });
});
