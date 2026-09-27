import { describe, expect, it } from 'vitest';

import { relativeRef } from '../sheet/A1';
import { COLUMN_WIDTH, ROW_HEIGHT } from './dimensions';
import { SheetDocument } from './SheetDocument';

/**
 * A chart sits over a cell, as it does in Excel, and moves with it: a
 * column widened to its left, rows inserted above it, the columns under
 * it deleted. Placed in pixels, it stayed where it was and ended up over
 * the numbers it was placed beside.
 */
describe('a chart over the cells', () => {
  /** A chart whose corner is 10 pixels into D, and 6 into row 5. */
  function charted(): SheetDocument {
    const document = new SheetDocument();
    document.columnWidths = Array.from({ length: 20 }, () => COLUMN_WIDTH);
    document.addChart({
      kind: 'line',
      title: '',
      range: { start: relativeRef(0, 0), end: relativeRef(3, 1) },
      place: { x: 3 * COLUMN_WIDTH + 10, y: 4 * ROW_HEIGHT + 6, width: 300, height: 200 },
      legend: false
    });
    return document;
  }
  const place = (document: SheetDocument) => document.charts[0].place;

  it('moves right when a column to its left is widened', () => {
    const document = charted();
    const widths = [...document.columnWidths];
    widths[0] += 50;
    document.columnWidths = widths;
    expect(place(document).x).toBe(3 * COLUMN_WIDTH + 10 + 50);
  });

  it('keeps its place in the column it starts in when that one is widened', () => {
    const document = charted();
    const widths = [...document.columnWidths];
    widths[3] = COLUMN_WIDTH * 2;
    document.columnWidths = widths;
    // Ten pixels into a column of 104 is the same share of one of 208.
    expect(place(document).x).toBe(3 * COLUMN_WIDTH + 20);
  });

  it('is left alone by a column to its right', () => {
    const document = charted();
    const widths = [...document.columnWidths];
    widths[9] = 300;
    document.columnWidths = widths;
    expect(place(document).x).toBe(3 * COLUMN_WIDTH + 10);
  });

  it('moves down with rows inserted above it, and back on undo', () => {
    const document = charted();
    document.applyShift({ axis: 'row', at: 1, by: 2 });
    expect(place(document).y).toBe(6 * ROW_HEIGHT + 6);
    document.undo();
    expect(place(document).y).toBe(4 * ROW_HEIGHT + 6);
    document.redo();
    expect(place(document).y).toBe(6 * ROW_HEIGHT + 6);
  });

  it('moves left with columns deleted before it, and to where they were when its own goes', () => {
    const document = charted();
    document.applyShift({ axis: 'column', at: 0, by: -1 });
    expect(place(document).x).toBe(2 * COLUMN_WIDTH + 10);
    document.applyShift({ axis: 'column', at: 1, by: -2 });
    expect(place(document).x).toBe(COLUMN_WIDTH);
  });

  it('is not moved by rows inserted below it', () => {
    const document = charted();
    document.applyShift({ axis: 'row', at: 10, by: 3 });
    expect(place(document)).toMatchObject({ x: 3 * COLUMN_WIDTH + 10, y: 4 * ROW_HEIGHT + 6 });
  });
});
