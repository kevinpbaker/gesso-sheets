import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { SheetService } from './SheetService';

/**
 * Rules and charts through inserts and deletes, and charts of another
 * sheet — after Part Four.
 *
 * The cells, the formats, the merges and the notes moved when a row
 * was inserted; the conditional formats, the validations and the
 * charts did not, so a rule over B2:B9 was over B2:B9 still after a row
 * went in above it, painting the wrong cells. Found while making a chart
 * able to read another sheet, which has to follow that sheet's rows.
 */
const at = (row: number, column: number, lastRow: number, lastColumn: number, sheet?: string) => ({
  start: { row, column, rowAbsolute: false, columnAbsolute: false, ...(sheet === undefined ? {} : { sheet }) },
  end: { row: lastRow, column: lastColumn, rowAbsolute: false, columnAbsolute: false }
});

describe('a structural edit', () => {
  function ruled(): SheetDocument {
    const document = new SheetDocument();
    document.addConditional({ range: at(3, 1, 9, 1), test: { kind: 'greaterThan', value: 1 }, paint: { bold: true } });
    document.addValidation({ range: at(3, 2, 9, 2), rule: { kind: 'number', min: 0 } });
    document.addChart({ kind: 'line', title: '', legend: false, range: at(3, 0, 9, 1), place: { x: 0, y: 0, width: 300, height: 200 } });
    return document;
  }

  it('moves the rules and the charts down with a row inserted above them', () => {
    const document = ruled();
    document.applyShift({ axis: 'row', at: 1, by: 2 });
    expect(document.conditional[0].range).toMatchObject({ start: { row: 5 }, end: { row: 11 } });
    expect(document.validations[0].range).toMatchObject({ start: { row: 5 }, end: { row: 11 } });
    expect(document.charts[0].range).toMatchObject({ start: { row: 5 }, end: { row: 11 } });
    expect(document.validationAt(10, 2)).not.toBeNull();
  });

  it('grows a range a row is inserted inside, and shrinks one a row is deleted from', () => {
    const document = ruled();
    document.applyShift({ axis: 'row', at: 5, by: 1 });
    expect(document.conditional[0].range).toMatchObject({ start: { row: 3 }, end: { row: 10 } });
    document.applyShift({ axis: 'row', at: 4, by: -3 });
    expect(document.conditional[0].range).toMatchObject({ start: { row: 3 }, end: { row: 7 } });
  });

  it('takes away a rule whose every row was deleted, and keeps the chart', () => {
    const document = ruled();
    document.applyShift({ axis: 'row', at: 3, by: -7 });
    expect(document.conditional).toEqual([]);
    expect(document.validations).toEqual([]);
    expect(document.charts).toHaveLength(1);
  });

  it('moves a column’s rules sideways with a column deleted before them', () => {
    const document = ruled();
    document.applyShift({ axis: 'column', at: 0, by: -1 });
    expect(document.conditional[0].range).toMatchObject({ start: { column: 0 }, end: { column: 0 } });
    expect(document.validations[0].range).toMatchObject({ start: { column: 1 }, end: { column: 1 } });
  });

  it('puts every rule and chart back on one undo, and moves them again on redo', () => {
    const document = ruled();
    document.applyShift({ axis: 'row', at: 3, by: -7 });
    document.undo();
    expect(document.conditional[0].range).toMatchObject({ start: { row: 3 }, end: { row: 9 } });
    expect(document.validations).toHaveLength(1);
    document.redo();
    expect(document.conditional).toEqual([]);
  });
});

describe('a chart of another sheet', () => {
  function workbook(): SheetDocument {
    const document = new SheetDocument();
    document.renameSheet(0, 'Summary');
    document.addSheet('Data');
    document.activate(1);
    document.setCell(0, 0, 'Month');
    document.setCell(0, 1, 'Units');
    ['Jan', 'Feb', 'Mar'].forEach((month, row) => {
      document.setCell(row + 1, 0, month);
      document.setCell(row + 1, 1, String((row + 1) * 10));
    });
    document.activate(0);
    document.addChart({ kind: 'column', title: 'Units', legend: false, range: at(0, 0, 3, 1, 'Data'), place: { x: 0, y: 0, width: 300, height: 200 } });
    return document;
  }

  it('draws the other sheet’s numbers', () => {
    const document = workbook();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    let drawn: number[] = [];
    service.chartSeries.subscribe(view => {
      const chart = Object.values(view.charts)[0];
      drawn = chart?.series[0]?.points.map(point => point.y) ?? [];
    });
    // Series are drawn after an edit; one on the summary sheet will do.
    service.setCell(10, 5, 'note');
    expect(drawn).toEqual([10, 20, 30]);
  });

  it('follows the other sheet when rows go in above its data', () => {
    const document = workbook();
    document.activate(1);
    document.applyShift({ axis: 'row', at: 0, by: 2 });
    document.activate(0);
    expect(document.charts[0].range).toMatchObject({ start: { row: 2, sheet: 'Data' }, end: { row: 5 } });
  });

  it('follows the other sheet when it is renamed', () => {
    const document = workbook();
    document.renameSheet(1, 'Figures');
    expect(document.charts[0].range.start.sheet).toBe('Figures');
  });
});
