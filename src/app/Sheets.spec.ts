import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { SheetService } from './SheetService';
import type { SheetTabs } from './SheetContract';
import type { CellFormat } from '../sheet/Format';

/**
 * Many sheets, from the document and the service.
 *
 * The engine's half is `Workbook.spec.ts`. This is the half above it:
 * that everything drawn *over* the cells — formats, merges, widths,
 * freezes, hidden rows and where the selection sits — is per sheet,
 * that the tab strip says so, and that undo does not put text back
 * on the wrong tab.
 */

const bold = (format: CellFormat): CellFormat => ({ ...format, paint: { ...format.paint, bold: true } });
const italic = (format: CellFormat): CellFormat => ({ ...format, paint: { ...format.paint, italic: true } });

function harness() {
  const document = new SheetDocument();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  return { document, service };
}

const tabs = (service: SheetService): SheetTabs => {
  let seen: SheetTabs | undefined;
  service.sheets.subscribe(value => (seen = value)).unsubscribe();
  return seen!;
};

describe('the tab strip', () => {
  it('starts with one sheet', () => {
    const { service } = harness();
    expect(tabs(service)).toEqual({ entries: [{ name: 'Sheet1', colour: null }], active: 0 });
  });

  it('adds one and shows it', () => {
    const { service } = harness();
    service.addSheet();
    expect(tabs(service).entries.map(entry => entry.name)).toEqual(['Sheet1', 'Sheet2']);
    expect(tabs(service).active).toBe(1);
  });

  it('renames one', () => {
    const { service } = harness();
    service.addSheet();
    service.renameSheet(1, 'Figures');
    expect(tabs(service).entries[1].name).toBe('Figures');
  });

  it('colours one', () => {
    const { service } = harness();
    service.setSheetColour(0, '#ea4335');
    expect(tabs(service).entries[0].colour).toBe('#ea4335');
  });

  it('moves one along the strip', () => {
    const { service } = harness();
    service.addSheet();
    service.addSheet();
    service.moveSheet(0, 2);
    expect(tabs(service).entries.map(entry => entry.name)).toEqual(['Sheet2', 'Sheet3', 'Sheet1']);
  });

  it('removes one', () => {
    const { service } = harness();
    service.addSheet();
    service.removeSheet(0);
    expect(tabs(service).entries.map(entry => entry.name)).toEqual(['Sheet2']);
    expect(tabs(service).active).toBe(0);
  });

  /** Nothing in the application is written to survive a workbook of none. */
  it('will not remove the last one', () => {
    const { service } = harness();
    service.removeSheet(0);
    expect(tabs(service).entries).toHaveLength(1);
  });
});

describe('what belongs to a sheet rather than to the workbook', () => {
  it('keeps a selection per sheet', () => {
    const { document } = harness();
    document.setSelection(3, 4, 3, 4);
    document.addSheet();
    expect(document.selection).toMatchObject({ row: 0, column: 0 });

    document.activate(0);
    expect(document.selection).toMatchObject({ row: 3, column: 4 });
  });

  it('keeps column widths per sheet', () => {
    const { document } = harness();
    document.columnWidths = [10, 20, 30];
    document.addSheet();
    document.columnWidths = [99];

    document.activate(0);
    expect(document.columnWidths).toEqual([10, 20, 30]);
  });

  it('keeps merges per sheet', () => {
    const { document } = harness();
    document.merges.add({ firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 });
    document.addSheet();
    expect(document.merges.size).toBe(0);

    document.activate(0);
    expect(document.merges.size).toBe(1);
  });

  it('keeps formats per sheet', () => {
    const { document } = harness();
    document.setFormat(0, 0, bold(document.formatAt(0, 0)));
    document.addSheet();
    expect(document.formatAt(0, 0).paint.bold).toBe(false);

    document.activate(0);
    expect(document.formatAt(0, 0).paint.bold).toBe(true);
  });

  it('keeps freezes and hidden rows per sheet', () => {
    const { document } = harness();
    document.frozenRows = 2;
    document.hiddenRows.add(4);
    document.addSheet();
    expect(document.frozenRows).toBe(0);
    expect(document.hiddenRows.has(4)).toBe(false);

    document.activate(0);
    expect(document.frozenRows).toBe(2);
    expect(document.hiddenRows.has(4)).toBe(true);
  });

  /** A name means the same cells wherever it is written. */
  it('keeps names for the whole workbook', () => {
    const { document } = harness();
    document.addSheet();
    expect(document.book.names).toBe(document.sheet.names);
    document.activate(0);
    expect(document.book.names).toBe(document.sheet.names);
  });
});

describe('duplicating a sheet', () => {
  it('brings everything drawn over the cells with it', () => {
    const { document } = harness();
    document.setCell(0, 0, '7');
    document.setFormat(0, 0, bold(document.formatAt(0, 0)));
    document.merges.add({ firstRow: 2, lastRow: 3, firstColumn: 0, lastColumn: 1 });
    document.columnWidths = [55, 66];
    document.frozenRows = 1;

    const copy = document.duplicateSheet(0);
    expect(copy).toBe(1);
    expect(document.sheet.input(0, 0)).toBe('7');
    expect(document.formatAt(0, 0).paint.bold).toBe(true);
    expect(document.merges.size).toBe(1);
    expect(document.columnWidths).toEqual([55, 66]);
    expect(document.frozenRows).toBe(1);
  });

  /** The copy's formats are its own, not a second view of the original's. */
  it('does not let the copy format the original', () => {
    const { document } = harness();
    document.setCell(0, 0, '7');
    const copy = document.duplicateSheet(0);
    document.setFormat(0, 0, italic(document.formatAt(0, 0)));

    document.activate(0);
    expect(document.formatAt(0, 0).paint.italic).toBe(false);
    document.activate(copy);
    expect(document.formatAt(0, 0).paint.italic).toBe(true);
  });
});

describe('undo across sheets', () => {
  it('takes a step back on the sheet it was made on', () => {
    const { document } = harness();
    document.setCell(0, 0, 'on one');
    document.addSheet();
    document.setCell(0, 0, 'on two');

    document.undo();
    expect(document.active).toBe(1);
    expect(document.sheet.input(0, 0)).toBe('');
    document.activate(0);
    expect(document.sheet.input(0, 0)).toBe('on one');
  });

  /**
   * A sheet removed or moved renumbers the indices the stack holds,
   * and a rename rewrites the text it holds. Dropping the history is
   * what Excel does with a sheet delete, and for the same reason:
   * saying it is gone beats a ctrl-Z that writes on the wrong tab.
   */
  it('forgets the history when the sheets themselves change', () => {
    const { document } = harness();
    document.setCell(0, 0, 'typed');
    expect(document.canUndo).toBe(true);

    document.addSheet();
    expect(document.canUndo).toBe(false);
  });
});
