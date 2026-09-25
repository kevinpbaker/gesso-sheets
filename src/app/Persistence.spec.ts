import { describe, expect, it } from 'vitest';

import { DEFAULT_FORMAT, GENERAL, PLAIN, type CellFormat } from '../sheet/Format';

import { COLUMN_WIDTH } from './dimensions';
import { applySnapshot, parseSnapshot, snapshotOf, type SheetSnapshot } from './SheetFile';
import { SheetDocument } from './SheetDocument';
import { InMemorySheetRepository } from './SheetRepository';
import { SheetService, type Schedule } from './SheetService';

/**
 * Persistence — Phase 6.
 *
 * The exit criterion is the notes example's, borrowed: type a marker,
 * reload, it is still there. A reload is a new document, a new
 * service, and the same repository, so that is what these do.
 */

function harness(repository = new InMemorySheetRepository()) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, repository, rowCount: 100, columnCount: 10 });
  const drain = () => {
    while (queue.length > 0) {
      queue.shift()!();
    }
  };
  return { document, service, repository, drain };
}

describe('what is written down', () => {
  it('keeps what was typed, not what was shown', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2');
    document.setCell(0, 1, '=A1*10');
    document.sheet.recalculate();

    const snapshot = snapshotOf(document, [COLUMN_WIDTH]);

    // The formula, not the 20 it displayed: a value can be rebuilt
    // from a formula and a formula cannot be rebuilt from a value.
    expect(snapshot.cells).toContainEqual({ row: 0, column: 1, input: '=A1*10' });
  });

  it('comes back as the sheet it was', () => {
    const before = new SheetDocument();
    before.setCell(0, 0, '2');
    before.setCell(0, 1, '=A1*10');
    before.sheet.recalculate();

    const after = new SheetDocument();
    applySnapshot(after, snapshotOf(before, []));

    expect(after.sheet.value(0, 1)).toBe(20);
    expect(after.sheet.input(0, 1)).toBe('=A1*10');
  });

  /**
   * Opening a sheet is not an edit. A person's first ctrl-Z after
   * opening one should do nothing, not empty it.
   */
  it('does not land on the undo stack', () => {
    const document = new SheetDocument();
    applySnapshot(document, {
      version: 2,
      cells: [{ row: 0, column: 0, input: 'a' }],
      palette: [],
      formats: [],
      regions: { sheet: 0, rows: [], columns: [] },
      merges: [],
      names: [],
      frozenRows: 0,
      frozenColumns: 0,
      hiddenRows: [],
      columnWidths: []
    });
    expect(document.canUndo).toBe(false);
  });
});

describe('reading a file that is not what this build writes', () => {
  it('reads one that is', () => {
    const snapshot: SheetSnapshot = {
      version: 2,
      cells: [{ row: 1, column: 2, input: '=A1' }],
      palette: [],
      formats: [],
      regions: { sheet: 0, rows: [], columns: [] },
      merges: [],
      names: [],
      frozenRows: 0,
      frozenColumns: 0,
      hiddenRows: [],
      columnWidths: [80, 90]
    };
    const read = parseSnapshot(JSON.stringify(snapshot), 4);
    expect(read?.cells).toEqual(snapshot.cells);
    expect(read?.columnWidths).toEqual([80, 90, COLUMN_WIDTH, COLUMN_WIDTH]);
  });

  /**
   * A file is outside the program: an older build wrote it, or a newer
   * one, or something truncated it. Losing a sheet that is mostly fine
   * over one cell that is not would be the wrong trade.
   */
  it('keeps the cells it can read and drops the ones it cannot', () => {
    const text = JSON.stringify({
      version: 1,
      cells: [{ row: 0, column: 0, input: 'good' }, { row: 'x', column: 0, input: 'bad' }, null, { row: 2 }],
      columnWidths: [60, 'wide', -5]
    });
    const read = parseSnapshot(text, 3);
    expect(read?.cells).toEqual([{ row: 0, column: 0, input: 'good' }]);
    expect(read?.columnWidths).toEqual([60, COLUMN_WIDTH, COLUMN_WIDTH]);
  });

  it('refuses one it does not recognise at all', () => {
    expect(parseSnapshot('not json', 4)).toBeNull();
    expect(parseSnapshot('{}', 4)).toBeNull();
    expect(parseSnapshot(JSON.stringify({ version: 99, cells: [] }), 4)).toBeNull();
  });
});

describe('a sheet that is opened again', () => {
  /** The exit criterion: type a marker, reload, it is still there. */
  it('still has what was typed', async () => {
    const repository = new InMemorySheetRepository();

    const first = harness(repository);
    await first.service.restore();
    first.service.setCell(3, 4, 'a marker');
    first.drain();

    // A reload: a new document, a new service, the same disk.
    const second = harness(repository);
    await second.service.restore();

    expect(second.document.sheet.input(3, 4)).toBe('a marker');
  });

  it('recalculates the formulas rather than storing their answers', async () => {
    const repository = new InMemorySheetRepository();
    const first = harness(repository);
    await first.service.restore();
    first.service.setCell(0, 0, '6');
    first.service.setCell(0, 1, '=A1*7');
    first.drain();

    const second = harness(repository);
    await second.service.restore();
    second.drain();

    expect(second.document.sheet.value(0, 1)).toBe(42);
  });

  it('keeps a column that was dragged wider', async () => {
    const repository = new InMemorySheetRepository();
    const first = harness(repository);
    await first.service.restore();
    first.service.setColumnWidth(2, 200);

    const second = harness(repository);
    await second.service.restore();

    let widths: readonly number[] = [];
    second.service.geometry.subscribe(geometry => (widths = geometry.columnWidths)).unsubscribe();
    expect(widths[2]).toBe(200);
    expect(widths[1]).toBe(COLUMN_WIDTH);
  });

  it('seeds a sheet nobody has opened, and writes the seed down', async () => {
    const repository = new InMemorySheetRepository();
    const { service, document } = harness(repository);

    await service.restore(sheet => sheet.setCell(0, 0, 'from the seed'));

    expect(document.sheet.input(0, 0)).toBe('from the seed');
    expect(repository.peek()?.cells).toContainEqual({ row: 0, column: 0, input: 'from the seed' });
  });

  it('does not seed over a sheet that exists', async () => {
    const repository = new InMemorySheetRepository({
      version: 2,
      cells: [{ row: 0, column: 0, input: 'mine' }],
      palette: [],
      formats: [],
      regions: { sheet: 0, rows: [], columns: [] },
      merges: [],
      names: [],
      frozenRows: 0,
      frozenColumns: 0,
      hiddenRows: [],
      columnWidths: []
    });
    const { service, document } = harness(repository);

    await service.restore(sheet => sheet.setCell(0, 0, 'from the seed'));

    expect(document.sheet.input(0, 0)).toBe('mine');
  });

  /**
   * The race the `restored` flag exists for. An edit arriving before
   * the load finished would otherwise write a half-built document over
   * a real file.
   */
  it('writes nothing before it has finished reading', () => {
    const repository = new InMemorySheetRepository();
    const { service } = harness(repository);

    service.setCell(0, 0, 'too early');

    expect(repository.saves).toBe(0);
  });

  it('writes after every kind of edit', async () => {
    const repository = new InMemorySheetRepository();
    const { service, drain } = harness(repository);
    await service.restore();
    const before = repository.saves;

    service.setCell(0, 0, 'a');
    drain();
    service.undo();
    drain();
    service.clearRange();
    drain();

    expect(repository.saves).toBeGreaterThanOrEqual(before + 3);
  });
});

describe('undoing and then closing the tab', () => {
  /**
   * An undo is an edit as far as the file is concerned. Written out,
   * taking something back and closing the tab would bring it back on
   * the next open — the opposite of what undo promises.
   */
  it('does not bring back what was undone', async () => {
    const repository = new InMemorySheetRepository();
    const first = harness(repository);
    await first.service.restore();
    first.service.setCell(0, 0, 'typed');
    first.drain();
    first.service.undo();
    first.drain();

    const second = harness(repository);
    await second.service.restore();

    expect(second.document.sheet.input(0, 0)).toBe('');
  });
});

/**
 * Phase 9 added formats to the file, which is what `version` was put
 * there for in Phase 6 — the file that needs a version field is the
 * one already on somebody's disk.
 */
describe('formats in the file', () => {
  const bold: CellFormat = { number: GENERAL, paint: { ...PLAIN, bold: true } };
  const money: CellFormat = { number: { kind: 'currency', places: 2, symbol: '$' }, paint: PLAIN };

  it('writes the palette and the cells pointing into it', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '1234.5');
    document.setFormat(0, 0, money);
    document.setFormat(1, 1, bold);

    const snapshot = snapshotOf(document, [], 100);
    expect(snapshot.version).toBe(2);
    expect(snapshot.palette).toHaveLength(3);
    expect(snapshot.formats).toHaveLength(2);
  });

  it('brings a formatted sheet back exactly as it was', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '1234.5');
    document.setFormat(0, 0, money);
    document.setFormat(2, 0, bold);

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, [], 100)), 4)!);

    expect(loaded.display(0, 0)).toBe('$1,234.50');
    expect(loaded.formatAt(2, 0).paint.bold).toBe(true);
    expect(loaded.formatAt(5, 5)).toEqual(DEFAULT_FORMAT);
  });

  /**
   * Formats have to be restored before inputs are read. The other way
   * round, `007` in a Text cell is parsed as the number seven and
   * then formatted as text, and the leading zeros somebody saved are
   * gone by the time the format says to keep them.
   */
  it('keeps what a Text cell was holding', () => {
    const document = new SheetDocument();
    document.setFormat(0, 0, { number: { kind: 'text' }, paint: PLAIN });
    document.setCell(0, 0, '007');
    expect(document.display(0, 0)).toBe('007');

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, [], 100)), 4)!);
    expect(loaded.display(0, 0)).toBe('007');
  });

  /** A v1 file is a v2 file with no formats in it, so it still opens. */
  it('reads a file written before formats existed', () => {
    const read = parseSnapshot(
      JSON.stringify({ version: 1, cells: [{ row: 0, column: 0, input: '5' }], columnWidths: [] }),
      4
    );
    expect(read?.cells).toEqual([{ row: 0, column: 0, input: '5' }]);
    expect(read?.formats).toEqual([]);
    expect(read?.palette).toHaveLength(1);
  });

  /**
   * A file is outside the program. A palette entry missing a field
   * would otherwise reach the render worker as `undefined` and be
   * drawn as nothing at all, so each field falls back rather than the
   * cell being lost.
   */
  it('repairs a palette entry it only half understands', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        palette: [null, { number: { kind: 'currency' }, paint: { bold: true } }],
        formats: [{ row: 0, column: 0, id: 1 }],
        columnWidths: []
      }),
      4
    );
    const entry = read!.palette[1];
    expect(entry.paint.bold).toBe(true);
    expect(entry.paint.align).toBe('auto');
    expect(entry.number).toEqual({ kind: 'currency', places: 2, symbol: '$' });
  });

  it('drops a cell pointing at a palette entry that is not there', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        palette: [],
        formats: [{ row: 0, column: 0, id: 9 }],
        columnWidths: []
      }),
      4
    );
    expect(read?.formats).toEqual([]);
  });
});

/**
 * The shape of a sheet — its merges, its frozen pane, its hidden
 * rows — is as much somebody's work as its cells are, and a sheet
 * reopened without them has lost something they set up.
 *
 * No version bump: an older file has none of these fields, and their
 * absence already means the right thing. A version is for a change
 * that would be read *wrongly*, not one that would be read as
 * nothing.
 */
describe('the shape of the sheet in the file', () => {
  it('brings back the merges', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, 'Title');
    document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 3 });

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, [], 100)), 4)!);

    expect(loaded.merges.at(0, 2)).toEqual({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 3 });
  });

  it('brings back the frozen pane and the hidden rows', () => {
    const document = new SheetDocument();
    document.frozenRows = 2;
    document.frozenColumns = 1;
    document.hiddenRows.add(4);

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, [], 100)), 4)!);

    expect(loaded.frozenRows).toBe(2);
    expect(loaded.frozenColumns).toBe(1);
    expect([...loaded.hiddenRows]).toEqual([4]);
  });

  it('reads a file written before any of them existed', () => {
    const read = parseSnapshot(
      JSON.stringify({ version: 2, cells: [{ row: 0, column: 0, input: '5' }], columnWidths: [] }),
      4
    );
    expect(read?.merges).toEqual([]);
    expect(read?.frozenRows).toBe(0);
    expect(read?.hiddenRows).toEqual([]);
  });

  it('drops a merge a file got wrong', () => {
    const read = parseSnapshot(
      JSON.stringify({
        version: 2,
        cells: [],
        merges: [{ firstRow: 5, lastRow: 1, firstColumn: 0, lastColumn: 1 }, null, { firstRow: 0 }],
        columnWidths: []
      }),
      4
    );
    expect(read?.merges).toEqual([]);
  });

  /** What lives past the end of the sheet is not somebody's data. */
  it('does not write a merge past the end of the sheet', () => {
    const document = new SheetDocument();
    document.merges.add({ firstRow: 500, lastRow: 501, firstColumn: 0, lastColumn: 1 });
    expect(snapshotOf(document, [], 100).merges).toEqual([]);
  });
});
