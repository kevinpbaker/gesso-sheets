import { describe, expect, it } from 'vitest';

import { DEFAULT_FORMAT, GENERAL, PLAIN, type CellFormat } from '../sheet/Format';

import { COLUMN_WIDTH } from './dimensions';
import { applySnapshot, parseSnapshot, snapshotOf, type SheetSnapshot } from './SheetFile';
import { relativeRef } from '../sheet/A1';
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

    const snapshot = snapshotOf(document);

    // The formula, not the 20 it displayed: a value can be rebuilt
    // from a formula and a formula cannot be rebuilt from a value.
    expect(snapshot.sheets[0].cells).toContainEqual({ row: 0, column: 1, input: '=A1*10' });
  });

  it('comes back as the sheet it was', () => {
    const before = new SheetDocument();
    before.setCell(0, 0, '2');
    before.setCell(0, 1, '=A1*10');
    before.sheet.recalculate();

    const after = new SheetDocument();
    applySnapshot(after, snapshotOf(before));

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
      version: 3,
      active: 0,
      names: [],
      sheets: [
        {
          name: 'Sheet1',
          colour: null,
          cells: [{ row: 0, column: 0, input: 'a' }],
          palette: [],
          formats: [],
          regions: { sheet: 0, rows: [], columns: [] },
          merges: [],
          conditional: [],
          validations: [],
          frozenRows: 0,
          frozenColumns: 0,
          hiddenRows: [],
          columnWidths: []
        }
      ]
    });
    expect(document.canUndo).toBe(false);
  });
});

describe('reading a file that is not what this build writes', () => {
  /**
   * A v2 file is a workbook of one sheet, and is read as one.
   *
   * The whole of v3's compatibility claim, and it is a claim about
   * files already on somebody's disk — so it is asserted against the
   * literal shape v2 wrote rather than against anything this build
   * can still produce.
   */
  it('reads a v2 file as a workbook of one sheet', () => {
    const v2 = {
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
    const read = parseSnapshot(JSON.stringify(v2), 4);
    expect(read?.version).toBe(3);
    expect(read?.sheets).toHaveLength(1);
    expect(read?.sheets[0].name).toBe('Sheet1');
    expect(read?.sheets[0].cells).toEqual(v2.cells);
    expect(read?.sheets[0].columnWidths).toEqual([80, 90, COLUMN_WIDTH, COLUMN_WIDTH]);
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
    expect(read?.sheets[0].cells).toEqual([{ row: 0, column: 0, input: 'good' }]);
    expect(read?.sheets[0].columnWidths).toEqual([60, COLUMN_WIDTH, COLUMN_WIDTH]);
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
    expect(repository.peek()?.sheets[0].cells).toContainEqual({ row: 0, column: 0, input: 'from the seed' });
  });

  it('does not seed over a sheet that exists', async () => {
    const repository = new InMemorySheetRepository({
      version: 3,
      active: 0,
      names: [],
      sheets: [
        {
          name: 'Sheet1',
          colour: null,
          cells: [{ row: 0, column: 0, input: 'mine' }],
          palette: [],
          formats: [],
          regions: { sheet: 0, rows: [], columns: [] },
          merges: [],
          conditional: [],
          validations: [],
          frozenRows: 0,
          frozenColumns: 0,
          hiddenRows: [],
          columnWidths: []
        }
      ]
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

    const snapshot = snapshotOf(document, 100);
    expect(snapshot.version).toBe(3);
    expect(snapshot.sheets[0].palette).toHaveLength(3);
    expect(snapshot.sheets[0].formats).toHaveLength(2);
  });

  it('brings a formatted sheet back exactly as it was', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '1234.5');
    document.setFormat(0, 0, money);
    document.setFormat(2, 0, bold);

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, 100)), 4)!);

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
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, 100)), 4)!);
    expect(loaded.display(0, 0)).toBe('007');
  });

  /** A v1 file is a v2 file with no formats in it, so it still opens. */
  it('reads a file written before formats existed', () => {
    const read = parseSnapshot(
      JSON.stringify({ version: 1, cells: [{ row: 0, column: 0, input: '5' }], columnWidths: [] }),
      4
    );
    expect(read?.sheets[0].cells).toEqual([{ row: 0, column: 0, input: '5' }]);
    expect(read?.sheets[0].formats).toEqual([]);
    expect(read?.sheets[0].palette).toHaveLength(1);
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
    const entry = read!.sheets[0].palette[1];
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
    expect(read?.sheets[0].formats).toEqual([]);
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
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, 100)), 4)!);

    expect(loaded.merges.at(0, 2)).toEqual({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 3 });
  });

  it('brings back the frozen pane and the hidden rows', () => {
    const document = new SheetDocument();
    document.frozenRows = 2;
    document.frozenColumns = 1;
    document.hiddenRows.add(4);

    const loaded = new SheetDocument();
    applySnapshot(loaded, parseSnapshot(JSON.stringify(snapshotOf(document, 100)), 4)!);

    expect(loaded.frozenRows).toBe(2);
    expect(loaded.frozenColumns).toBe(1);
    expect([...loaded.hiddenRows]).toEqual([4]);
  });

  it('reads a file written before any of them existed', () => {
    const read = parseSnapshot(
      JSON.stringify({ version: 2, cells: [{ row: 0, column: 0, input: '5' }], columnWidths: [] }),
      4
    );
    expect(read?.sheets[0].merges).toEqual([]);
    expect(read?.sheets[0].frozenRows).toBe(0);
    expect(read?.sheets[0].hiddenRows).toEqual([]);
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
    expect(read?.sheets[0].merges).toEqual([]);
  });

  /** What lives past the end of the sheet is not somebody's data. */
  it('does not write a merge past the end of the sheet', () => {
    const document = new SheetDocument();
    document.merges.add({ firstRow: 500, lastRow: 501, firstColumn: 0, lastColumn: 1 });
    expect(snapshotOf(document, 100).sheets[0].merges).toEqual([]);
  });
});

/**
 * Phase 13's other exit criterion: Phase 6's reload proof, run over a
 * three-sheet workbook with cross-references in both directions.
 *
 * Written down, read back, and every claim made against the reloaded
 * workbook rather than against the snapshot — a file that round-trips
 * its own JSON proves nothing about whether it holds a workbook.
 */
describe('a workbook of three sheets, written down and opened again', () => {
  function built(): SheetDocument {
    const document = new SheetDocument();
    document.renameSheet(0, 'Input');
    document.setCell(0, 0, '10');
    document.setCell(1, 0, '20');
    document.columnWidths = [140, COLUMN_WIDTH];
    document.frozenRows = 1;
    document.merges.add({ firstRow: 4, lastRow: 4, firstColumn: 0, lastColumn: 2 });

    document.addSheet('Working');
    // Reads back down the book…
    document.setCell(0, 0, '=SUM(Input!A1:A2)');
    document.setFormat(0, 0, {
      ...document.formatAt(0, 0),
      number: { kind: 'currency', places: 2, symbol: '$' }
    });
    document.hiddenRows.add(7);

    document.addSheet('Report');
    document.setCell(0, 0, '=Working!A1*2');
    document.setSheetColour(2, '#34a853');

    // …and back up it, so the references run both ways.
    document.activate(0);
    document.setCell(3, 0, '=Report!A1+1');
    document.defineName('Readings', {
      start: { ...relativeRef(0, 0), sheet: 'Input' },
      end: { ...relativeRef(1, 0), sheet: 'Input' }
    });
    document.activate(1);
    document.setCell(2, 0, '=COUNT(Readings)');
    document.activate(2);
    document.book.recalculate();
    return document;
  }

  function reopened(): SheetDocument {
    const written = JSON.stringify(snapshotOf(built(), 1_000));
    const read = parseSnapshot(written, 2);
    expect(read).not.toBeNull();
    const after = new SheetDocument();
    applySnapshot(after, read!);
    return after;
  }

  it('comes back with all three sheets, named and in order', () => {
    const after = reopened();
    expect(after.sheets().map(entry => entry.name)).toEqual(['Input', 'Working', 'Report']);
  });

  it('opens on the sheet that was showing', () => {
    expect(reopened().active).toBe(2);
  });

  it('works out the whole chain again, across all three', () => {
    const after = reopened();
    after.activate(0);
    expect(after.sheet.value(0, 0)).toBe(10);
    after.activate(1);
    expect(after.sheet.value(0, 0)).toBe(30);
    after.activate(2);
    expect(after.sheet.value(0, 0)).toBe(60);
  });

  /** The reference that runs back up the book, which a naive load breaks. */
  it('resolves a reference pointing at a sheet written after it', () => {
    const after = reopened();
    after.activate(0);
    expect(after.sheet.value(3, 0)).toBe(61);
  });

  it('keeps the formulas as text, not the numbers they showed', () => {
    const after = reopened();
    after.activate(1);
    expect(after.sheet.input(0, 0)).toBe('=SUM(Input!A1:A2)');
  });

  it('brings back a name that points at a sheet', () => {
    const after = reopened();
    after.activate(1);
    expect(after.sheet.value(2, 0)).toBe(2);
    expect(after.book.names.rangeOf('Readings')?.start.sheet).toBe('Input');
  });

  /**
   * The names arrive after the cells, so the formulas that read them
   * were wired when the name meant nothing — and a formula that is
   * *evaluated* correctly but not *wired* is the silent kind of
   * wrong: right on the screen until somebody edits what it reads.
   *
   * So the claim is about the edge and not about the value, which is
   * why this writes into the named range rather than reading it.
   */
  it('wires a formula to the name it reads, not only evaluates it', () => {
    const after = reopened();
    after.activate(0);
    after.setCell(2, 0, '30');
    after.book.recalculate();

    after.activate(1);
    expect(after.sheet.value(2, 0)).toBe(2);
    after.activate(0);
    after.setCell(1, 0, '');
    after.book.recalculate();
    after.activate(1);
    expect(after.sheet.value(2, 0)).toBe(1);
  });

  it('keeps what is drawn over each sheet with that sheet', () => {
    const after = reopened();
    after.activate(0);
    expect(after.columnWidths[0]).toBe(140);
    expect(after.frozenRows).toBe(1);
    expect(after.merges.size).toBe(1);
    expect(after.hiddenRows.size).toBe(0);

    after.activate(1);
    expect(after.columnWidths[0]).toBe(COLUMN_WIDTH);
    expect(after.frozenRows).toBe(0);
    expect(after.merges.size).toBe(0);
    expect(after.hiddenRows.has(7)).toBe(true);
  });

  it('keeps each sheet’s formats to itself', () => {
    const after = reopened();
    after.activate(1);
    expect(after.formatAt(0, 0).number).toEqual({ kind: 'currency', places: 2, symbol: '$' });
    expect(after.display(0, 0)).toBe('$30.00');

    after.activate(0);
    expect(after.formatAt(0, 0).number).toEqual({ kind: 'general' });
  });

  it('keeps a tab colour', () => {
    expect(reopened().sheets()[2].colour).toBe('#34a853');
  });

  /** Opening a workbook is not an edit, however many sheets it has. */
  it('does not land on the undo stack', () => {
    expect(reopened().canUndo).toBe(false);
  });
});

/**
 * Phase 6's reload proof, run through the service and the repository
 * over a workbook of three sheets.
 *
 * The document-level specs above say the snapshot is right. This
 * says the *round trip* is: a real service writes to a real
 * repository, a second service reads it back, and every claim is
 * made against what the second one holds.
 */
describe('a workbook of three sheets, reloaded through the service', () => {
  async function saved(): Promise<InMemorySheetRepository> {
    const repository = new InMemorySheetRepository();
    const first = harness(repository);
    await first.service.restore();

    first.service.renameSheet(0, 'Input');
    first.service.setCell(0, 0, '10');
    first.service.setCell(1, 0, '20');
    first.service.setColumnWidth(0, 140);

    first.service.addSheet();
    first.service.renameSheet(1, 'Working');
    first.service.setCell(0, 0, '=SUM(Input!A1:A2)');

    first.service.addSheet();
    first.service.renameSheet(2, 'Report');
    first.service.setSheetColour(2, '#34a853');
    first.service.setCell(0, 0, '=Working!A1*2');

    first.service.activateSheet(0);
    first.service.setCell(3, 0, '=Report!A1+1');
    first.drain();
    return repository;
  }

  it('comes back with every sheet, and works the chain out again', async () => {
    const second = harness(await saved());
    await second.service.restore();
    second.drain();

    expect(second.document.sheets().map(entry => entry.name)).toEqual(['Input', 'Working', 'Report']);

    second.document.activate(1);
    expect(second.document.sheet.value(0, 0)).toBe(30);
    second.document.activate(2);
    expect(second.document.sheet.value(0, 0)).toBe(60);
    // The reference that runs back up the book.
    second.document.activate(0);
    expect(second.document.sheet.value(3, 0)).toBe(61);
  });

  it('keeps each sheet’s own widths and colour', async () => {
    const second = harness(await saved());
    await second.service.restore();
    second.drain();

    second.document.activate(0);
    expect(second.document.columnWidths[0]).toBe(140);
    second.document.activate(1);
    expect(second.document.columnWidths[0]).toBe(COLUMN_WIDTH);
    expect(second.document.sheets()[2].colour).toBe('#34a853');
  });

  it('tells the render worker about the tabs', async () => {
    const second = harness(await saved());
    await second.service.restore();
    second.drain();

    let tabs: readonly { name: string }[] = [];
    second.service.sheets.subscribe(view => (tabs = view.entries)).unsubscribe();
    expect(tabs.map(tab => tab.name)).toEqual(['Input', 'Working', 'Report']);
  });
});

/**
 * The rules, written down and read back.
 *
 * A file is untrusted input in exactly the way a keystroke is, so
 * every rule is rebuilt field by field and one that is
 * half-understood is dropped — half a rule paints the wrong cells
 * rather than none.
 */
describe('formats that think, in a file', () => {
  function withRules(): SheetDocument {
    const document = new SheetDocument();
    document.setCell(0, 0, '1');
    document.setCell(1, 0, '9');
    document.addConditional({
      range: { start: relativeRef(0, 0), end: relativeRef(4, 0) },
      test: { kind: 'greaterThan', value: 5 },
      paint: { fill: '#fce8e6', color: '#c5221f' }
    });
    document.addConditional({
      range: { start: relativeRef(0, 1), end: relativeRef(4, 1) },
      test: null,
      scale: { from: '#ffffff', middle: '#cccccc', to: '#000000' }
    });
    document.addValidation({
      range: { start: relativeRef(0, 2), end: relativeRef(4, 2) },
      rule: { kind: 'list', values: ['North', 'South'] },
      strict: true,
      message: 'A region, please.'
    });
    return document;
  }

  const reopened = (): SheetDocument => {
    const read = parseSnapshot(JSON.stringify(snapshotOf(withRules(), 1_000)), 4);
    expect(read).not.toBeNull();
    const after = new SheetDocument();
    applySnapshot(after, read!);
    return after;
  };

  it('brings the conditional formats back', () => {
    const after = reopened();
    expect(after.conditional).toHaveLength(2);
    expect(after.conditional[0].test).toEqual({ kind: 'greaterThan', value: 5 });
    expect(after.conditional[0].paint).toEqual({ fill: '#fce8e6', color: '#c5221f' });
  });

  it('brings a colour scale back, middle stop and all', () => {
    expect(reopened().conditional[1].scale).toEqual({
      from: '#ffffff',
      middle: '#cccccc',
      to: '#000000'
    });
  });

  it('brings the validations back, and whether they refuse', () => {
    const after = reopened();
    expect(after.validations).toHaveLength(1);
    expect(after.validations[0].rule).toEqual({ kind: 'list', values: ['North', 'South'] });
    expect(after.validations[0].strict).toBe(true);
    expect(after.validations[0].message).toBe('A region, please.');
  });

  it('is absent from a file written before they existed', () => {
    const stored = JSON.parse(JSON.stringify(snapshotOf(withRules()))) as {
      sheets: Record<string, unknown>[];
    };
    delete stored.sheets[0].conditional;
    delete stored.sheets[0].validations;
    const read = parseSnapshot(JSON.stringify(stored), 4);
    expect(read?.sheets[0].conditional).toEqual([]);
    expect(read?.sheets[0].validations).toEqual([]);
  });

  it('drops a rule a file should not have held', () => {
    const stored = JSON.parse(JSON.stringify(snapshotOf(withRules()))) as {
      sheets: Record<string, unknown>[];
    };
    stored.sheets[0].conditional = [
      { range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, test: { kind: 'nonsense' } },
      { test: { kind: 'notEmpty' } },
      {
        range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
        test: { kind: 'notEmpty' },
        paint: { fill: '#eeeeee' }
      }
    ];
    stored.sheets[0].validations = [
      { range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, rule: { kind: 'list', values: [] } },
      { range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } }, rule: { kind: 'number', min: 3 } }
    ];
    const read = parseSnapshot(JSON.stringify(stored), 4);
    expect(read?.sheets[0].conditional.map(rule => rule.test?.kind)).toEqual(['notEmpty']);
    expect(read?.sheets[0].validations.map(rule => rule.rule.kind)).toEqual(['number']);
  });
});
