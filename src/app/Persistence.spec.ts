import { describe, expect, it } from 'vitest';

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
    applySnapshot(document, { version: 1, cells: [{ row: 0, column: 0, input: 'a' }], columnWidths: [] });
    expect(document.canUndo).toBe(false);
  });
});

describe('reading a file that is not what this build writes', () => {
  it('reads one that is', () => {
    const snapshot: SheetSnapshot = {
      version: 1,
      cells: [{ row: 1, column: 2, input: '=A1' }],
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
      version: 1,
      cells: [{ row: 0, column: 0, input: 'mine' }],
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
