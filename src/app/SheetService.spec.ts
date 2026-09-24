import { describe, expect, it } from 'vitest';

import { cellIn, type SheetEditor, type SheetStatus, type SheetWindow } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';

function harness(budget = 2_000) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, budget });
  const drain = () => {
    let guard = 0;
    while (queue.length > 0) {
      queue.shift()!();
      if (guard++ > 10_000) {
        throw new Error('the pump never finished');
      }
    }
  };
  return { document, service, drain, queue };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

describe('SheetService', () => {
  it('publishes an empty window until a viewport arrives', () => {
    const { service } = harness();
    expect(latest<SheetWindow>(service.window).lastRow).toBe(-1);
  });

  it('publishes the cells a viewport covers, as display strings', () => {
    const { service, drain } = harness();
    service.setCell(0, 0, '=1+2');
    drain();
    service.setViewport(0, 1, 0, 1);

    const window = latest<SheetWindow>(service.window);
    expect(cellIn(window, 0, 0)).toBe('3');
    expect(cellIn(window, 1, 1)).toBe('');
  });

  it('shows an error as its code, because the wire carries strings', () => {
    const { service, drain } = harness();
    service.setViewport(0, 0, 0, 0);
    service.setCell(0, 0, '=1/0');
    drain();
    expect(cellIn(latest<SheetWindow>(service.window), 0, 0)).toBe('#DIV/0!');
  });

  it('moves the formula bar with the selection, and not the window', () => {
    const { service, drain } = harness();
    service.setViewport(0, 4, 0, 4);
    service.setCell(1, 1, '=2*3');
    drain();
    const windowBefore = latest<SheetWindow>(service.window);

    service.setSelection(1, 1, 1, 1);

    expect(latest<SheetEditor>(service.editor)).toEqual({ row: 1, column: 1, input: '=2*3' });
    expect(latest<SheetWindow>(service.window)).toBe(windowBefore);
  });

  it('reports what is still to do, and that it settles', () => {
    const { document, service, drain } = harness(100);
    // Built through the model so the pump never runs, then woken by
    // one command: a backlog that outlasts a single slice.
    document.sheet.setCell(0, 0, '1');
    for (let row = 1; row <= 500; row++) {
      document.sheet.setCell(row, 0, `=A${row}+1`);
    }
    document.sheet.recalculate();
    service.setViewport(0, 4, 0, 0);

    service.setCell(0, 0, '2');

    expect(latest<SheetStatus>(service.status).pending).toBeGreaterThan(0);
    expect(service.recalculating).toBe(true);
    drain();
    expect(latest<SheetStatus>(service.status).pending).toBe(0);
    expect(service.recalculating).toBe(false);
  });

  it('reports whether there is anything to undo', () => {
    const { service, drain } = harness();
    expect(latest<SheetStatus>(service.status).canUndo).toBe(false);
    service.setCell(0, 0, 'a');
    drain();
    expect(latest<SheetStatus>(service.status).canUndo).toBe(true);
    service.undo();
    drain();
    expect(latest<SheetStatus>(service.status).canUndo).toBe(false);
    expect(latest<SheetStatus>(service.status).canRedo).toBe(true);
  });

  it('puts an undone value back in the window', () => {
    const { service, drain } = harness();
    service.setViewport(0, 0, 0, 0);
    service.setCell(0, 0, 'before');
    drain();
    service.setCell(0, 0, 'after');
    drain();
    expect(cellIn(latest<SheetWindow>(service.window), 0, 0)).toBe('after');

    service.undo();
    drain();

    expect(cellIn(latest<SheetWindow>(service.window), 0, 0)).toBe('before');
  });

  /**
   * The pump must not be re-entered. A second edit arriving while a
   * recalc is in slices adds to the same queue; starting a second pump
   * would run two interleaved walks of one plan.
   */
  it('runs one pump however many edits arrive', () => {
    const { service, queue, drain } = harness(10);
    for (let row = 1; row <= 200; row++) {
      service.setCell(row, 0, `=A${row}+1`);
    }
    expect(queue.length).toBeLessThanOrEqual(1);
    drain();
    expect(service.recalculating).toBe(false);
  });
});
