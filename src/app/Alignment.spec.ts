import { describe, expect, it } from 'vitest';

import type { CellPaint } from '../sheet/Format';
import type { SheetFormatWindow, SheetPalette } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';

/**
 * Where a cell sits when nobody has said: numbers right, text left.
 *
 * The grid decides that from the display string, which is all it is
 * sent, and a number format is exactly what makes the string mislead.
 * These are the cells where it would, read through what the service
 * publishes — which is where the fix is.
 */

function harness() {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, rowCount: 100, columnCount: 10 });
  service.setViewport(0, 0, 20, 0, 5);
  const drain = () => {
    while (queue.length > 0) {
      queue.shift()!();
    }
  };
  /** The paint the grid would draw a cell with. */
  const paintAt = (row: number, column: number): CellPaint => {
    drain();
    let formats!: SheetFormatWindow;
    let palette!: SheetPalette;
    service.formats.subscribe(value => (formats = value)).unsubscribe();
    service.palette.subscribe(value => (palette = value)).unsubscribe();
    return palette.entries[formats.cells[row]?.[column] ?? 0];
  };
  const formats = (): SheetFormatWindow => {
    drain();
    let seen!: SheetFormatWindow;
    service.formats.subscribe(value => (seen = value)).unsubscribe();
    return seen;
  };
  return { service, document, drain, paintAt, formats };
}

describe('the alignment a format would hide', () => {
  it('puts a number formatted as currency on the right', () => {
    const { service, paintAt } = harness();
    service.setCell(0, 0, '4.5');
    service.format({ number: { kind: 'currency', places: 2, symbol: '$' } });
    expect(paintAt(0, 0).align).toBe('end');
  });

  it('puts a percentage and a date on the right', () => {
    const { service, paintAt } = harness();
    service.setCell(0, 0, '0.096');
    service.format({ number: { kind: 'percent', places: 1 } });
    service.setSelection(1, 0, 1, 0);
    service.setCell(1, 0, '2026-09-24');
    expect(paintAt(0, 0).align).toBe('end');
    expect(paintAt(1, 0).align).toBe('end');
  });

  /** A code with its leading zero is text, and text sits on the left. */
  it('puts text that looks like a number on the left', () => {
    const { service, paintAt } = harness();
    service.format({ number: { kind: 'text' } });
    service.setCell(0, 0, '007');
    expect(paintAt(0, 0).align).toBe('start');
  });

  it('leaves an alignment somebody chose alone', () => {
    const { service, paintAt } = harness();
    service.setCell(0, 0, '4.5');
    service.format({ number: { kind: 'currency', places: 2, symbol: '$' }, align: 'center' });
    expect(paintAt(0, 0).align).toBe('center');
  });

  it('follows the value when a formula under a format stops being a number', () => {
    const { service, paintAt } = harness();
    service.setCell(0, 1, '2');
    service.setCell(0, 0, '=B1*2');
    service.format({ number: { kind: 'currency', places: 2, symbol: '$' } });
    expect(paintAt(0, 0).align).toBe('end');
    service.setCell(0, 1, 'two');
    expect(paintAt(0, 0).align).toBe('auto');
  });

  /** What keeps this off `pnpm proof`: a plain sheet sends no formats at all. */
  it('sends nothing for a sheet nobody has formatted', () => {
    const { service, formats } = harness();
    service.setCell(0, 0, '42');
    service.setCell(0, 1, 'text');
    service.setCell(0, 2, '=A1*2');
    const cells = formats().cells;
    expect(Object.values(cells).every(line => Object.keys(line).length === 0)).toBe(true);
  });
});
