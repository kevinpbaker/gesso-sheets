import { describe, expect, it } from 'vitest';

import { addressOf, explainCell } from './Explain';
import { Sheet } from './Sheet';

/**
 * Why a cell is showing an error.
 *
 * Five characters is a diagnosis, not an explanation: `#DIV/0!` in a
 * cell that divides nothing means the division is three cells away,
 * and finding it is a mechanical walk backwards through the formulas.
 */
function sheetOf(cells: Readonly<Record<string, string>>): Sheet {
  const sheet = new Sheet();
  for (const [address, input] of Object.entries(cells)) {
    const match = /^([A-Z])(\d+)$/.exec(address)!;
    sheet.setCell(Number(match[2]) - 1, match[1].charCodeAt(0) - 65, input);
  }
  sheet.recalculate();
  return sheet;
}

const at = (sheet: Sheet, address: string) => {
  const match = /^([A-Z])(\d+)$/.exec(address)!;
  return explainCell(sheet, Number(match[2]) - 1, match[1].charCodeAt(0) - 65);
};

describe('explaining an error', () => {
  it('says nothing about a cell that is fine', () => {
    expect(at(sheetOf({ A1: '1' }), 'A1')).toBeNull();
  });

  it('says what the code means', () => {
    const found = at(sheetOf({ A1: '=1/0' }), 'A1');
    expect(found?.code).toBe('#DIV/0!');
    expect(found?.meaning).toContain('divided by zero');
  });

  /** A cell that broke on its own needs no pointing at. */
  it('blames nobody when the cell broke itself', () => {
    expect(at(sheetOf({ A1: '=1/0' }), 'A1')?.blame).toBeNull();
  });

  it('follows the error back to the cell that made it', () => {
    const sheet = sheetOf({ A1: '=1/0', B1: '=A1+1', C1: '=B1*2' });
    expect(at(sheet, 'C1')?.blame).toEqual({ row: 0, column: 0 });
  });

  /** The nearest origin, which is the one somebody can act on. */
  it('stops at the first cell that produced it', () => {
    const sheet = sheetOf({ A1: '=1/0', B1: '=A1', C1: '=B1' });
    expect(at(sheet, 'C1')?.blame).toEqual({ row: 0, column: 0 });
    expect(at(sheet, 'B1')?.blame).toEqual({ row: 0, column: 0 });
  });

  it('does not blame a cell carrying a different error', () => {
    // B1 is #VALUE!, C1 divides by zero on its own.
    const sheet = sheetOf({ A1: 'text', B1: '=A1+1', C1: '=1/0' });
    expect(at(sheet, 'B1')?.code).toBe('#VALUE!');
    expect(at(sheet, 'C1')?.blame).toBeNull();
  });

  it('explains a circular cell without walking forever', () => {
    const sheet = sheetOf({ A1: '=B1', B1: '=A1' });
    const found = at(sheet, 'A1');
    expect(found?.code).toBe('#CIRC!');
    expect(found?.meaning).toContain('depends on itself');
  });

  it('explains a lookup that found nothing', () => {
    const sheet = sheetOf({ A1: '1', B1: '=VLOOKUP(99, A1:A1, 1, FALSE)' });
    expect(at(sheet, 'B1')?.code).toBe('#N/A');
    expect(at(sheet, 'B1')?.meaning).toContain('table has no such row');
  });

  it('explains a misspelt function', () => {
    expect(at(sheetOf({ A1: '=NOSUCH(1)' }), 'A1')?.code).toBe('#NAME?');
  });

  /** A long chain must not cost the selection a visible pause. */
  it('gives up rather than walking a whole sheet', () => {
    const cells: Record<string, string> = { A1: '=1/0' };
    for (let row = 2; row <= 60; row++) {
      cells[`A${row}`] = `=A${row - 1}+1`;
    }
    const sheet = sheetOf(cells);
    // Well within the limit: still traced.
    expect(explainCell(sheet, 59, 0)?.blame).toEqual({ row: 0, column: 0 });
    // And with a limit of two it stops looking rather than hanging.
    expect(explainCell(sheet, 59, 0, 2)?.blame).toBeNull();
  });
});

describe('naming a cell in a sentence', () => {
  it('writes it as somebody would', () => {
    expect(addressOf(0, 0)).toBe('A1');
    expect(addressOf(6, 1)).toBe('B7');
    expect(addressOf(0, 26)).toBe('AA1');
  });
});
