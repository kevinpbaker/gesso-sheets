import { describe, expect, it } from 'vitest';

import { Sheet } from './Sheet';
import { CIRC, DIV0, NAME, REF, VALUE } from './Values';

/** A sheet built from a grid literal, recalculated once. */
function sheetOf(grid: Record<string, string>): Sheet {
  const sheet = new Sheet();
  for (const [address, input] of Object.entries(grid)) {
    const column = address.charCodeAt(0) - 65;
    sheet.setCell(Number(address.slice(1)) - 1, column, input);
  }
  sheet.recalculate();
  return sheet;
}

describe('typing into a cell', () => {
  it('reads something that looks like a number as a number', () => {
    const sheet = sheetOf({ A1: '42', A2: '-1.5', A3: '1e3' });
    expect(sheet.value(0, 0)).toBe(42);
    expect(sheet.value(1, 0)).toBe(-1.5);
    expect(sheet.value(2, 0)).toBe(1000);
  });

  it('reads anything else as text', () => {
    const sheet = sheetOf({ A1: 'hello', A2: '3 apples', A3: '0x10' });
    expect(sheet.value(0, 0)).toBe('hello');
    expect(sheet.value(1, 0)).toBe('3 apples');
    // Number('0x10') is 16, which is not what somebody who typed it meant.
    expect(sheet.value(2, 0)).toBe('0x10');
  });

  it('reads TRUE and FALSE as booleans', () => {
    const sheet = sheetOf({ A1: 'TRUE', A2: 'false' });
    expect(sheet.value(0, 0)).toBe(true);
    expect(sheet.value(1, 0)).toBe(false);
  });

  it('keeps what was typed, which is what an editor reopens with', () => {
    const sheet = sheetOf({ A1: '1', B1: '=A1+1' });
    expect(sheet.input(0, 0)).toBe('1');
    expect(sheet.input(0, 1)).toBe('=A1+1');
    expect(sheet.display(0, 1)).toBe('2');
  });

  it('empties a cell when the text is emptied', () => {
    const sheet = sheetOf({ A1: '5' });
    expect(sheet.size).toBe(1);
    sheet.setCell(0, 0, '');
    expect(sheet.value(0, 0)).toBeNull();
    expect(sheet.size).toBe(0);
  });

  /**
   * Losing what somebody wrote because they have not finished writing
   * it would be worse than showing an error beside it.
   */
  it('keeps a formula that does not parse, and marks it', () => {
    const sheet = sheetOf({ A1: '=1+' });
    expect(sheet.value(0, 0)).toBe(VALUE);
    expect(sheet.input(0, 0)).toBe('=1+');
  });
});

describe('recalculation', () => {
  it('follows a chain in order', () => {
    const sheet = sheetOf({ A1: '1', A2: '=A1+1', A3: '=A2+1', A4: '=A3+1' });
    expect(sheet.value(3, 0)).toBe(4);

    sheet.setCell(0, 0, '10');
    sheet.recalculate();

    expect(sheet.value(3, 0)).toBe(13);
  });

  it('updates a formula when a cell it reads through a range changes', () => {
    const sheet = sheetOf({ A1: '1', A2: '2', A3: '3', B1: '=SUM(A1:A3)' });
    expect(sheet.value(0, 1)).toBe(6);

    sheet.setCell(1, 0, '20');
    sheet.recalculate();

    expect(sheet.value(0, 1)).toBe(24);
  });

  it('notices a cell that was empty and is now not', () => {
    const sheet = sheetOf({ B1: '=SUM(A1:A3)' });
    expect(sheet.value(0, 1)).toBe(0);

    sheet.setCell(0, 0, '5');
    sheet.recalculate();

    expect(sheet.value(0, 1)).toBe(5);
  });

  it('notices a cell that was cleared', () => {
    const sheet = sheetOf({ A1: '5', B1: '=A1+1' });
    sheet.clearCell(0, 0);
    sheet.recalculate();
    expect(sheet.value(0, 1)).toBe(1);
  });

  /**
   * Retyping a formula has to drop the references it no longer has, or
   * the old precedent goes on marking it dirty forever — a leak that
   * shows up as a recalc that grows with the session rather than with
   * the sheet.
   */
  it('stops following a reference a rewritten formula dropped', () => {
    const sheet = sheetOf({ A1: '1', B1: '2', C1: '=A1+B1' });
    expect(sheet.value(0, 2)).toBe(3);

    sheet.setCell(0, 2, '=B1');
    sheet.recalculate();
    expect(sheet.dependentsOf(0, 0)).toEqual([]);

    const before = sheet.stats.evaluated;
    sheet.setCell(0, 0, '99');
    sheet.recalculate();
    expect(sheet.stats.evaluated).toBe(before);
    expect(sheet.value(0, 2)).toBe(2);
  });

  it('turns a formula back into a literal', () => {
    const sheet = sheetOf({ A1: '1', B1: '=A1+1' });
    sheet.setCell(0, 1, '7');
    sheet.recalculate();
    expect(sheet.value(0, 1)).toBe(7);

    const before = sheet.stats.evaluated;
    sheet.setCell(0, 0, '5');
    sheet.recalculate();
    expect(sheet.stats.evaluated).toBe(before);
    expect(sheet.value(0, 1)).toBe(7);
  });
});

describe('cycles', () => {
  it('marks a cell that reads itself', () => {
    const sheet = sheetOf({ A1: '=A1+1' });
    expect(sheet.value(0, 0)).toBe(CIRC);
  });

  it('marks every cell in a ring', () => {
    const sheet = sheetOf({ A1: '=B1', B1: '=C1', C1: '=A1' });
    expect(sheet.value(0, 0)).toBe(CIRC);
    expect(sheet.value(0, 1)).toBe(CIRC);
    expect(sheet.value(0, 2)).toBe(CIRC);
  });

  it('marks what reads a ring, which has no value either', () => {
    const sheet = sheetOf({ A1: '=B1', B1: '=A1', C1: '=A1+1' });
    expect(sheet.value(0, 2)).toBe(CIRC);
  });

  it('leaves the rest of the sheet alone', () => {
    const sheet = sheetOf({ A1: '=B1', B1: '=A1', D1: '2', E1: '=D1*2' });
    expect(sheet.value(0, 4)).toBe(4);
  });

  it('recovers when the cycle is broken', () => {
    const sheet = sheetOf({ A1: '=B1', B1: '=A1' });
    expect(sheet.value(0, 0)).toBe(CIRC);

    sheet.setCell(0, 1, '5');
    sheet.recalculate();

    expect(sheet.value(0, 1)).toBe(5);
    expect(sheet.value(0, 0)).toBe(5);
  });

  it('does not leave a cycle pending forever', () => {
    const sheet = new Sheet();
    sheet.setCell(0, 0, '=B1');
    sheet.setCell(0, 1, '=A1');
    const result = sheet.recalculate();
    expect(result.done).toBe(true);
    expect(sheet.pending).toBe(0);
  });
});

describe('the five error values', () => {
  it('names a reference off the sheet', () => {
    expect(sheetOf({ A1: '=XFE1' }).value(0, 0)).toBe(REF);
  });

  it('names a division by zero', () => {
    expect(sheetOf({ A1: '1', B1: '0', C1: '=A1/B1' }).value(0, 2)).toBe(DIV0);
  });

  it('names something it does not know', () => {
    expect(sheetOf({ A1: '=FROBNICATE(1)' }).value(0, 0)).toBe(NAME);
  });

  it('names a value of the wrong kind', () => {
    expect(sheetOf({ A1: 'text', B1: '=A1*2' }).value(0, 1)).toBe(VALUE);
  });

  it('names a circular reference', () => {
    expect(sheetOf({ A1: '=A1' }).value(0, 0)).toBe(CIRC);
  });

  it('shows them as their codes', () => {
    expect(sheetOf({ A1: '=1/0' }).display(0, 0)).toBe('#DIV/0!');
  });

  /**
   * A cell may legitimately hold the *text* `#REF!` — somebody pasted
   * it out of a report — and a sheet that could not tell that apart
   * from the error could not be trusted about either.
   */
  it('tells the text apart from the error', () => {
    const sheet = sheetOf({ A1: '#REF!', B1: '=1/0' });
    // The same eight characters on the screen, and not the same value:
    // one is text a person typed, the other is an error the engine
    // produced, and only the second should propagate.
    expect(sheet.display(0, 0)).toBe('#REF!');
    expect(sheet.value(0, 0)).toBe('#REF!');
    expect(sheet.value(0, 1)).toBe(DIV0);
    sheet.setCell(1, 0, '=A1&"!"');
    sheet.setCell(1, 1, '=B1&"!"');
    sheet.recalculate();
    expect(sheet.value(1, 0)).toBe('#REF!!');
    expect(sheet.value(1, 1)).toBe(DIV0);
  });
});

describe('reading the sheet back', () => {
  it('lists what it holds, for a repository to write out', () => {
    const sheet = sheetOf({ A1: '1', B2: '=A1+1' });
    expect([...sheet.entries()].sort((a, b) => a.row - b.row)).toEqual([
      { row: 0, column: 0, input: '1' },
      { row: 1, column: 1, input: '=A1+1' }
    ]);
  });

  it('says what a cell reads and what reads it', () => {
    const sheet = sheetOf({ A1: '1', B1: '=A1+1', C1: '=B1+1' });
    expect(sheet.precedentsOf(0, 1)).toHaveLength(1);
    expect(sheet.dependentsOf(0, 1)).toHaveLength(1);
    expect(sheet.dependentsOf(0, 2)).toHaveLength(0);
  });
});
