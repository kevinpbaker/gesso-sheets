import { describe, expect, it } from 'vitest';

import { relativeRef, type RangeRef } from '../sheet/A1';
import { addressOf, cycleAbsolute, pick, pickDecision, repick } from './FormulaEditing';

/**
 * Reference picking, as a table of (caret context, click) → result.
 *
 * Phase 12's exit criterion, and written as a table because the thing
 * being specified is a **mode**: the same click either moves the
 * selection or types into somebody's formula. One wrong case in one
 * direction throws away a formula somebody was writing; one wrong
 * case in the other freezes the selection and the sheet feels broken.
 *
 * The caret is written as `|` in the first column.
 */

const cell = (row: number, column: number): RangeRef => ({
  start: relativeRef(row, column),
  end: relativeRef(row, column)
});

const area = (row: number, column: number, lastRow: number, lastColumn: number): RangeRef => ({
  start: relativeRef(row, column),
  end: relativeRef(lastRow, lastColumn)
});

/** Splits `=SUM(|)` into the text and the caret offset. */
function at(withCaret: string): [string, number] {
  const caret = withCaret.indexOf('|');
  return [withCaret.slice(0, caret) + withCaret.slice(caret + 1), caret];
}

const decisionFor = (withCaret: string) => pickDecision(...at(withCaret));

describe('what a click does, by where the caret is', () => {
  /**
   * The positions where a formula is expecting a value. A click here
   * types an address.
   */
  const inserts: readonly string[] = [
    '=|',
    '=1+|',
    '=1-|',
    '=1*|',
    '=1/|',
    '=2^|',
    '="a"&|',
    '=SUM(|',
    '=SUM(A1,|',
    '=SUM(A1, |',
    '=IF(A1>|',
    '=IF(A1<=|',
    '=SUM(A1:|'
  ];

  it.each(inserts)('inserts at %o', text => {
    expect(decisionFor(text).kind).toBe('insert');
  });

  /**
   * The positions where the formula is finished enough that a click
   * means "I am done here". Inserting would produce `=1A1`.
   */
  const selects: readonly string[] = [
    '=1|',
    '=SUM(A1)|',
    '="text"|',
    '=A1+1|',
    '=SUM(A1) |',
    // Not a formula at all.
    'North|',
    '120|',
    '|'
  ];

  it.each(selects)('selects at %o', text => {
    expect(decisionFor(text).kind).toBe('select');
  });

  /**
   * The caret on a reference replaces it, which is how somebody fixes
   * a formula pointing at the wrong cell without retyping it.
   */
  const replaces: readonly [string, number, number][] = [
    ['=|A1', 1, 3],
    ['=A|1', 1, 3],
    ['=A1|', 1, 3],
    ['=SUM(A1:B9|)', 5, 10],
    ['=SUM(|A1:B9)', 5, 10],
    ['=B2*|C2', 4, 6]
  ];

  it.each(replaces)('replaces at %o, over [%i, %i]', (text, start, end) => {
    const decision = decisionFor(text);
    expect(decision).toEqual({ kind: 'replace', span: { start, end } });
  });
});

describe('inserting an address', () => {
  it('writes a cell at the caret', () => {
    expect(pick(...at('=SUM(|'), cell(1, 1))).toEqual({
      text: '=SUM(B2',
      caret: 7,
      span: { start: 5, end: 7 }
    });
  });

  it('writes a range when one was dragged', () => {
    expect(pick(...at('=SUM(|'), area(1, 1, 6, 3))?.text).toBe('=SUM(B2:D7');
  });

  /** Relative, always: a click is not a statement about copying. */
  it('writes it without dollars', () => {
    expect(addressOf(cell(0, 0))).toBe('A1');
    expect(addressOf(area(0, 0, 8, 1))).toBe('A1:B9');
  });

  it('writes over the reference the caret was on', () => {
    expect(pick(...at('=A1|+1'), cell(4, 2))).toEqual({
      text: '=C5+1',
      caret: 3,
      span: { start: 1, end: 3 }
    });
  });

  it('writes over a range, however much shorter the new one is', () => {
    expect(pick(...at('=SUM(A1:Z99|)'), cell(0, 0))?.text).toBe('=SUM(A1)');
  });

  it('refuses where a click is only a click', () => {
    expect(pick(...at('=1|'), cell(0, 0))).toBeNull();
  });
});

/**
 * A drag rewrites the address it already wrote, rather than adding
 * another corner each time the pointer moves.
 */
describe('dragging a range into a formula', () => {
  it('grows the address in place', () => {
    const down = pick(...at('=SUM(|'), cell(1, 1))!;
    expect(down.text).toBe('=SUM(B2');

    const moved = repick(down.text, down.span, area(1, 1, 3, 1));
    expect(moved.text).toBe('=SUM(B2:B4');

    const further = repick(moved.text, moved.span, area(1, 1, 6, 3));
    expect(further.text).toBe('=SUM(B2:D7');
    expect(further.caret).toBe(further.text.length);
  });

  it('leaves what was already in the formula alone', () => {
    const down = pick(...at('=SUM(A1, |)'), cell(4, 4))!;
    expect(down.text).toBe('=SUM(A1, E5)');

    const moved = repick(down.text, down.span, area(4, 4, 5, 4));
    expect(moved.text).toBe('=SUM(A1, E5:E6)');
  });
});

/**
 * F4, in Excel's order — which is not the order anybody would choose,
 * and is the one people's fingers know.
 */
describe('cycling a reference through its dollars', () => {
  const cycle = (withCaret: string) => cycleAbsolute(...at(withCaret))?.text;

  it('goes A1 to $A$1 to A$1 to $A1 and round again', () => {
    expect(cycle('=A|1')).toBe('=$A$1');
    expect(cycle('=$A$|1')).toBe('=A$1');
    expect(cycle('=A$|1')).toBe('=$A1');
    expect(cycle('=$A|1')).toBe('=A1');
  });

  it('leaves the rest of the formula where it was', () => {
    expect(cycle('=SUM(A|1, B2)')).toBe('=SUM($A$1, B2)');
  });

  it('moves both corners of a range together', () => {
    expect(cycle('=SUM(A1:B|9)')).toBe('=SUM($A$1:$B$9)');
  });

  it('puts the caret after what it rewrote', () => {
    expect(cycleAbsolute(...at('=A|1'))?.caret).toBe(5);
  });

  it('does nothing when the caret is not on a reference', () => {
    expect(cycleAbsolute(...at('=1+|1'))).toBeNull();
    expect(cycleAbsolute(...at('North|'))).toBeNull();
  });
});
