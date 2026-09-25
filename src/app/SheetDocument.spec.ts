import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';

describe('SheetDocument', () => {
  it('commits a cell and settles it', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '=1+2');
    document.sheet.recalculate();
    expect(document.sheet.display(0, 0)).toBe('3');
    expect(document.sheet.input(0, 0)).toBe('=1+2');
  });

  describe('undo', () => {
    it('puts back what a cell held', () => {
      const document = new SheetDocument();
      document.setCell(0, 0, 'first');
      document.setCell(0, 0, 'second');

      expect(document.undo()).toBe(true);
      expect(document.sheet.input(0, 0)).toBe('first');
      expect(document.undo()).toBe(true);
      expect(document.sheet.input(0, 0)).toBe('');
    });

    it('has nothing to undo at the start', () => {
      const document = new SheetDocument();
      expect(document.canUndo).toBe(false);
      expect(document.undo()).toBe(false);
    });

    it('redoes what it undid', () => {
      const document = new SheetDocument();
      document.setCell(0, 0, 'a');
      document.undo();
      expect(document.canRedo).toBe(true);
      expect(document.redo()).toBe(true);
      expect(document.sheet.input(0, 0)).toBe('a');
    });

    /**
     * Editing after undoing drops the future. Keeping it would offer a
     * redo that reinstates a value on top of an edit made since, which
     * is a way to lose work rather than recover it.
     */
    it('drops the future when an edit replaces it', () => {
      const document = new SheetDocument();
      document.setCell(0, 0, 'a');
      document.undo();
      document.setCell(0, 0, 'b');
      expect(document.canRedo).toBe(false);
    });

    /**
     * A commit that changes nothing is not an edit. Pressing Enter on
     * a cell without touching it would otherwise leave an entry that
     * undoes to itself, and a person pressing ctrl-Z would watch
     * nothing happen and press it again.
     */
    it('records nothing when the text did not change', () => {
      const document = new SheetDocument();
      document.setCell(0, 0, 'a');
      document.setCell(0, 0, 'a');
      document.undo();
      expect(document.sheet.input(0, 0)).toBe('');
      expect(document.canUndo).toBe(false);
    });

    /** Undoing moves the selection to what it changed, so it is seen. */
    it('takes the selection to the cell it put back', () => {
      const document = new SheetDocument();
      document.setCell(4, 6, 'x');
      document.setSelection(0, 0, 0, 0);
      document.undo();
      expect(document.selection).toEqual({ row: 4, column: 6, anchorRow: 4, anchorColumn: 6 });
    });

    it('recalculates what the undone cell fed', () => {
      const document = new SheetDocument();
      document.setCell(0, 0, '1');
      document.setCell(0, 1, '=A1*10');
      document.sheet.recalculate();
      document.setCell(0, 0, '5');
      document.sheet.recalculate();
      expect(document.sheet.value(0, 1)).toBe(50);

      document.undo();
      document.sheet.recalculate();
      expect(document.sheet.value(0, 1)).toBe(10);
    });
  });

  it('reports the active cell as it was typed, not as it displays', () => {
    const document = new SheetDocument();
    document.setCell(2, 3, '=1+2');
    document.sheet.recalculate();
    document.setSelection(2, 3, 2, 3);
    expect(document.activeInput).toBe('=1+2');
    expect(document.sheet.display(2, 3)).toBe('3');
  });
});

/**
 * Typing a date, which is the one place the engine and the format
 * axis touch.
 *
 * A date is a number with a format — Excel's serial from 1899-12-30 —
 * so neither half is optional. The engine turns the text into a
 * serial and the document puts the format on, and a spec that checked
 * only one of them would pass for a cell showing 46,289 to somebody
 * who typed a date.
 */
describe('typing a date', () => {
  it('stores the serial and shows the date', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2026-09-24');
    document.sheet.recalculate();

    expect(document.sheet.value(0, 0)).toBe(46_289);
    expect(document.display(0, 0)).toBe('2026-09-24');
  });

  it('gives it back for editing exactly as it was typed', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '24/9/2026');
    expect(document.sheet.input(0, 0)).toBe('24/9/2026');
  });

  it('shows it in the shape it was typed in', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '24/9/2026');
    document.setCell(1, 0, 'Sep 24, 2026');
    document.setCell(2, 0, '2026-09-24');

    expect(document.display(0, 0)).toBe('24 Sep 2026');
    expect(document.display(1, 0)).toBe('Sep 24, 2026');
    expect(document.display(2, 0)).toBe('2026-09-24');
  });

  it('keeps the clock when one was typed', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2026-09-24 09:30');
    expect(document.display(0, 0)).toBe('2026-09-24 09:30');
  });

  it('shows a time on its own as a time', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '13:45');
    expect(document.display(0, 0)).toBe('13:45');
  });

  /** A date is a number, which is what makes date arithmetic work. */
  it('subtracts two dates into a number of days', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2026-09-01');
    document.setCell(1, 0, '2026-09-24');
    document.setCell(2, 0, '=A2-A1');
    document.sheet.recalculate();

    // The difference is a count, not a date: the formula cell was
    // never formatted, so it shows the number.
    expect(document.display(2, 0)).toBe('23');
  });

  /** Both halves, in one press. */
  it('takes the value and the format back on one undo', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2026-09-24');
    document.undo();

    expect(document.sheet.input(0, 0)).toBe('');
    expect(document.formatAt(0, 0).number.kind).toBe('general');
  });

  /**
   * A format somebody chose on purpose is an answer already given.
   * This looks odd and is what every spreadsheet does; the
   * alternative is a format that silently undoes a decision.
   */
  it('leaves a format somebody chose alone', () => {
    const document = new SheetDocument();
    document.setFormat(0, 0, { ...document.formatAt(0, 0), number: { kind: 'number', places: 0, thousands: true } });
    document.setCell(0, 0, '2026-09-24');

    expect(document.formatAt(0, 0).number.kind).toBe('number');
    expect(document.display(0, 0)).toBe('46,289');
  });

  /** And text stays text: the number test runs before the date one. */
  it('leaves a bare year as a number', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '2026');
    expect(document.sheet.value(0, 0)).toBe(2026);
    expect(document.formatAt(0, 0).number.kind).toBe('general');
  });
});
