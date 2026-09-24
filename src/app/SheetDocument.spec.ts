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
