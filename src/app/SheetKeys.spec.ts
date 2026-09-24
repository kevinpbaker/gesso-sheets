import { describe, expect, it } from 'vitest';

import { isPrintable, keyAction, PAGE_ROWS, type SheetAction } from './SheetKeys';

function onGrid(key: string, modifiers = {}): SheetAction | null {
  return keyAction(key, modifiers, false);
}

function inCell(key: string, modifiers = {}): SheetAction | null {
  return keyAction(key, modifiers, true);
}

describe('keys on a selected cell', () => {
  it('moves with the arrows', () => {
    expect(onGrid('ArrowUp')).toEqual({ kind: 'move', rows: -1, columns: 0, extend: false });
    expect(onGrid('ArrowDown')).toEqual({ kind: 'move', rows: 1, columns: 0, extend: false });
    expect(onGrid('ArrowLeft')).toEqual({ kind: 'move', rows: 0, columns: -1, extend: false });
    expect(onGrid('ArrowRight')).toEqual({ kind: 'move', rows: 0, columns: 1, extend: false });
  });

  /** Shift keeps the anchor and moves the far corner, as it does everywhere. */
  it('extends the selection when shift is held', () => {
    expect(onGrid('ArrowDown', { shift: true })).toEqual({ kind: 'move', rows: 1, columns: 0, extend: true });
    expect(onGrid('ArrowRight', { shift: true })).toEqual({ kind: 'move', rows: 0, columns: 1, extend: true });
    expect(onGrid('End', { shift: true })).toEqual({ kind: 'jump', to: 'rowEnd', extend: true });
    expect(onGrid('PageDown', { shift: true })).toEqual({ kind: 'move', rows: PAGE_ROWS, columns: 0, extend: true });
  });

  it('copies, cuts and selects everything', () => {
    expect(onGrid('c', { control: true })).toEqual({ kind: 'copy', cut: false });
    expect(onGrid('x', { meta: true })).toEqual({ kind: 'copy', cut: true });
    expect(onGrid('a', { control: true })).toEqual({ kind: 'selectAll' });
  });

  /**
   * Paste is not a key here. The text arrives from the system a moment
   * after the key goes down, as a Paste event, and a handler that
   * claimed the key would have nothing to put anywhere.
   */
  it('leaves paste to the event that carries the text', () => {
    expect(onGrid('v', { control: true })).toBeNull();
  });

  it('moves down on Enter and right on Tab, and back with shift', () => {
    // Shift reverses these two rather than extending, which is what it
    // means on Enter and Tab.
    expect(onGrid('Enter')).toEqual({ kind: 'move', rows: 1, columns: 0, extend: false });
    expect(onGrid('Enter', { shift: true })).toEqual({ kind: 'move', rows: -1, columns: 0, extend: false });
    expect(onGrid('Tab')).toEqual({ kind: 'move', rows: 0, columns: 1, extend: false });
    expect(onGrid('Tab', { shift: true })).toEqual({ kind: 'move', rows: 0, columns: -1, extend: false });
  });

  it('pages by less than a screen, so something stays in common', () => {
    expect(onGrid('PageDown')).toEqual({ kind: 'move', rows: PAGE_ROWS, columns: 0, extend: false });
    expect(onGrid('PageUp')).toEqual({ kind: 'move', rows: -PAGE_ROWS, columns: 0, extend: false });
  });

  it('jumps to the edges', () => {
    expect(onGrid('Home')).toEqual({ kind: 'jump', to: 'rowStart', extend: false });
    expect(onGrid('End')).toEqual({ kind: 'jump', to: 'rowEnd', extend: false });
    expect(onGrid('Home', { control: true })).toEqual({ kind: 'jump', to: 'sheetStart', extend: false });
    expect(onGrid('End', { meta: true })).toEqual({ kind: 'jump', to: 'sheetEnd', extend: false });
  });

  it('opens the cell on F2 with what is already in it', () => {
    expect(onGrid('F2')).toEqual({ kind: 'edit' });
  });

  /**
   * The character typed is the first character of the new value, not a
   * keystroke spent on opening the cell. Typing `5` over a cell holding
   * `=A1+1` leaves `5` in it, not `=A1+15` and not an empty editor.
   */
  it('replaces the cell with the character typed', () => {
    expect(onGrid('5')).toEqual({ kind: 'replace', text: '5' });
    expect(onGrid('=')).toEqual({ kind: 'replace', text: '=' });
    expect(onGrid('£')).toEqual({ kind: 'replace', text: '£' });
  });

  it('empties the cell on Delete without opening it', () => {
    expect(onGrid('Delete')).toEqual({ kind: 'clear' });
    expect(onGrid('Backspace')).toEqual({ kind: 'clear' });
  });

  it('undoes and redoes', () => {
    expect(onGrid('z', { control: true })).toEqual({ kind: 'undo' });
    expect(onGrid('Z', { meta: true })).toEqual({ kind: 'undo' });
    expect(onGrid('z', { control: true, shift: true })).toEqual({ kind: 'redo' });
    expect(onGrid('y', { control: true })).toEqual({ kind: 'redo' });
  });

  it('lets a key it has no meaning for through', () => {
    expect(onGrid('Shift')).toBeNull();
    expect(onGrid('F5')).toBeNull();
    expect(onGrid('Escape')).toBeNull();
    expect(onGrid('q', { control: true })).toBeNull();
  });
});

describe('keys inside an open cell', () => {
  it('commits and moves down on Enter, up with shift', () => {
    expect(inCell('Enter')).toEqual({ kind: 'commit', rows: 1, columns: 0 });
    expect(inCell('Enter', { shift: true })).toEqual({ kind: 'commit', rows: -1, columns: 0 });
  });

  it('commits and moves right on Tab, left with shift', () => {
    expect(inCell('Tab')).toEqual({ kind: 'commit', rows: 0, columns: 1 });
    expect(inCell('Tab', { shift: true })).toEqual({ kind: 'commit', rows: 0, columns: -1 });
  });

  it('puts back what was there on Escape', () => {
    expect(inCell('Escape')).toEqual({ kind: 'cancel' });
  });

  /**
   * The important null. Once a cell is open the caret, the selection,
   * the clipboard and IME composition all belong to the text; a grid
   * that went on reading arrows as movement would be a grid you cannot
   * type an arrow key into, and one that read `z` as undo would eat
   * the letter.
   */
  it('gives every other key to the text', () => {
    for (const key of ['ArrowLeft', 'ArrowUp', 'Home', 'End', 'Delete', 'Backspace', 'a', 'F2', '5']) {
      expect(inCell(key), key).toBeNull();
    }
    expect(inCell('z', { control: true })).toBeNull();
  });
});

describe('isPrintable', () => {
  it('is one code point, not one UTF-16 unit', () => {
    expect(isPrintable('a')).toBe(true);
    expect(isPrintable('€')).toBe(true);
    // Two units, one character: a key a touch keyboard can send.
    expect(isPrintable('😀')).toBe(true);
    expect(isPrintable('ArrowUp')).toBe(false);
    expect(isPrintable('')).toBe(false);
  });
});
