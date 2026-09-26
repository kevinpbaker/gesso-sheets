import { describe, expect, it } from 'vitest';

import { COMMANDS } from './SheetCommands';
import { isPrintable, keyAction, NAVIGATION, PAGE_ROWS, stampText, type SheetAction } from './SheetKeys';

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
    expect(onGrid('c', { ctrl: true })).toEqual({ kind: 'copy', cut: false });
    expect(onGrid('x', { meta: true })).toEqual({ kind: 'copy', cut: true });
    expect(onGrid('a', { ctrl: true })).toEqual({ kind: 'selectAll' });
  });

  /**
   * Paste is not a key here. The text arrives from the system a moment
   * after the key goes down, as a Paste event, and a handler that
   * claimed the key would have nothing to put anywhere.
   */
  it('leaves paste to the event that carries the text', () => {
    expect(onGrid('v', { ctrl: true })).toBeNull();
  });

  it('moves down on Enter and right on Tab, and back with shift', () => {
    // Shift reverses these two rather than extending, which is what it
    // means on Enter and Tab.
    expect(onGrid('Enter')).toEqual({ kind: 'move', rows: 1, columns: 0, extend: false });
    expect(onGrid('Enter', { shift: true })).toEqual({ kind: 'move', rows: -1, columns: 0, extend: false });
    expect(onGrid('Tab')).toEqual({ kind: 'move', rows: 0, columns: 1, extend: false, tab: true });
    expect(onGrid('Tab', { shift: true })).toEqual({ kind: 'move', rows: 0, columns: -1, extend: false, tab: true });
  });

  it('pages by less than a screen, so something stays in common', () => {
    expect(onGrid('PageDown')).toEqual({ kind: 'move', rows: PAGE_ROWS, columns: 0, extend: false });
    expect(onGrid('PageUp')).toEqual({ kind: 'move', rows: -PAGE_ROWS, columns: 0, extend: false });
  });

  it('jumps to the edges', () => {
    expect(onGrid('Home')).toEqual({ kind: 'jump', to: 'rowStart', extend: false });
    expect(onGrid('End')).toEqual({ kind: 'jump', to: 'rowEnd', extend: false });
    expect(onGrid('Home', { ctrl: true })).toEqual({ kind: 'jump', to: 'sheetStart', extend: false });
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
    expect(onGrid('z', { ctrl: true })).toEqual({ kind: 'undo' });
    expect(onGrid('Z', { meta: true })).toEqual({ kind: 'undo' });
    expect(onGrid('z', { ctrl: true, shift: true })).toEqual({ kind: 'redo' });
    expect(onGrid('y', { ctrl: true })).toEqual({ kind: 'redo' });
  });

  it('lets a key it has no meaning for through', () => {
    expect(onGrid('Shift')).toBeNull();
    expect(onGrid('F5')).toBeNull();
    expect(onGrid('Escape')).toBeNull();
    expect(onGrid('q', { ctrl: true })).toBeNull();
  });
});

describe('keys inside an open cell', () => {
  it('commits and moves down on Enter, up with shift', () => {
    expect(inCell('Enter')).toEqual({ kind: 'commit', rows: 1, columns: 0 });
    expect(inCell('Enter', { shift: true })).toEqual({ kind: 'commit', rows: -1, columns: 0 });
  });

  it('commits and moves right on Tab, left with shift', () => {
    expect(inCell('Tab')).toEqual({ kind: 'commit', rows: 0, columns: 1, tab: true });
    expect(inCell('Tab', { shift: true })).toEqual({ kind: 'commit', rows: 0, columns: -1, tab: true });
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
    expect(inCell('z', { ctrl: true })).toBeNull();
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

/**
 * The shortcut sheet advertises `NAVIGATION`, and a help page that
 * advertises a key which does nothing is worse than no help page:
 * somebody believes it. Pressing every advertised key here is what
 * makes the sheet answerable to the table beside it.
 */
describe('the keys the shortcut sheet advertises', () => {
  it('all do something', () => {
    for (const entry of NAVIGATION) {
      const action = keyAction(entry.probe, entry.probeModifiers ?? {}, entry.whileEditing === true);
      expect(action, `${entry.keys} (${entry.label})`).not.toBeNull();
    }
  });

  /**
   * The other direction: a key the table answers and the sheet does not
   * advertise is a key nobody presses. Every kind of thing a key can
   * do has to be reachable from an advertised key — a navigation entry,
   * or a command whose accelerator the table answers — so a new key
   * cannot be added here without somebody saying so in the help.
   */
  it('advertises every kind of thing a key can do', () => {
    const advertised = new Set<string>();
    // An entry that is not only for an open cell says what the key does
    // in one as well — "commit and move down" — so it is asked both ways.
    for (const entry of NAVIGATION) {
      for (const editing of entry.whileEditing === true ? [true] : [false, true]) {
        const action = keyAction(entry.probe, entry.probeModifiers ?? {}, editing);
        if (action !== null) {
          advertised.add(action.kind);
        }
      }
    }
    for (const command of Object.values(COMMANDS)) {
      if (command.viaKeyTable === true && command.accelerator !== undefined) {
        const { key, ...modifiers } = command.accelerator;
        const action = keyAction(key, modifiers, false);
        if (action !== null) {
          advertised.add(action.kind);
        }
      }
    }
    // Typing is not a shortcut, and the help does not list the alphabet.
    advertised.add('replace');

    const keys = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown', 'Enter', 'Tab'];
    keys.push('Escape', 'F2', 'Delete', 'Backspace', ' ', ';', ':', 'a', 'c', 'x', 'y', 'z');
    const chords = [{}, { shift: true }, { ctrl: true }, { ctrl: true, shift: true }, { alt: true }, { meta: true }];
    for (const key of keys) {
      for (const modifiers of chords) {
        for (const editing of [false, true]) {
          const action = keyAction(key, modifiers, editing);
          if (action !== null) {
            expect(advertised, `${JSON.stringify(modifiers)} ${key}${editing ? ' in a cell' : ''}`).toContain(action.kind);
          }
        }
      }
    }
  });

  it('says something about every one of them', () => {
    for (const entry of NAVIGATION) {
      expect(entry.label.length).toBeGreaterThan(0);
      expect(entry.keys.length).toBeGreaterThan(0);
    }
  });
});

describe('the new keys', () => {
  it('jumps to the edge of the data on Ctrl+Arrow, extending with Shift', () => {
    expect(onGrid('ArrowDown', { ctrl: true })).toEqual({ kind: 'edge', rows: 1, columns: 0, extend: false });
    expect(onGrid('ArrowLeft', { meta: true, shift: true })).toEqual({ kind: 'edge', rows: 0, columns: -1, extend: true });
  });

  it('leaves Ctrl+Arrow to the text in an open cell', () => {
    expect(inCell('ArrowRight', { ctrl: true })).toBeNull();
  });

  it('selects a column on Ctrl+Space and a row on Shift+Space', () => {
    expect(onGrid(' ', { ctrl: true })).toEqual({ kind: 'selectLine', axis: 'columns' });
    expect(onGrid(' ', { shift: true })).toEqual({ kind: 'selectLine', axis: 'rows' });
    // A plain space is still the first character of what is typed.
    expect(onGrid(' ')).toEqual({ kind: 'replace', text: ' ' });
  });

  it('stamps the date on Ctrl+; and the time on either spelling of Ctrl+Shift+;', () => {
    for (const editing of [false, true]) {
      expect(keyAction(';', { ctrl: true }, editing)).toEqual({ kind: 'stamp', what: 'date' });
      expect(keyAction(':', { ctrl: true, shift: true }, editing)).toEqual({ kind: 'stamp', what: 'time' });
      expect(keyAction(';', { ctrl: true, shift: true }, editing)).toEqual({ kind: 'stamp', what: 'time' });
    }
  });

  it('breaks the line on Alt+Enter and fills the selection on Ctrl+Enter, in a cell', () => {
    expect(inCell('Enter', { alt: true })).toEqual({ kind: 'insert', text: '\n' });
    expect(inCell('Enter', { ctrl: true })).toEqual({ kind: 'commitAll' });
  });

  it('writes the stamp as the sheet reads it, in local time', () => {
    const at = new Date(2026, 0, 5, 9, 7);
    expect(stampText('date', at)).toBe('2026-01-05');
    expect(stampText('time', at)).toBe('09:07');
  });
});
