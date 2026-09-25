import { describe, expect, it } from 'vitest';

import {
  acceleratorLabel,
  COMMANDS,
  commandFor,
  keyTableAction,
  MENUS,
  menuForMnemonic,
  SEPARATOR,
  type CommandId
} from './SheetCommands';
import { keyAction } from './SheetKeys';

/**
 * The spec that keeps two tables from drifting apart.
 *
 * `SheetKeys` says what a key press does and `SheetCommands` says what
 * the menu advertises. Nothing in the language makes those agree, and
 * the failure is the quiet kind: the menu goes on printing `Ctrl+Z`
 * next to Undo long after the key has been rebound, and somebody
 * learns the wrong thing from the application itself.
 *
 * So every accelerator marked `viaKeyTable` is pressed here, through
 * the real key table, and asserted to mean what the menu says it
 * means. A rebinding that does not update both fails this file.
 */
describe('the command table', () => {
  /** What each key-table accelerator is claimed to do, in the key table's own terms. */
  const claims: Partial<Record<CommandId, (action: ReturnType<typeof keyAction>) => boolean>> = {
    undo: action => action?.kind === 'undo',
    redo: action => action?.kind === 'redo',
    cut: action => action?.kind === 'copy' && action.cut,
    copy: action => action?.kind === 'copy' && !action.cut,
    clear: action => action?.kind === 'clear',
    selectAll: action => action?.kind === 'selectAll',
    // Paste is the odd one: the key table deliberately answers null,
    // because at the moment the key goes down there is nothing to
    // paste. The text arrives a moment later as a Paste event. The
    // menu still advertises Ctrl+V, and that is correct.
    paste: action => action === null
  };

  it('advertises the key the key table actually answers', () => {
    for (const command of Object.values(COMMANDS)) {
      if (command.viaKeyTable !== true) {
        continue;
      }
      const accelerator = command.accelerator;
      expect(accelerator, `${command.id} is marked viaKeyTable with no accelerator`).toBeDefined();
      const claim = claims[command.id];
      expect(claim, `${command.id} is marked viaKeyTable but nothing here says what it should do`).toBeDefined();
      expect(claim!(keyTableAction(accelerator!)), `${command.id}: ${acceleratorLabel(accelerator!)}`).toBe(true);
    }
  });

  it('never claims an accelerator the key table already answers', () => {
    for (const command of Object.values(COMMANDS)) {
      if (command.viaKeyTable !== true || command.accelerator === undefined) {
        continue;
      }
      const { key, ctrl, shift } = command.accelerator;
      expect(commandFor(key, { ctrl: ctrl === true, shift: shift === true }), command.id).toBeNull();
    }
  });

  it('answers the accelerators that are its own', () => {
    expect(commandFor('d', { ctrl: true })).toBe('fillDown');
    expect(commandFor('r', { ctrl: true })).toBe('fillRight');
    expect(commandFor('f', { ctrl: true })).toBe('find');
    expect(commandFor('h', { ctrl: true })).toBe('replace');
    expect(commandFor('g', { ctrl: true })).toBe('gotoCell');
    expect(commandFor('F9', {})).toBe('recalculate');
    expect(commandFor('/', { ctrl: true })).toBe('shortcuts');
    expect(commandFor('F10', {})).toBe('menuBar');
  });

  it('takes meta for ctrl, so one table covers both keyboards', () => {
    expect(commandFor('f', { meta: true })).toBe('find');
  });

  /**
   * Caps lock reports `D` where the table wrote `d`. An accelerator
   * that stops working under caps lock is a bug nobody can describe,
   * so it is a spec rather than a hope.
   */
  it('matches a letter whatever case the browser reports', () => {
    expect(commandFor('D', { ctrl: true })).toBe('fillDown');
  });

  it('does not fire on the bare letter', () => {
    expect(commandFor('d', {})).toBeNull();
    expect(commandFor('f', {})).toBeNull();
  });

  it('does not fire when an extra modifier is held', () => {
    expect(commandFor('d', { ctrl: true, shift: true })).toBeNull();
  });

  it('prints an accelerator the way an application prints one', () => {
    expect(acceleratorLabel({ key: 'z', ctrl: true })).toBe('Ctrl+Z');
    expect(acceleratorLabel({ key: 'Delete' })).toBe('Delete');
    expect(acceleratorLabel({ key: 'F9' })).toBe('F9');
    expect(acceleratorLabel({ key: 'z', ctrl: true, shift: true })).toBe('Ctrl+Shift+Z');
  });
});

describe('the menus', () => {
  it('names only commands that exist', () => {
    for (const menu of MENUS) {
      for (const entry of menu.entries) {
        if (entry === SEPARATOR) {
          continue;
        }
        expect(COMMANDS[entry], `${menu.id} names ${entry}`).toBeDefined();
      }
    }
  });

  /**
   * A command reachable from two menus is a command somebody will
   * find in the wrong one, and a command in none is one they cannot
   * find at all. The shortcut sheet reads `COMMANDS`, so an orphan
   * would be advertised by the sheet and absent from the bar.
   */
  it('puts every command in exactly one menu, or deliberately in none', () => {
    const seen = new Map<string, number>();
    for (const menu of MENUS) {
      for (const entry of menu.entries) {
        if (entry !== SEPARATOR) {
          seen.set(entry, (seen.get(entry) ?? 0) + 1);
        }
      }
    }
    for (const command of Object.values(COMMANDS)) {
      const wanted = command.hidden === true ? 0 : 1;
      expect(seen.get(command.id) ?? 0, `${command.id} is in ${seen.get(command.id) ?? 0} menus`).toBe(wanted);
    }
  });

  /**
   * The shortcut sheet reads `COMMANDS` rather than `MENUS`, so a
   * hidden command is still advertised. A key that works and is
   * documented nowhere is a key nobody presses.
   */
  it('gives every command an accelerator, so the sheet can advertise it', () => {
    for (const command of Object.values(COMMANDS)) {
      expect(command.accelerator, command.id).toBeDefined();
    }
  });

  it('gives every menu a mnemonic of its own', () => {
    const letters = MENUS.map(menu => menu.mnemonic);
    expect(new Set(letters).size).toBe(letters.length);
    for (const menu of MENUS) {
      expect(menu.label.toLowerCase()).toContain(menu.mnemonic);
      expect(menuForMnemonic(menu.mnemonic.toUpperCase())).toBe(MENUS.indexOf(menu));
    }
  });

  it('has no menu that opens onto a rule', () => {
    for (const menu of MENUS) {
      expect(menu.entries[0], menu.id).not.toBe(SEPARATOR);
      expect(menu.entries[menu.entries.length - 1], menu.id).not.toBe(SEPARATOR);
    }
  });
});
