/**
 * What the application can be asked to do, as a table.
 *
 * Three things read this and no two of them may disagree: the menu
 * bar, the toolbar, and the shortcut sheet. Written as three lists,
 * they drift — a key gets rebound and the menu goes on advertising the
 * old one, which is worse than advertising nothing, because somebody
 * learns it.
 *
 * So the accelerator is written here once and the sheet of shortcuts
 * is *generated* from it. What that cannot do by itself is agree with
 * `SheetKeys`, which had a table of its own before this file existed
 * and is still the thing a key press actually goes through. The two
 * are reconciled by a spec rather than by a refactor: `viaKeyTable`
 * marks the accelerators `keyAction` already answers, and
 * `SheetCommands.spec.ts` fails the build when the menu's label and
 * the key table's behaviour stop meaning the same thing.
 */

import { isPrintable, keyAction, type KeyModifiers } from './SheetKeys';

export type CommandId =
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'
  | 'clear'
  | 'selectAll'
  | 'fillDown'
  | 'fillRight'
  | 'find'
  | 'replace'
  | 'gotoCell'
  | 'recalculate'
  | 'shortcuts'
  | 'menuBar';

/**
 * How long a chain the proof command builds.
 *
 * Two hundred thousand cells, each depending on the one before it, so
 * the recalculation cannot be parallelised or skipped — it is the
 * longest possible critical path through the sheet, and none of it is
 * on screen. It lives here rather than in the screen that draws the
 * button because the menu item and the button have to say the same
 * number.
 */
export const STRESS_CELLS = 200_000;

/**
 * A key and the modifiers held with it.
 *
 * `ctrl` means ctrl *or* meta, as it does everywhere else in this
 * application — the two are the same accelerator wearing a different
 * hat on a different keyboard, and a table that separated them would
 * need every row twice.
 */
export interface Accelerator {
  readonly key: string;
  readonly ctrl?: boolean;
  readonly shift?: boolean;
}

export interface Command {
  readonly id: CommandId;
  readonly label: string;
  readonly accelerator?: Accelerator;
  /**
   * True when `keyAction` already answers this accelerator.
   *
   * Those rows are documentation here and behaviour there, which is
   * why `commandFor` below refuses to claim them: a key answered
   * twice is a cell that gets cleared and then cleared again, or an
   * undo that takes back two steps.
   */
  readonly viaKeyTable?: boolean;
  /**
   * A command reachable by its key and deliberately in no menu.
   *
   * There is exactly one: the key that *opens* the menu bar cannot
   * sensibly live inside it. It is still advertised by the shortcut
   * sheet, which reads this table rather than the menus, and still
   * held to the rules by the spec.
   */
  readonly hidden?: true;
}

export const COMMANDS: Readonly<Record<CommandId, Command>> = {
  undo: { id: 'undo', label: 'Undo', accelerator: { key: 'z', ctrl: true }, viaKeyTable: true },
  redo: { id: 'redo', label: 'Redo', accelerator: { key: 'y', ctrl: true }, viaKeyTable: true },
  cut: { id: 'cut', label: 'Cut', accelerator: { key: 'x', ctrl: true }, viaKeyTable: true },
  copy: { id: 'copy', label: 'Copy', accelerator: { key: 'c', ctrl: true }, viaKeyTable: true },
  paste: { id: 'paste', label: 'Paste', accelerator: { key: 'v', ctrl: true }, viaKeyTable: true },
  clear: { id: 'clear', label: 'Clear contents', accelerator: { key: 'Delete' }, viaKeyTable: true },
  selectAll: { id: 'selectAll', label: 'Select all', accelerator: { key: 'a', ctrl: true }, viaKeyTable: true },
  fillDown: { id: 'fillDown', label: 'Fill down', accelerator: { key: 'd', ctrl: true } },
  fillRight: { id: 'fillRight', label: 'Fill right', accelerator: { key: 'r', ctrl: true } },
  find: { id: 'find', label: 'Find…', accelerator: { key: 'f', ctrl: true } },
  replace: { id: 'replace', label: 'Replace…', accelerator: { key: 'h', ctrl: true } },
  gotoCell: { id: 'gotoCell', label: 'Go to…', accelerator: { key: 'g', ctrl: true } },
  recalculate: {
    id: 'recalculate',
    label: `Recalculate ${STRESS_CELLS.toLocaleString('en-US')} cells`,
    accelerator: { key: 'F9' }
  },
  shortcuts: { id: 'shortcuts', label: 'Keyboard shortcuts…', accelerator: { key: '/', ctrl: true } },
  menuBar: { id: 'menuBar', label: 'Go to the menu bar', accelerator: { key: 'F10' }, hidden: true }
};

/** A rule drawn across a menu. Not choosable, and not a tab stop. */
export const SEPARATOR = '-' as const;

export type MenuEntry = CommandId | typeof SEPARATOR;

export interface MenuDefinition {
  readonly id: string;
  readonly label: string;
  /**
   * The letter Alt opens this menu with, and the one drawn underlined.
   *
   * Written out rather than taken as the label's first character,
   * because two menus will eventually start with the same letter and
   * the one that loses would silently stop being reachable.
   */
  readonly mnemonic: string;
  readonly entries: readonly MenuEntry[];
}

/**
 * The bar, as it stands after Phase 8.
 *
 * Three menus and not the seven a spreadsheet ends up with, because
 * the other four would be menus of things that do not work yet. File
 * arrives with Phase 16, which is when there is a file to open;
 * Insert and Format arrive with Phases 10 and 9. A menu of disabled
 * items is a worse answer than no menu: it advertises, and then it
 * refuses.
 */
export const MENUS: readonly MenuDefinition[] = [
  {
    id: 'edit',
    label: 'Edit',
    mnemonic: 'e',
    entries: [
      'undo',
      'redo',
      SEPARATOR,
      'cut',
      'copy',
      'paste',
      SEPARATOR,
      'clear',
      'selectAll',
      SEPARATOR,
      'find',
      'replace',
      'gotoCell'
    ]
  },
  {
    id: 'data',
    label: 'Data',
    mnemonic: 'd',
    entries: ['fillDown', 'fillRight', SEPARATOR, 'recalculate']
  },
  {
    id: 'help',
    label: 'Help',
    mnemonic: 'h',
    entries: ['shortcuts']
  }
];

/**
 * The accelerator as a person reads it: `Ctrl+Z`, `Delete`, `F9`.
 *
 * One-character keys are upper-cased because that is how every
 * application in the world prints them, and `Ctrl+z` looks like a
 * mistake even though it is the more literal answer.
 */
export function acceleratorLabel(accelerator: Accelerator): string {
  const parts: string[] = [];
  if (accelerator.ctrl === true) {
    parts.push('Ctrl');
  }
  if (accelerator.shift === true) {
    parts.push('Shift');
  }
  parts.push(accelerator.key.length === 1 ? accelerator.key.toUpperCase() : accelerator.key);
  return parts.join('+');
}

/**
 * The command a key press means, or null when it means nothing here.
 *
 * Only the accelerators `keyAction` does *not* already answer.
 * Claiming one it does would run the command twice — the grid asks
 * this first and the key table second, and both would say yes.
 */
export function commandFor(key: string, modifiers: KeyModifiers): CommandId | null {
  const accel = modifiers.ctrl === true || modifiers.meta === true;
  const shift = modifiers.shift === true;
  for (const command of Object.values(COMMANDS)) {
    const wanted = command.accelerator;
    if (wanted === undefined || command.viaKeyTable === true) {
      continue;
    }
    if (
      matchesKey(wanted.key, key) &&
      (wanted.ctrl === true) === accel &&
      (wanted.shift === true) === shift
    ) {
      return command.id;
    }
  }
  return null;
}

/**
 * Whether a key press is the key an accelerator asked for.
 *
 * Single characters compare case-insensitively: with ctrl held, a
 * browser reports `d` or `D` depending on whether caps lock is on,
 * and an accelerator that stopped working under caps lock is the kind
 * of bug nobody can describe.
 */
function matchesKey(wanted: string, pressed: string): boolean {
  return wanted.length === 1 ? wanted.toLowerCase() === pressed.toLowerCase() : wanted === pressed;
}

/**
 * Whether a key opens a menu from the bar by its letter.
 *
 * Alt on its own focuses the bar, which is the platform convention
 * and is handled by the screen; this is Alt+E and its siblings.
 */
export function menuForMnemonic(key: string): number {
  if (!isPrintable(key)) {
    return -1;
  }
  return MENUS.findIndex(menu => menu.mnemonic === key.toLowerCase());
}

/**
 * What the key table does with an accelerator, for the spec that
 * reconciles the two.
 *
 * Exported so the check lives in a spec rather than in a comment
 * asking the next person to remember.
 */
export function keyTableAction(accelerator: Accelerator) {
  return keyAction(accelerator.key, { ctrl: accelerator.ctrl === true, shift: accelerator.shift === true }, false);
}
