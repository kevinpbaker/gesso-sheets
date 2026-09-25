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
  | 'menuBar'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'wrap'
  | 'formatGeneral'
  | 'formatNumber'
  | 'formatCurrency'
  | 'formatPercent'
  | 'formatScientific'
  | 'formatDate'
  | 'formatTime'
  | 'formatText'
  | 'moreDecimals'
  | 'fewerDecimals'
  | 'clearFormat'
  | 'insertRowAbove'
  | 'insertRowBelow'
  | 'insertColumnLeft'
  | 'insertColumnRight'
  | 'deleteRows'
  | 'deleteColumns'
  | 'borderAll'
  | 'borderOutline'
  | 'borderTop'
  | 'borderBottom'
  | 'borderThickBottom'
  | 'borderNone';

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
  readonly alt?: boolean;
  /**
   * What the browser actually reports for this key with shift held.
   *
   * A keyboard event's `key` is the character produced, not the key
   * pressed: hold shift and press 4 and a browser says `$`. So the
   * number-format accelerators — Ctrl+Shift+1 through 7, which every
   * spreadsheet binds and which this table would otherwise advertise
   * and never answer — are matched by either. The label is still
   * written from `key`, because `Ctrl+Shift+4` is what people have
   * learned and `Ctrl+Shift+$` is what a bug report looks like.
   *
   * It is layout-dependent and unavoidably so; `code` would be the
   * layout-independent answer and is not on `UiKeyboardEvent`. A
   * layout where neither character is produced loses the shortcut and
   * keeps the menu item, which is the right way round to fail.
   */
  readonly shifted?: string;
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

/**
 * Whether a command can be reached at all.
 *
 * A menu item or a key is enough; a command with neither is one
 * nobody can run and one the shortcut sheet would not list. Checked
 * by the spec rather than by the type, because "in a menu" is a fact
 * about `MENUS` and not about this row.
 */
export function isReachable(command: Command, inAMenu: boolean): boolean {
  return inAMenu || command.accelerator !== undefined;
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
  menuBar: { id: 'menuBar', label: 'Go to the menu bar', accelerator: { key: 'F10' }, hidden: true },

  /**
   * Alt, and not the Ctrl+Shift+= and Ctrl+- a desktop spreadsheet
   * uses.
   *
   * Those are Chrome's zoom shortcuts, and a page cannot prevent
   * them: the browser takes Ctrl+Plus and Ctrl+Minus before the
   * document sees them, so an insert bound there is an insert that
   * zooms the page instead. Google Sheets moved to Ctrl+Alt for the
   * same reason and this follows it — there is no cleverness
   * available, only somebody else's key or nobody's.
   *
   * `=` and not `+`: `+` is what the key *produces* with shift held,
   * which is the trap the number formats fell into. The label prints
   * the key somebody presses.
   */
  insertRowAbove: { id: 'insertRowAbove', label: 'Row above', accelerator: { key: '=', ctrl: true, alt: true } },
  insertRowBelow: { id: 'insertRowBelow', label: 'Row below' },
  insertColumnLeft: {
    id: 'insertColumnLeft',
    label: 'Column left',
    accelerator: { key: '=', ctrl: true, alt: true, shift: true, shifted: '+' }
  },
  insertColumnRight: { id: 'insertColumnRight', label: 'Column right' },
  deleteRows: { id: 'deleteRows', label: 'Delete rows', accelerator: { key: '-', ctrl: true, alt: true } },
  deleteColumns: {
    id: 'deleteColumns',
    label: 'Delete columns',
    accelerator: { key: '-', ctrl: true, alt: true, shift: true, shifted: '_' }
  },

  /**
   * Six borders and not a grid of sixteen buttons.
   *
   * What people actually draw is a box round a block, a rule under a
   * heading and a heavy rule above a total; everything else in a
   * border picker is there because the picker exists. These six are
   * those three, their obvious neighbours, and the one that takes
   * them off again.
   */
  borderAll: { id: 'borderAll', label: 'All borders' },
  borderOutline: { id: 'borderOutline', label: 'Outline' },
  borderTop: { id: 'borderTop', label: 'Top border' },
  borderBottom: { id: 'borderBottom', label: 'Bottom border' },
  borderThickBottom: { id: 'borderThickBottom', label: 'Thick bottom border' },
  borderNone: { id: 'borderNone', label: 'No borders' },

  bold: { id: 'bold', label: 'Bold', accelerator: { key: 'b', ctrl: true } },
  italic: { id: 'italic', label: 'Italic', accelerator: { key: 'i', ctrl: true } },
  underline: { id: 'underline', label: 'Underline', accelerator: { key: 'u', ctrl: true } },
  alignLeft: { id: 'alignLeft', label: 'Align left', accelerator: { key: 'l', ctrl: true, shift: true } },
  alignCenter: { id: 'alignCenter', label: 'Align centre', accelerator: { key: 'e', ctrl: true, shift: true } },
  alignRight: { id: 'alignRight', label: 'Align right', accelerator: { key: 'r', ctrl: true, shift: true } },
  wrap: { id: 'wrap', label: 'Wrap text', accelerator: { key: 'w', ctrl: true, shift: true } },
  /**
   * The number formats take Ctrl+Shift+1 through 7, which is what
   * every spreadsheet binds them to and the one part of this table
   * nobody has to learn.
   */
  formatGeneral: {
    id: 'formatGeneral',
    label: 'General',
    accelerator: { key: '`', ctrl: true, shift: true, shifted: '~' }
  },
  formatNumber: {
    id: 'formatNumber',
    label: 'Number',
    accelerator: { key: '1', ctrl: true, shift: true, shifted: '!' }
  },
  formatCurrency: {
    id: 'formatCurrency',
    label: 'Currency',
    accelerator: { key: '4', ctrl: true, shift: true, shifted: '$' }
  },
  formatPercent: {
    id: 'formatPercent',
    label: 'Percent',
    accelerator: { key: '5', ctrl: true, shift: true, shifted: '%' }
  },
  formatScientific: {
    id: 'formatScientific',
    label: 'Scientific',
    accelerator: { key: '6', ctrl: true, shift: true, shifted: '^' }
  },
  formatDate: {
    id: 'formatDate',
    label: 'Date',
    accelerator: { key: '3', ctrl: true, shift: true, shifted: '#' }
  },
  formatTime: {
    id: 'formatTime',
    label: 'Time',
    accelerator: { key: '2', ctrl: true, shift: true, shifted: '@' }
  },
  formatText: {
    id: 'formatText',
    label: 'Plain text',
    accelerator: { key: '7', ctrl: true, shift: true, shifted: '&' }
  },
  moreDecimals: { id: 'moreDecimals', label: 'More decimal places', accelerator: { key: ']', ctrl: true } },
  fewerDecimals: { id: 'fewerDecimals', label: 'Fewer decimal places', accelerator: { key: '[', ctrl: true } },
  clearFormat: { id: 'clearFormat', label: 'Clear formatting', accelerator: { key: '\\', ctrl: true } }
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
 * Five menus and not the six a spreadsheet ends up with: File
 * arrives with Phase 16, which is when there is a file to open. A
 * menu of disabled items is a worse answer than no menu — it
 * advertises, and then it refuses.
 *
 * Format's mnemonic is `o` rather than `f`, because File is coming
 * and will want `f` — and a mnemonic that moves once people have
 * learned it is worse than one that was never the obvious letter.
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
    id: 'insert',
    label: 'Insert',
    mnemonic: 'i',
    entries: [
      'insertRowAbove',
      'insertRowBelow',
      SEPARATOR,
      'insertColumnLeft',
      'insertColumnRight',
      SEPARATOR,
      'deleteRows',
      'deleteColumns'
    ]
  },
  {
    id: 'format',
    label: 'Format',
    mnemonic: 'o',
    entries: [
      'bold',
      'italic',
      'underline',
      SEPARATOR,
      'alignLeft',
      'alignCenter',
      'alignRight',
      'wrap',
      SEPARATOR,
      'formatGeneral',
      'formatNumber',
      'formatCurrency',
      'formatPercent',
      'formatScientific',
      'formatDate',
      'formatTime',
      'formatText',
      SEPARATOR,
      'moreDecimals',
      'fewerDecimals',
      SEPARATOR,
      'borderAll',
      'borderOutline',
      'borderTop',
      'borderBottom',
      'borderThickBottom',
      'borderNone',
      SEPARATOR,
      'clearFormat'
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
  if (accelerator.alt === true) {
    parts.push('Alt');
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
  const alt = modifiers.alt === true;
  for (const command of Object.values(COMMANDS)) {
    const wanted = command.accelerator;
    if (wanted === undefined || command.viaKeyTable === true) {
      continue;
    }
    if (
      (matchesKey(wanted.key, key) || (wanted.shifted !== undefined && wanted.shifted === key)) &&
      (wanted.ctrl === true) === accel &&
      (wanted.shift === true) === shift &&
      (wanted.alt === true) === alt
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
