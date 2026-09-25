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
  | 'sheetTabs'
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
  | 'formatDateTime'
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
  | 'defineName'
  | 'borderAll'
  | 'borderOutline'
  | 'borderTop'
  | 'borderBottom'
  | 'borderThickBottom'
  | 'borderNone'
  | 'mergeCells'
  | 'unmergeCells'
  | 'sortAscending'
  | 'sortDescending'
  | 'hideColumns'
  | 'showColumns'
  | 'hideRows'
  | 'showRows'
  | 'autofitColumns'
  | 'filterToSelection'
  | 'clearFilter'
  | 'freezeHere'
  | 'freezeTopRow'
  | 'freezeFirstColumn'
  | 'unfreeze'
  | 'insertSheet'
  | 'renameSheet'
  | 'duplicateSheet'
  | 'deleteSheet'
  | 'moveSheetLeft'
  | 'moveSheetRight'
  | 'nextSheet'
  | 'previousSheet'
  | 'sheetColourNone'
  | 'sheetColourBlue'
  | 'sheetColourRed'
  | 'sheetColourGreen'
  | 'sheetColourPurple'
  | 'sheetColourOrange';

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
   * The door to the tab strip, and it needs one.
   *
   * **Tab does not get there**, and cannot: Tab moves the selection
   * one cell right, which is what it does in every spreadsheet, so
   * the grid consumes it and a keyboard standing in the sheet can
   * never leave by that route. Found in a browser and not by the
   * specs, which drive focus rather than pressing the key — so the
   * strip was a region that read perfectly in the accessibility tree
   * and nobody using a keyboard could reach.
   *
   * Alt+F10 rather than a letter, because it is the same kind of
   * thing F10 is and reads as the same gesture: F10 for the menus
   * along the top, Alt+F10 for the tabs along the bottom. Hidden from
   * the menus for the reason `menuBar` is — an item that takes you
   * somewhere cannot usefully live in the place it takes you from —
   * and still advertised by the shortcut sheet, which reads this
   * table.
   */
  sheetTabs: { id: 'sheetTabs', label: 'Go to the sheet tabs', accelerator: { key: 'F10', alt: true }, hidden: true },

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
   * The menu route to a gesture that already existed.
   *
   * Naming a range is done in the name box — select, type what it is,
   * press Enter — and a gesture nobody can find is a feature nobody
   * has. This item is the route: it puts the keyboard in the box and
   * says, under it, which range is about to be named. The work is
   * still the box's, which is why this is a menu entry and not a
   * dialog.
   *
   * No accelerator. Excel's is Ctrl+F3, F3 is the browser's find, and
   * a shortcut that opens the browser's own search bar instead is
   * worse than no shortcut at all.
   */
  defineName: { id: 'defineName', label: 'Name the selection…' },

  /**
   * The sheets, as commands as well as tabs.
   *
   * The strip along the bottom is the surface a pointer wants; this
   * is the one a keyboard wants, and it is the complete one. Every
   * tab operation is here, because Phase 4's standard is that
   * anything the mouse can do the keyboard can, and a colour chosen
   * from a swatch nobody can tab to is a colour only half the people
   * using this can set.
   */
  insertSheet: { id: 'insertSheet', label: 'Insert sheet' },
  renameSheet: { id: 'renameSheet', label: 'Rename sheet…', accelerator: { key: 'F2', shift: true } },
  duplicateSheet: { id: 'duplicateSheet', label: 'Duplicate sheet' },
  deleteSheet: { id: 'deleteSheet', label: 'Delete sheet…' },
  moveSheetLeft: { id: 'moveSheetLeft', label: 'Move sheet left' },
  moveSheetRight: { id: 'moveSheetRight', label: 'Move sheet right' },
  /**
   * Alt, and not the Ctrl+PageUp and Ctrl+PageDown every desktop
   * spreadsheet uses.
   *
   * Those switch *browser* tabs, and a page cannot prevent it. A
   * shortcut that took somebody out of the spreadsheet altogether
   * would be worse than no shortcut, which is the same reasoning the
   * insert-row accelerators are on Alt for.
   */
  nextSheet: { id: 'nextSheet', label: 'Next sheet', accelerator: { key: 'PageDown', alt: true } },
  previousSheet: { id: 'previousSheet', label: 'Previous sheet', accelerator: { key: 'PageUp', alt: true } },
  sheetColourNone: { id: 'sheetColourNone', label: 'Tab colour: none' },
  sheetColourBlue: { id: 'sheetColourBlue', label: 'Tab colour: blue' },
  sheetColourRed: { id: 'sheetColourRed', label: 'Tab colour: red' },
  sheetColourGreen: { id: 'sheetColourGreen', label: 'Tab colour: green' },
  sheetColourPurple: { id: 'sheetColourPurple', label: 'Tab colour: purple' },
  sheetColourOrange: { id: 'sheetColourOrange', label: 'Tab colour: orange' },

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
  mergeCells: { id: 'mergeCells', label: 'Merge cells', accelerator: { key: 'm', ctrl: true, alt: true } },
  unmergeCells: { id: 'unmergeCells', label: 'Unmerge', accelerator: { key: 'm', ctrl: true, alt: true, shift: true } },

  sortAscending: { id: 'sortAscending', label: 'Sort A to Z' },
  sortDescending: { id: 'sortDescending', label: 'Sort Z to A' },
  hideColumns: { id: 'hideColumns', label: 'Hide columns', accelerator: { key: '0', ctrl: true, alt: true } },
  showColumns: {
    id: 'showColumns',
    label: 'Show columns',
    accelerator: { key: '0', ctrl: true, alt: true, shift: true, shifted: ')' }
  },
  freezeHere: { id: 'freezeHere', label: 'Freeze up to here' },
  freezeTopRow: { id: 'freezeTopRow', label: 'Freeze the top row' },
  freezeFirstColumn: { id: 'freezeFirstColumn', label: 'Freeze the first column' },
  unfreeze: { id: 'unfreeze', label: 'Unfreeze' },
  autofitColumns: { id: 'autofitColumns', label: 'Fit columns to contents' },
  filterToSelection: {
    id: 'filterToSelection',
    label: 'Keep only rows like this one',
    accelerator: { key: 'k', ctrl: true, alt: true }
  },
  clearFilter: { id: 'clearFilter', label: 'Show every row', accelerator: { key: 'k', ctrl: true, alt: true, shift: true } },
  hideRows: { id: 'hideRows', label: 'Hide rows', accelerator: { key: '9', ctrl: true, alt: true } },
  showRows: {
    id: 'showRows',
    label: 'Show rows',
    accelerator: { key: '9', ctrl: true, alt: true, shift: true, shifted: '(' }
  },

  bold: { id: 'bold', label: 'Bold', accelerator: { key: 'b', ctrl: true } },
  italic: { id: 'italic', label: 'Italic', accelerator: { key: 'i', ctrl: true } },
  underline: { id: 'underline', label: 'Underline', accelerator: { key: 'u', ctrl: true } },
  alignLeft: { id: 'alignLeft', label: 'Align left', accelerator: { key: 'l', ctrl: true, shift: true } },
  alignCenter: { id: 'alignCenter', label: 'Align centre', accelerator: { key: 'e', ctrl: true, shift: true } },
  alignRight: { id: 'alignRight', label: 'Align right', accelerator: { key: 'r', ctrl: true, shift: true } },
  /**
   * Set by nothing, for now.
   *
   * The format travels and the file keeps it, and `Grid` binds it —
   * but `LazySheet` takes one row height for every row, so a wrapped
   * cell has nowhere to put its second line. It is `hidden` rather
   * than deleted because the model is right and only the engine is
   * missing: when row heights vary, it goes back in the menu and
   * nothing else has to change. A control that silently does nothing
   * is worse than one that is not there.
   */
  wrap: { id: 'wrap', label: 'Wrap text', accelerator: { key: 'w', ctrl: true, shift: true }, hidden: true },
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
  /**
   * No accelerator, and that is the decision rather than an omission.
   *
   * Date and Time have had ctrl-shift-3 and ctrl-shift-2 since they
   * were spreadsheet keys, and there is no third one anybody knows.
   * Inventing a chord for the rarer format would put something
   * unguessable on a key somebody is used to reaching past.
   */
  formatDateTime: {
    id: 'formatDateTime',
    label: 'Date and time'
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
 * Six menus, and the last one missing is File — which arrives with
 * Phase 16, when there is a file to open. A menu of disabled items is
 * a worse answer than no menu: it advertises, and then it refuses.
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
      'deleteColumns',
      SEPARATOR,
      'defineName'
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
      SEPARATOR,
      'formatGeneral',
      'formatNumber',
      'formatCurrency',
      'formatPercent',
      'formatScientific',
      'formatDate',
      'formatTime',
      'formatDateTime',
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
      'mergeCells',
      'unmergeCells',
      SEPARATOR,
      'clearFormat'
    ]
  },
  {
    id: 'sheet',
    label: 'Sheet',
    mnemonic: 's',
    entries: [
      'insertSheet',
      'duplicateSheet',
      'renameSheet',
      'deleteSheet',
      SEPARATOR,
      'moveSheetLeft',
      'moveSheetRight',
      SEPARATOR,
      'previousSheet',
      'nextSheet',
      SEPARATOR,
      'sheetColourNone',
      'sheetColourBlue',
      'sheetColourRed',
      'sheetColourGreen',
      'sheetColourPurple',
      'sheetColourOrange'
    ]
  },
  {
    id: 'view',
    label: 'View',
    mnemonic: 'v',
    entries: ['freezeHere', 'freezeTopRow', 'freezeFirstColumn', SEPARATOR, 'unfreeze']
  },
  {
    id: 'data',
    label: 'Data',
    mnemonic: 'd',
    entries: [
      'sortAscending',
      'sortDescending',
      SEPARATOR,
      'filterToSelection',
      'clearFilter',
      SEPARATOR,
      'fillDown',
      'fillRight',
      SEPARATOR,
      'autofitColumns',
      SEPARATOR,
      'hideRows',
      'showRows',
      'hideColumns',
      'showColumns',
      SEPARATOR,
      'recalculate'
    ]
  },
  {
    id: 'help',
    label: 'Help',
    mnemonic: 'h',
    entries: ['shortcuts']
  }
];

/**
 * The commands that belong to the proof route and to no other.
 *
 * `/` is a spreadsheet. `/proof` is the same spreadsheet with the
 * evidence attached — the strip along the top, and the one command
 * whose entire purpose is to give the strip something to measure.
 * Nobody opening a spreadsheet wants to recalculate two hundred
 * thousand cells, and an item that exists to be photographed belongs
 * on the page that photographs it.
 *
 * A list here rather than a flag on the row, because the tables above
 * say what a command *is* and this says where it is offered. The
 * command itself is unchanged on both routes: `COMMANDS` still holds
 * it, `commandFor` still answers F9 with it, and the specs that keep
 * the two tables honest still see every row.
 */
export const PROOF_ONLY: readonly CommandId[] = ['recalculate'];

/** Whether a route offers a command at all. */
export function offers(id: CommandId, proof: boolean): boolean {
  return proof || !PROOF_ONLY.includes(id);
}

/**
 * The menus as a route shows them.
 *
 * Dropping an entry is the easy half; the rules it leaves behind are
 * the half that shows. Data ends `…, SEPARATOR, 'recalculate'`, so a
 * plain filter gives a menu that ends in a line drawn under nothing —
 * which is exactly what `SheetCommands.spec` refuses to allow in the
 * table itself, and there is no reason to allow it in the table's
 * output either. So the rules are collapsed: none at the start, none
 * at the end, never two together.
 */
export function menusFor(proof: boolean): readonly MenuDefinition[] {
  if (proof) {
    return MENUS;
  }
  return MENUS.map(menu => ({
    ...menu,
    entries: withoutStrayRules(menu.entries.filter(entry => entry === SEPARATOR || offers(entry, proof)))
  })).filter(menu => menu.entries.length > 0);
}

function withoutStrayRules(entries: readonly MenuEntry[]): readonly MenuEntry[] {
  const kept: MenuEntry[] = [];
  for (const entry of entries) {
    if (entry === SEPARATOR && (kept.length === 0 || kept[kept.length - 1] === SEPARATOR)) {
      continue;
    }
    kept.push(entry);
  }
  while (kept.length > 0 && kept[kept.length - 1] === SEPARATOR) {
    kept.pop();
  }
  return kept;
}

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
