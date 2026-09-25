import type { UiKeyModifiers } from 'gesso-core';

/**
 * What a key means, as a pure function.
 *
 * The commit semantics of a spreadsheet are muscle memory: Enter goes
 * down, Tab goes right, Esc puts back what was there, F2 opens the cell
 * you are on. Getting one of them wrong is not a bug someone reports,
 * it is a tool that feels broken. So they are a table rather than a
 * chain of `if`s inside an event handler, and the table has its own
 * spec that no rendering is needed to run.
 *
 * The same key means different things while a cell is open, which is
 * why `editing` is a parameter rather than two functions: Enter
 * *commits and then* moves, and writing that as two tables invites the
 * two to disagree.
 */

/**
 * The engine's own modifier names, not a set of our own.
 *
 * It was a set of our own, with `control` where the engine says
 * `ctrl`, and the mistake survived every spec: `fireEvent.press` takes
 * the engine's shape, the specs passed theirs through a
 * `Record<string, boolean>`, and an index signature is assignable to
 * anything — so both sides agreed on a field the browser never sets.
 * Not one accelerator worked in a real browser: ctrl+C typed a `c`.
 */
export type KeyModifiers = Partial<UiKeyModifiers>;

export type SheetAction =
  /**
   * Move the selection by a delta, clamped by the caller.
   *
   * `extend` keeps the anchor where it is and moves the other corner,
   * which is what shift does to every selection everywhere.
   */
  | { readonly kind: 'move'; readonly rows: number; readonly columns: number; readonly extend: boolean }
  /** Move to an edge of the sheet. */
  | { readonly kind: 'jump'; readonly to: 'rowStart' | 'sheetStart' | 'rowEnd' | 'sheetEnd'; readonly extend: boolean }
  /** Open the cell with what is already in it. */
  | { readonly kind: 'edit' }
  /** Open the cell, replacing its contents with this text. */
  | { readonly kind: 'replace'; readonly text: string }
  /** Commit what is being typed, then move. */
  | { readonly kind: 'commit'; readonly rows: number; readonly columns: number }
  /** Close the cell and put back what was there. */
  | { readonly kind: 'cancel' }
  /** Empty the selected cell without opening it. */
  | { readonly kind: 'clear' }
  | { readonly kind: 'undo' }
  | { readonly kind: 'redo' }
  /** Put the selection on the clipboard, and empty it when cutting. */
  | { readonly kind: 'copy'; readonly cut: boolean }
  /** Select every cell in the sheet. */
  | { readonly kind: 'selectAll' };

/**
 * A page, in rows. Not the viewport's height: a page that moved by
 * exactly a screenful would leave nothing in common between the old
 * view and the new one, and a person paging through a sheet keeps a
 * row or two to navigate by. Every spreadsheet does this.
 */
export const PAGE_ROWS = 24;

/**
 * The action a key press means, or null to let it through.
 *
 * Null is the important return while a cell is open: every key that is
 * not a commit or a cancel belongs to the text, and a grid that
 * swallowed them would be a grid you cannot type an arrow into.
 */
export function keyAction(key: string, modifiers: KeyModifiers, editing: boolean): SheetAction | null {
  const shift = modifiers.shift === true;
  const accel = modifiers.ctrl === true || modifiers.meta === true;

  if (editing) {
    switch (key) {
      case 'Enter':
        return { kind: 'commit', rows: shift ? -1 : 1, columns: 0 };
      case 'Tab':
        return { kind: 'commit', rows: 0, columns: shift ? -1 : 1 };
      case 'Escape':
        return { kind: 'cancel' };
      default:
        // The caret, the selection, the clipboard and IME all live in
        // the text. Nothing else here is the grid's.
        return null;
    }
  }

  if (accel) {
    switch (key.toLowerCase()) {
      case 'z':
        return shift ? { kind: 'redo' } : { kind: 'undo' };
      case 'y':
        return { kind: 'redo' };
      case 'c':
        return { kind: 'copy', cut: false };
      case 'x':
        return { kind: 'copy', cut: true };
      case 'a':
        return { kind: 'selectAll' };
      // Paste is not here. The text arrives from the system a moment
      // later, as a Paste event, and a key handler has nothing to
      // paste at the moment the key goes down.
      case 'v':
        return null;
      case 'home':
        return { kind: 'jump', to: 'sheetStart', extend: shift };
      case 'end':
        return { kind: 'jump', to: 'sheetEnd', extend: shift };
      default:
        return null;
    }
  }

  switch (key) {
    case 'ArrowUp':
      return { kind: 'move', rows: -1, columns: 0, extend: shift };
    case 'ArrowDown':
      return { kind: 'move', rows: 1, columns: 0, extend: shift };
    case 'ArrowLeft':
      return { kind: 'move', rows: 0, columns: -1, extend: shift };
    case 'ArrowRight':
      return { kind: 'move', rows: 0, columns: 1, extend: shift };
    case 'PageUp':
      return { kind: 'move', rows: -PAGE_ROWS, columns: 0, extend: shift };
    case 'PageDown':
      return { kind: 'move', rows: PAGE_ROWS, columns: 0, extend: shift };
    case 'Home':
      return { kind: 'jump', to: 'rowStart', extend: shift };
    case 'End':
      return { kind: 'jump', to: 'rowEnd', extend: shift };
    // Enter and Tab move without opening anything, which is what they
    // do on a cell nobody is typing into. Shift reverses them rather
    // than extending: that is what it means on these two keys.
    case 'Enter':
      return { kind: 'move', rows: shift ? -1 : 1, columns: 0, extend: false };
    case 'Tab':
      return { kind: 'move', rows: 0, columns: shift ? -1 : 1, extend: false };
    case 'F2':
      return { kind: 'edit' };
    case 'Delete':
    case 'Backspace':
      return { kind: 'clear' };
    case 'Escape':
      return null;
    default:
      // Typing over a selected cell replaces it, and the character
      // typed is the first one of the new value rather than being
      // swallowed by the transition into edit mode.
      return isPrintable(key) ? { kind: 'replace', text: key } : null;
  }
}

/**
 * The navigation keys, as the shortcut sheet prints them.
 *
 * Beside the table that answers them rather than in the help screen
 * that shows them, so that a key which stops working is a key the
 * spec below notices. `SheetKeys.spec.ts` presses every one of these
 * and fails the build if any of them has become a key that does
 * nothing — which is the failure a hand-written help page hides.
 *
 * Only the keys the sheet itself answers. The accelerators live in
 * `SheetCommands`, where the menus can read their labels too.
 */
export interface NavigationKey {
  /** As the shortcut sheet prints it. */
  readonly keys: string;
  readonly label: string;
  /** One key from `keys`, for the spec that presses them all. */
  readonly probe: string;
  /** True when the key only means something with a cell open. */
  readonly whileEditing?: true;
}

export const NAVIGATION: readonly NavigationKey[] = [
  { keys: 'Arrows', label: 'Move one cell', probe: 'ArrowDown' },
  { keys: 'Shift+Arrows', label: 'Extend the selection', probe: 'ArrowDown' },
  { keys: 'Enter', label: 'Move down; commit and move down', probe: 'Enter' },
  { keys: 'Tab', label: 'Move right; commit and move right', probe: 'Tab' },
  { keys: 'Home / End', label: 'Start or end of the row', probe: 'Home' },
  { keys: 'Ctrl+Home / Ctrl+End', label: 'Start or end of the sheet', probe: 'Home' },
  { keys: 'PageUp / PageDown', label: 'Move a screen at a time', probe: 'PageDown' },
  { keys: 'F2', label: 'Edit the cell you are on', probe: 'F2' },
  { keys: 'Escape', label: 'Put back what was there', probe: 'Escape', whileEditing: true }
];

/**
 * Whether a key event carries a character rather than a command.
 *
 * A `key` of one code point is the rule browsers already follow:
 * `'a'`, `'7'` and `'£'` are one, `'Shift'` and `'ArrowUp'` are not.
 * Counting code points rather than UTF-16 units matters for the keys
 * that are one character and two units — an emoji on a touch keyboard.
 */
export function isPrintable(key: string): boolean {
  return [...key].length === 1 && key !== '\t' && key !== '\n';
}
