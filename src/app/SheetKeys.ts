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

export interface KeyModifiers {
  readonly shift?: boolean;
  readonly control?: boolean;
  readonly meta?: boolean;
  readonly alt?: boolean;
}

export type SheetAction =
  /** Move the selection by a delta, clamped by the caller. */
  | { readonly kind: 'move'; readonly rows: number; readonly columns: number }
  /** Move to an edge of the sheet. */
  | { readonly kind: 'jump'; readonly to: 'rowStart' | 'sheetStart' | 'rowEnd' | 'sheetEnd' }
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
  | { readonly kind: 'redo' };

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
  const accel = modifiers.control === true || modifiers.meta === true;

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
      case 'home':
        return { kind: 'jump', to: 'sheetStart' };
      case 'end':
        return { kind: 'jump', to: 'sheetEnd' };
      default:
        return null;
    }
  }

  switch (key) {
    case 'ArrowUp':
      return { kind: 'move', rows: -1, columns: 0 };
    case 'ArrowDown':
      return { kind: 'move', rows: 1, columns: 0 };
    case 'ArrowLeft':
      return { kind: 'move', rows: 0, columns: -1 };
    case 'ArrowRight':
      return { kind: 'move', rows: 0, columns: 1 };
    case 'PageUp':
      return { kind: 'move', rows: -PAGE_ROWS, columns: 0 };
    case 'PageDown':
      return { kind: 'move', rows: PAGE_ROWS, columns: 0 };
    case 'Home':
      return { kind: 'jump', to: 'rowStart' };
    case 'End':
      return { kind: 'jump', to: 'rowEnd' };
    // Enter and Tab move without opening anything, which is what they
    // do on a cell nobody is typing into.
    case 'Enter':
      return { kind: 'move', rows: shift ? -1 : 1, columns: 0 };
    case 'Tab':
      return { kind: 'move', rows: 0, columns: shift ? -1 : 1 };
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
