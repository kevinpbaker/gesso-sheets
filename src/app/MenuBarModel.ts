/**
 * How a menu bar behaves under the keyboard, as a pure function.
 *
 * This is `SheetKeys`' argument applied to the chrome. A menu bar is
 * more of a state machine than it looks — the arrows mean four
 * different things depending on whether a menu is open, Escape means
 * two, and a letter means "jump to the item starting with it" or
 * "open the menu starting with it" — and written inside event
 * handlers that is a pile of `if`s nobody can check. Written here it
 * is a table with a spec, and the component below it does nothing but
 * draw the answer.
 *
 * **Why it is not `Menu` from `gesso-components`.** That component
 * owns its own keyboard and traps focus inside itself, which is
 * correct for a popup opened by a button and wrong for a bar: with
 * focus trapped in the popup, ArrowLeft cannot reach the bar to move
 * to the menu next door, and moving between menus with the arrows is
 * most of what makes a menu bar a menu bar. Its `MenuItem` also has
 * no separator, no accelerator column and no notion of a mnemonic.
 * The shape wanted here is now known, which is the point at which
 * the roadmap said the upstreaming decision could be taken.
 */

import { SEPARATOR, type CommandId, type MenuDefinition } from './SheetCommands';
import { isPrintable } from './SheetKeys';

export interface MenuBarState {
  /**
   * The menu the bar's focus is on.
   *
   * There is always one, open or not: a menu bar is a single tab stop
   * with a roving highlight inside it, not seven tab stops, so that
   * Tab past the bar takes one press rather than one per menu.
   */
  readonly focused: number;
  readonly open: boolean;
  /**
   * The highlighted entry in the open menu, as an index into
   * `entries` — separators included.
   *
   * Into the whole list rather than into the choosable subset,
   * for the reason `Menu` gives for the same decision: the row paints
   * its highlight by comparing its own position, and walking one list
   * while painting by another makes them disagree the moment a rule
   * or a disabled item sits anywhere but the end.
   */
  readonly active: number;
}

export const CLOSED: MenuBarState = { focused: 0, open: false, active: -1 };

export interface MenuBarStep {
  readonly state: MenuBarState;
  /** The command the key chose, when it chose one. */
  readonly choose?: CommandId;
  /** True when the key means "give the keyboard back to the sheet". */
  readonly dismiss?: boolean;
}

/**
 * What the model needs to know about the commands it is walking.
 *
 * Passed in rather than imported so the model is testable against
 * menus of its own, and so nothing here depends on the particular
 * three menus this application happens to ship.
 */
export interface MenuBarContext {
  /** Whether a command can be chosen right now. */
  readonly enabled: (id: CommandId) => boolean;
  /** A command's label, which is what type-ahead matches against. */
  readonly labelOf: (id: CommandId) => string;
}

/**
 * What a key does to the bar, or null to let it through.
 *
 * Null is as important here as it is in `keyAction`: a key the bar
 * does not claim has to reach whatever is underneath, and a bar that
 * swallowed everything while it had focus would be a bar you cannot
 * Tab out of.
 */
export function menuBarStep(
  state: MenuBarState,
  key: string,
  menus: readonly MenuDefinition[],
  context: MenuBarContext
): MenuBarStep | null {
  if (menus.length === 0) {
    return null;
  }
  const focused = clampMenu(state.focused, menus);

  if (!state.open) {
    switch (key) {
      case 'ArrowLeft':
        return { state: { focused: wrap(focused - 1, menus.length), open: false, active: -1 } };
      case 'ArrowRight':
        return { state: { focused: wrap(focused + 1, menus.length), open: false, active: -1 } };
      case 'Home':
        return { state: { focused: 0, open: false, active: -1 } };
      case 'End':
        return { state: { focused: menus.length - 1, open: false, active: -1 } };
      case 'ArrowDown':
      case 'Enter':
      case ' ':
        return { state: opened(focused, menus, context, 1) };
      case 'ArrowUp':
        return { state: opened(focused, menus, context, -1) };
      case 'Escape':
        return { state: CLOSED, dismiss: true };
      default: {
        // A bare letter opens the menu it names. With the bar focused
        // and nothing open there is nothing else a letter could mean,
        // and it is how every menu bar has worked for thirty years.
        const named = isPrintable(key)
          ? menus.findIndex(menu => menu.mnemonic === key.toLowerCase())
          : -1;
        return named === -1 ? null : { state: opened(named, menus, context, 1) };
      }
    }
  }

  const entries = menus[focused].entries;
  switch (key) {
    case 'ArrowDown':
      return { state: { focused, open: true, active: seek(entries, state.active + 1, 1, context) } };
    case 'ArrowUp':
      return { state: { focused, open: true, active: seek(entries, state.active - 1, -1, context) } };
    case 'Home':
      return { state: { focused, open: true, active: seek(entries, 0, 1, context) } };
    case 'End':
      return { state: { focused, open: true, active: seek(entries, entries.length - 1, -1, context) } };
    /**
     * Left and right walk the *bar* while a menu is open, closing
     * this one and opening the next. This is the case a focus-trapped
     * popup cannot serve, and the reason this model exists.
     */
    case 'ArrowLeft':
      return { state: opened(wrap(focused - 1, menus.length), menus, context, 1) };
    case 'ArrowRight':
      return { state: opened(wrap(focused + 1, menus.length), menus, context, 1) };
    case 'Escape':
      // The first Escape closes the menu and leaves the bar focused,
      // the second gives the keyboard back to the sheet. Closing
      // straight to the sheet loses the place of anybody who opened
      // the wrong menu, which is most people, most of the time.
      return { state: { focused, open: false, active: -1 } };
    case 'Enter':
    case ' ': {
      const entry = entries[state.active];
      if (entry === undefined || entry === SEPARATOR || !context.enabled(entry)) {
        return { state };
      }
      return { state: { focused, open: false, active: -1 }, choose: entry, dismiss: true };
    }
    case 'Tab':
      // Tab out of an open menu closes it rather than moving inside
      // it. A menu is not a set of tab stops.
      return { state: CLOSED, dismiss: true };
    default: {
      if (!isPrintable(key)) {
        return null;
      }
      const found = typeAhead(entries, state.active, key, context);
      return found === -1 ? { state } : { state: { focused, open: true, active: found } };
    }
  }
}

/** Opening a menu with its first choosable entry highlighted. */
function opened(
  index: number,
  menus: readonly MenuDefinition[],
  context: MenuBarContext,
  direction: 1 | -1
): MenuBarState {
  const entries = menus[index].entries;
  const from = direction === 1 ? 0 : entries.length - 1;
  return { focused: index, open: true, active: seek(entries, from, direction, context) };
}

/**
 * The next entry that can be chosen, wrapping, or -1 when none can.
 *
 * Wrapping is what makes ArrowDown on the last item land on the
 * first, which every menu does and which people rely on to reach the
 * bottom of a long menu from the top.
 */
function seek(
  entries: readonly MenuEntryOf[],
  from: number,
  direction: 1 | -1,
  context: MenuBarContext
): number {
  const count = entries.length;
  if (count === 0) {
    return -1;
  }
  for (let moved = 0; moved < count; moved++) {
    const index = wrap(from + direction * moved, count);
    const entry = entries[index];
    if (entry !== SEPARATOR && context.enabled(entry)) {
      return index;
    }
  }
  return -1;
}

/**
 * The next entry whose label starts with a letter.
 *
 * Searched from *after* the highlight so that pressing the same
 * letter twice walks the items sharing it rather than sitting on the
 * first one.
 */
function typeAhead(entries: readonly MenuEntryOf[], active: number, key: string, context: MenuBarContext): number {
  const wanted = key.toLowerCase();
  for (let moved = 1; moved <= entries.length; moved++) {
    const index = wrap(active + moved, entries.length);
    const entry = entries[index];
    if (entry === SEPARATOR || !context.enabled(entry)) {
      continue;
    }
    if (context.labelOf(entry).toLowerCase().startsWith(wanted)) {
      return index;
    }
  }
  return -1;
}

type MenuEntryOf = CommandId | typeof SEPARATOR;

function wrap(value: number, count: number): number {
  return ((value % count) + count) % count;
}

function clampMenu(index: number, menus: readonly MenuDefinition[]): number {
  return Math.min(Math.max(index, 0), menus.length - 1);
}
