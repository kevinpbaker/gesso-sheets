import { describe, expect, it } from 'vitest';

import { CLOSED, menuBarStep, type MenuBarContext, type MenuBarState } from './MenuBarModel';
import { COMMANDS, MENUS, SEPARATOR, type CommandId, type MenuDefinition } from './SheetCommands';

/**
 * The menu bar, as a table.
 *
 * No rendering: the point of pulling the traversal out of the
 * component is that it can be asserted here, where a wrong answer is
 * one line rather than a screenshot. What the component does with
 * these answers is Phase 8's other spec's business.
 */

const everything: MenuBarContext = {
  enabled: () => true,
  labelOf: id => COMMANDS[id].label
};

function withNothingEnabled(...disabled: CommandId[]): MenuBarContext {
  return { ...everything, enabled: id => !disabled.includes(id) };
}

/** Press a run of keys, from a starting state. */
function press(from: MenuBarState, keys: readonly string[], context = everything, menus = MENUS) {
  let state = from;
  let choose: CommandId | undefined;
  let dismissed = false;
  for (const key of keys) {
    const step = menuBarStep(state, key, menus, context);
    if (step === null) {
      continue;
    }
    state = step.state;
    choose = step.choose ?? choose;
    dismissed = step.dismiss === true || dismissed;
  }
  return { state, choose, dismissed };
}

/** The command the highlight is on, or null when it is on nothing. */
function highlighted(state: MenuBarState, menus: readonly MenuDefinition[] = MENUS): CommandId | null {
  const entry = menus[state.focused].entries[state.active];
  return entry === undefined || entry === SEPARATOR ? null : entry;
}

describe('the bar with nothing open', () => {
  it('walks the menus with the arrows, and wraps', () => {
    expect(press(CLOSED, ['ArrowRight']).state.focused).toBe(1);
    expect(press(CLOSED, ['ArrowRight', 'ArrowRight']).state.focused).toBe(2);
    // Off the end and round to the start, which is what a bar does.
    expect(press(CLOSED, ['ArrowRight', 'ArrowRight', 'ArrowRight']).state.focused).toBe(0);
    expect(press(CLOSED, ['ArrowLeft']).state.focused).toBe(MENUS.length - 1);
  });

  it('jumps to the ends with Home and End', () => {
    expect(press(CLOSED, ['End']).state.focused).toBe(MENUS.length - 1);
    expect(press(CLOSED, ['End', 'Home']).state.focused).toBe(0);
  });

  it('opens downward on Enter, Space and ArrowDown', () => {
    for (const key of ['Enter', ' ', 'ArrowDown']) {
      const { state } = press(CLOSED, [key]);
      expect(state.open, key).toBe(true);
      expect(highlighted(state), key).toBe('undo');
    }
  });

  /** ArrowUp opens onto the *last* item, which is what it means. */
  it('opens upward on ArrowUp', () => {
    const { state } = press(CLOSED, ['ArrowUp']);
    expect(state.open).toBe(true);
    expect(highlighted(state)).toBe('gotoCell');
  });

  it('opens a menu by its letter', () => {
    const { state } = press(CLOSED, ['d']);
    expect(state.focused).toBe(1);
    expect(highlighted(state)).toBe('fillDown');
  });

  it('gives the keyboard back on Escape', () => {
    const { state, dismissed } = press(CLOSED, ['Escape']);
    expect(dismissed).toBe(true);
    expect(state.open).toBe(false);
  });

  /**
   * A key the bar does not claim has to reach what is underneath. A
   * bar that answered everything while focused is a bar nobody can
   * Tab out of.
   */
  it('lets a key it does not claim through', () => {
    expect(menuBarStep(CLOSED, 'Tab', MENUS, everything)).toBeNull();
    expect(menuBarStep(CLOSED, 'q', MENUS, everything)).toBeNull();
  });
});

describe('a menu that is open', () => {
  const open = press(CLOSED, ['ArrowDown']).state;

  it('walks its items, skipping the rules between them', () => {
    // Edit opens on Undo; Redo is next; the third press has to clear
    // the separator and land on Cut rather than on the rule.
    expect(highlighted(press(open, ['ArrowDown']).state)).toBe('redo');
    expect(highlighted(press(open, ['ArrowDown', 'ArrowDown']).state)).toBe('cut');
  });

  it('wraps from the last item to the first', () => {
    const atEnd = press(open, ['End']).state;
    expect(highlighted(atEnd)).toBe('gotoCell');
    expect(highlighted(press(atEnd, ['ArrowDown']).state)).toBe('undo');
  });

  it('skips a command that cannot be chosen', () => {
    const context = withNothingEnabled('redo');
    const state = press(CLOSED, ['ArrowDown'], context).state;
    expect(highlighted(press(state, ['ArrowDown'], context).state)).toBe('cut');
  });

  it('opens onto the first command that can be chosen', () => {
    const context = withNothingEnabled('undo', 'redo');
    expect(highlighted(press(CLOSED, ['ArrowDown'], context).state)).toBe('cut');
  });

  /**
   * The case a focus-trapped popup cannot serve, and the reason this
   * model exists rather than `Menu`: left and right walk the *bar*
   * while a menu is showing.
   */
  it('moves to the menu next door and stays open', () => {
    const next = press(open, ['ArrowRight']).state;
    expect(next.focused).toBe(1);
    expect(next.open).toBe(true);
    expect(highlighted(next)).toBe('fillDown');

    const back = press(next, ['ArrowLeft']).state;
    expect(back.focused).toBe(0);
    expect(back.open).toBe(true);
  });

  it('chooses on Enter and closes', () => {
    const { state, choose, dismissed } = press(open, ['ArrowDown', 'Enter']);
    expect(choose).toBe('redo');
    expect(state.open).toBe(false);
    expect(dismissed).toBe(true);
  });

  it('refuses to choose a command that is not enabled', () => {
    const context = withNothingEnabled('undo');
    // Force the highlight onto the disabled item the only way the
    // model allows — it never lands there on its own, which is the
    // point, but a stale state must not choose it either.
    const stuck: MenuBarState = { focused: 0, open: true, active: 0 };
    const step = menuBarStep(stuck, 'Enter', MENUS, context);
    expect(step?.choose).toBeUndefined();
    expect(step?.state.open).toBe(true);
  });

  /**
   * Two Escapes rather than one. The first closes the menu and leaves
   * the bar focused; only the second hands the keyboard back. Closing
   * straight through to the sheet loses the place of anybody who
   * opened the wrong menu, which is most people most of the time.
   */
  it('closes to the bar on the first Escape and to the sheet on the second', () => {
    const first = menuBarStep(open, 'Escape', MENUS, everything);
    expect(first?.state.open).toBe(false);
    expect(first?.state.focused).toBe(0);
    expect(first?.dismiss).toBeUndefined();

    const second = menuBarStep(first!.state, 'Escape', MENUS, everything);
    expect(second?.dismiss).toBe(true);
  });

  it('closes rather than tabbing inside itself', () => {
    const step = menuBarStep(open, 'Tab', MENUS, everything);
    expect(step?.state.open).toBe(false);
    expect(step?.dismiss).toBe(true);
  });

  it('jumps to an item by its first letter', () => {
    expect(highlighted(press(open, ['p']).state)).toBe('paste');
    expect(highlighted(press(open, ['s']).state)).toBe('selectAll');
  });

  /** The same letter twice walks the items sharing it. */
  it('walks the items that share a letter', () => {
    const menus: readonly MenuDefinition[] = [
      { id: 'x', label: 'X', mnemonic: 'x', entries: ['cut', 'copy', 'clear'] }
    ];
    const first = press(CLOSED, ['ArrowDown', 'c'], everything, menus).state;
    expect(highlighted(first, menus)).toBe('copy');
    const second = press(first, ['c'], everything, menus).state;
    expect(highlighted(second, menus)).toBe('clear');
    // And round again, rather than stopping at the bottom.
    expect(highlighted(press(second, ['c'], everything, menus).state, menus)).toBe('cut');
  });

  it('leaves the highlight alone when no item starts with the letter', () => {
    const before = press(open, ['ArrowDown']).state;
    expect(highlighted(press(before, ['q']).state)).toBe('redo');
  });
});
