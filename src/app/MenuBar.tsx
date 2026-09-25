import { BehaviorSubject, combineLatest, map, type Observable } from 'rxjs';

import { clickOutside, Column, type UiKeyboardEvent, type UiNode } from 'gesso-core';
import { useOverlay } from 'gesso-components';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { CLOSED, menuBarStep, type MenuBarState } from './MenuBarModel';
import {
  acceleratorLabel,
  COMMANDS,
  MENUS,
  SEPARATOR,
  type CommandId,
  type MenuDefinition,
  type MenuEntry
} from './SheetCommands';

/**
 * The menu bar, drawn in the render worker like everything else.
 *
 * One tab stop, not seven: a menu bar is a single stop with a roving
 * highlight inside it, so Tab past it costs one press however many
 * menus it grows. The keys are answered by `menuBarStep`, which is a
 * table with its own spec; this file draws what the table says and
 * owns nothing but the overlay.
 *
 * **Focus stays on the bar while a menu is open.** The open menu is
 * an overlay that draws and does not take the keyboard, which is the
 * opposite of what a popup menu usually does and is the whole reason
 * this is not `Menu` from the component set: with focus trapped
 * inside the popup, ArrowLeft has nowhere to go, and walking to the
 * menu next door with the arrows is most of what makes a bar a bar.
 */
export interface MenuBarProps {
  /** Whether a command can be chosen at this moment. */
  readonly enabled: (id: CommandId) => boolean;
  readonly onChoose: (id: CommandId) => void;
  /** Called when the bar is done with the keyboard, so the sheet takes it. */
  readonly onDismiss: () => void;
  readonly menus?: readonly MenuDefinition[];
  /** Receives the bar itself, so F10 can put the keyboard on it. */
  readonly ref?: (node: UiNode | null) => void;
}

export function MenuBar(inputs: Inputs<MenuBarProps>, ctx: ComponentContext) {
  const menus = inputs.menus.value ?? MENUS;
  const overlay = useOverlay(ctx, 'sheet-menu');
  const state = new BehaviorSubject<MenuBarState>(CLOSED);
  /** The node each menu's title is drawn as, so the panel can sit under it. */
  const titles: (UiNode | null)[] = menus.map(() => null);
  /**
   * The title the pointer is over, or -1.
   *
   * Separate from the state above because hovering a title while
   * nothing is open is not the same as opening it: the title lights
   * up, and that is all. Once a menu *is* open, hovering a different
   * title switches to it, which is what every menu bar does and what
   * makes dragging along the bar work.
   */
  const hoveredTitle = new BehaviorSubject(-1);
  /**
   * True while `show` is taking a panel down in order to put another
   * one up.
   *
   * `overlay.hide()` reports itself through `onClose`, which is how a
   * press outside the menu gets back here — and walking from Edit to
   * Data hides one panel and shows the next, so without this the
   * hide's own report arrived *after* the new state was written and
   * closed the menu that had just opened. ArrowRight shut the bar
   * instead of moving along it.
   */
  let swapping = false;

  const context = {
    enabled: (id: CommandId) => inputs.enabled.value(id),
    labelOf: (id: CommandId) => COMMANDS[id].label
  };

  /**
   * The overlay follows the state rather than being opened beside it.
   *
   * Every path that opens or closes a menu — a key, a click, a
   * choice, a press outside — writes the state and nothing else, so
   * there is one place where "open" becomes a panel on screen and no
   * way for the two to disagree about which menu that is.
   */
  const show = (next: MenuBarState): void => {
    const was = state.value;
    state.next(next);
    if (!next.open) {
      if (overlay.isOpen()) {
        overlay.hide();
      }
      return;
    }
    if (overlay.isOpen() && was.focused === next.focused) {
      return;
    }
    // A different menu is a different panel in a different place, so
    // it is closed and reopened rather than moved.
    if (overlay.isOpen()) {
      swapping = true;
      overlay.hide();
      swapping = false;
    }
    const anchor = titles[next.focused];
    overlay.show(panel(next.focused), {
      anchor,
      environment: anchor,
      placement: 'bottom-start',
      offset: 2,
      /**
       * No backdrop, and `clickOutside` instead.
       *
       * `dismissOnOutsidePress` inserts a full-screen box over
       * everything to catch the press — and a box over everything is
       * a box over the menu bar, so the titles stopped receiving
       * `pointerEnter` and moving along the bar with the pointer did
       * nothing. `clickOutside` hears the press at the root and lets
       * it through, which is what its own documentation says a menu
       * wants. The titles are `except`ed, or pressing the open menu's
       * title would close it and reopen it in one press.
       */
      dismissOnOutsidePress: false,
      onClose: () => {
        if (!swapping && state.value.open) {
          state.next({ ...state.value, open: false, active: -1 });
        }
      }
    });
  };

  const choose = (id: CommandId): void => {
    if (!context.enabled(id)) {
      return;
    }
    show({ ...state.value, open: false, active: -1 });
    inputs.onChoose.value(id);
  };

  const onKeyDown = (event: UiKeyboardEvent): void => {
    const step = menuBarStep(state.value, event.key, menus, context);
    if (step === null) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    show(step.state);
    if (step.choose !== undefined) {
      inputs.onChoose.value(step.choose);
    }
    if (step.dismiss === true) {
      inputs.onDismiss.value();
    }
  };

  /** The panel for one menu: its commands, their keys, and the rules between. */
  const panel = (index: number) =>
    Column(
      {
        minWidth: 232,
        padding: 4,
        gap: 1,
        backgroundColor: 'surface',
        borderColor: 'border',
        borderWidth: 1,
        borderRadius: 8,
        role: 'menu',
        label: menus[index].label,
        modifiers: [
          clickOutside({
            onOutside: () => show({ ...state.value, open: false, active: -1 }),
            except: () => titles
          })
        ]
      },
      ...menus[index].entries.map((entry, at) => item(entry, at, index))
    );

  const item = (entry: MenuEntry, at: number, menu: number) => {
    if (entry === SEPARATOR) {
      return (
        <box
          key={`rule-${at}`}
          height={1}
          marginTop={3}
          marginBottom={3}
          backgroundColor="border"
          role="separator"
        />
      );
    }
    const command = COMMANDS[entry];
    const on = context.enabled(entry);
    const highlighted: Observable<boolean> = state.pipe(
      map(current => current.open && current.focused === menu && current.active === at)
    );
    return (
      <row
        key={entry}
        gap={24}
        paddingLeft={10}
        paddingRight={10}
        paddingTop={5}
        paddingBottom={5}
        borderRadius={4}
        y="center"
        cursor={on ? 'pointer' : 'default'}
        backgroundColor={highlighted.pipe(map(is => (is ? 'controlBackgroundHovered' : 'transparent')))}
        role="menuitem"
        label={command.label}
        disabled={!on}
        /**
         * Hovering moves the *same* highlight the arrows move.
         *
         * One highlight and not two: a menu with a keyboard highlight
         * on Undo and a hover highlight on Paste is a menu that
         * cannot say what Enter will do. A disabled item is skipped,
         * on the rule `seek` already follows — the highlight only
         * ever rests where Enter would work.
         */
        onPointerEnter={() => {
          if (on && state.value.active !== at) {
            state.next({ ...state.value, active: at });
          }
        }}
        onClick={() => choose(entry)}>
        <text
          text={command.label}
          flex={1}
          fontSize={12}
          color={on ? 'controlForeground' : 'controlForegroundDisabled'}
          selectable={false}
        />
        <text
          text={command.accelerator === undefined ? '' : acceleratorLabel(command.accelerator)}
          fontSize={11}
          color="textMuted"
          selectable={false}
        />
      </row>
    );
  };

  ctx.onUnmount(() => {
    if (overlay.isOpen()) {
      overlay.hide();
    }
  });

  return (
    <row
      ref={inputs.ref.value ?? undefined}
      gap={2}
      paddingLeft={4}
      paddingRight={4}
      y="center"
      focusable={true}
      role="menubar"
      label="Main menu"
      onKeyDown={onKeyDown}>
      {menus.map((menu, index) => (
        <row
          key={menu.id}
          ref={(node: UiNode | null) => (titles[index] = node)}
          paddingLeft={9}
          paddingRight={9}
          paddingTop={4}
          paddingBottom={4}
          borderRadius={5}
          cursor="pointer"
          backgroundColor={combineLatest([state, hoveredTitle]).pipe(
            map(([current, hovered]) =>
              (current.open && current.focused === index) || hovered === index
                ? 'controlBackgroundHovered'
                : 'transparent'
            )
          )}
          onPointerEnter={() => {
            hoveredTitle.next(index);
            // With a menu already open, moving along the bar opens the
            // one under the pointer. Without that, a bar is something
            // you have to click four times to read.
            if (state.value.open && state.value.focused !== index) {
              show({ focused: index, open: true, active: -1 });
            }
          }}
          onPointerLeave={() => {
            if (hoveredTitle.value === index) {
              hoveredTitle.next(-1);
            }
          }}
          onClick={() => {
            const current = state.value;
            // Clicking the menu that is already open closes it, which
            // is what a title bar does everywhere.
            show(
              current.open && current.focused === index
                ? { focused: index, open: false, active: -1 }
                : { focused: index, open: true, active: -1 }
            );
          }}>
          <text text={menu.label} fontSize={12} color="text" selectable={false} />
        </row>
      ))}
    </row>
  );
}
