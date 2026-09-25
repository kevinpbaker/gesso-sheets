/**
 * The menu bar's keyboard, which now lives in `gesso-components`.
 *
 * This file was the traversal itself for two phases, and the roadmap
 * said the upstreaming decision could be taken once the shape was
 * known. It is known, so it has been taken: `menuBarStep` is a peer of
 * `Menu` there, generic in the command type, with its own specs
 * against its own menus. The reason it could not be `Menu` has not
 * changed and is written down at both ends — a popup that traps focus
 * cannot let ArrowLeft reach the bar, and walking between menus with
 * the arrows is most of what makes a bar a bar.
 *
 * What is left here is the names this application uses. `CLOSED` and
 * `MenuBarContext` read better in `MenuBar.tsx` than
 * `MENU_BAR_CLOSED` and `MenuBarContext<CommandId>` do, and binding
 * the command type once is worth a file.
 */

import {
  MENU_BAR_CLOSED,
  menuBarStep as step,
  type MenuBarContext as GenericContext,
  type MenuBarState,
  type MenuBarStep as GenericStep
} from 'gesso-components';

import type { CommandId, MenuDefinition } from './SheetCommands';

export type { MenuBarState };

/** This application's commands, in the generic model's terms. */
export type MenuBarContext = GenericContext<CommandId>;
export type MenuBarStep = GenericStep<CommandId>;

export const CLOSED = MENU_BAR_CLOSED;

/** `menuBarStep`, bound to this application's commands and menus. */
export function menuBarStep(
  state: MenuBarState,
  key: string,
  menus: readonly MenuDefinition[],
  context: MenuBarContext
): MenuBarStep | null {
  return step(state, key, menus, context);
}
