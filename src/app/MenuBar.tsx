import type { UiNode } from 'gesso-core';
import { MenuBar as Bar } from 'gesso-components';
import { createComponent, type ComponentContext, type Inputs } from 'gesso-framework';

import { acceleratorLabel, COMMANDS, MENUS, type CommandId, type MenuDefinition } from './SheetCommands';

/**
 * The menu bar, which is `gesso-components`' now.
 *
 * This file was the whole thing for two phases: the titles, the
 * panel, the overlay that draws without taking the keyboard, and the
 * hover that moves the same highlight the arrows move. None of that
 * was ever this application's — a bar with focus on the strip rather
 * than in the popup is what a bar *is*, and the reason it could not
 * be `Menu`. The roadmap said the upstreaming decision could be taken
 * once the shape was known, and two phases of using it is knowing.
 *
 * What is left here is what a spreadsheet's commands are called and
 * how this application writes a shortcut. Those are the three props
 * the component asks for and the three it deliberately does not
 * guess: what a shortcut is called depends on the platform and on
 * what an application has decided to call its modifiers.
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

export function MenuBar(inputs: Inputs<MenuBarProps>, _ctx: ComponentContext) {
  return createComponent(Bar<CommandId>, {
    menus: inputs.menus.value ?? MENUS,
    enabled: (id: CommandId) => inputs.enabled.value(id),
    labelOf: (id: CommandId) => COMMANDS[id].label,
    acceleratorOf: (id: CommandId) => {
      const accelerator = COMMANDS[id].accelerator;
      return accelerator === undefined ? undefined : acceleratorLabel(accelerator);
    },
    onChoose: (id: CommandId) => inputs.onChoose.value(id),
    onDismiss: () => inputs.onDismiss.value(),
    barRef: inputs.ref.value ?? undefined
  });
}
