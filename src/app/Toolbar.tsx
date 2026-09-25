import { BehaviorSubject, combineLatest, map, of, type Observable } from 'rxjs';

import { type UiKeyboardEvent, type UiNode, type UiSemanticState } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

/**
 * The row of buttons under the menu bar.
 *
 * **One tab stop, with the arrows moving inside it.** The same
 * decision as the menu bar's and for a plainer reason: the toolbar
 * grew from three buttons to fifteen in this phase, and as fifteen
 * tab stops it put the grid fifteen presses away from the keyboard.
 * A spreadsheet whose sheet is the last thing you can reach is a
 * spreadsheet nobody will use without a mouse, and the specs said so
 * before a person could — `Tab never reached the grid` is what
 * Phase 9 got for adding a Format section.
 *
 * It is also what ARIA's toolbar pattern asks for, which
 * `gesso-components`' `Toolbar` deliberately does not do: that one
 * groups the buttons for a screen reader and leaves them as ordinary
 * stops. Grouping is the half this application already had.
 */

export interface ToolbarItem {
  readonly id: string;
  /** What is drawn in the button. */
  readonly text: string;
  /** What a screen reader says, when that is not the text. */
  readonly label?: string;
  readonly pressed?: Observable<boolean>;
  readonly enabled?: Observable<boolean>;
  readonly onRun: () => void;
  readonly weight?: number | 'normal' | 'bold';
  /** A gap before this item, for the groups people read by. */
  readonly startsGroup?: boolean;
}

export interface ToolbarProps {
  readonly items: readonly ToolbarItem[];
  readonly label?: string;
  readonly ref?: (node: UiNode | null) => void;
}

export function Toolbar(inputs: Inputs<ToolbarProps>, _ctx: ComponentContext) {
  const items = inputs.items.value;
  /** Which button the roving highlight is on. */
  const at = new BehaviorSubject(0);

  const step = (by: number): void => {
    if (items.length === 0) {
      return;
    }
    at.next(((at.value + by) % items.length + items.length) % items.length);
  };

  const onKeyDown = (event: UiKeyboardEvent): void => {
    switch (event.key) {
      case 'ArrowRight':
        step(1);
        break;
      case 'ArrowLeft':
        step(-1);
        break;
      case 'Home':
        at.next(0);
        break;
      case 'End':
        at.next(items.length - 1);
        break;
      case 'Enter':
      case ' ':
        items[at.value]?.onRun();
        break;
      default:
        // Everything else belongs to whatever is underneath, which is
        // what lets Tab out of the toolbar keep working.
        return;
    }
    event.preventDefault();
    event.stopPropagation();
  };

  return (
    <row
      ref={inputs.ref.value ?? undefined}
      y="center"
      focusable={true}
      role="toolbar"
      label={inputs.label.value ?? 'Toolbar'}
      onKeyDown={onKeyDown}>
      {items.map((item, index) => button(item, index, at))}
    </row>
  );
}

function button(item: ToolbarItem, index: number, at: BehaviorSubject<number>) {
  const pressed = item.pressed ?? of(false);
  const enabled = item.enabled ?? of(true);
  /**
   * The roving highlight is drawn as a ring rather than as focus,
   * because focus is on the toolbar itself. Without something
   * visible, the arrows would move a selection nobody can see.
   */
  const carried = combineLatest([pressed, at]).pipe(
    map(([is, current]) => (is ? 'controlBackgroundHovered' : current === index ? 'controlBackgroundHovered' : 'controlBackground'))
  );
  return (
    <button
      key={item.id}
      /**
       * Not a tab stop. The toolbar is the stop and the arrows move
       * inside it, so every button here is reachable and none of them
       * stands between the keyboard and the sheet. `focusable: false`
       * is the opt-out `isNodeFocusable` offers for exactly this —
       * a button is focusable by type unless it says otherwise.
       */
      focusable={false}
      onClick={() => {
        at.next(index);
        item.onRun();
      }}
      label={item.label ?? item.text}
      paddingLeft={9}
      paddingRight={9}
      paddingTop={4}
      paddingBottom={4}
      marginLeft={item.startsGroup === true ? 10 : 2}
      borderRadius={6}
      backgroundColor={carried}
      borderColor={at.pipe(map(current => (current === index ? 'focusRing' : 'controlBorder')))}
      borderWidth={1}
      cursor="pointer"
      opacity={enabled.pipe(map(is => (is ? 1 : 0.4)))}
      states={pressed.pipe(map((is): readonly UiSemanticState[] => (is ? ['pressed'] : [])))}>
      <text
        text={item.text}
        fontSize={12}
        fontWeight={item.weight ?? 'normal'}
        textAlign="center"
        textWrap="none"
        color="controlForeground"
        selectable={false}
      />
    </button>
  );
}
