import { BehaviorSubject, combineLatest, map, of, type Observable } from 'rxjs';

import { type UiKeyboardEvent, type UiNode, type UiSemanticState } from 'gesso-core';
import { Icon, tooltip } from 'gesso-components';
import { createComponent, type ComponentContext, type Inputs } from 'gesso-framework';

import type { Glyph } from './icons';

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

interface ToolbarItemBase {
  readonly id: string;
  /**
   * What the tooltip says, when the label alone is not what somebody
   * hovering wants to read.
   *
   * The label is the button's name and a screen reader reads it on
   * arrival; a tooltip is read by somebody who has stopped on the
   * button because they are not sure, and the accelerator is the
   * thing they get for stopping. Absent, the label is used, so a
   * button always has a tooltip and never a silent one.
   */
  readonly tip?: string;
  readonly pressed?: Observable<boolean>;
  readonly enabled?: Observable<boolean>;
  readonly onRun: () => void;
  /** A gap before this item, for the groups people read by. */
  readonly startsGroup?: boolean;
}

/** A button drawn as a glyph; see `icons.ts` for where the glyphs come from. */
export interface ToolbarIconItem extends ToolbarItemBase {
  readonly icon: Glyph;
  /**
   * What a screen reader says. Required, unlike the text item's,
   * because a picture says nothing: an icon button with no label is a
   * button called nothing at all, and the accessibility tree is what
   * this application's own specs read it through.
   */
  readonly label: string;
  readonly text?: never;
  readonly weight?: never;
}

/**
 * A button drawn as text.
 *
 * Still here, and deliberately: the decimal-place buttons say `.0←`
 * and `.00→` because what they do is about digits, and no icon set
 * has a glyph for *one fewer decimal place* that reads faster than
 * the digits themselves do.
 */
export interface ToolbarTextItem extends ToolbarItemBase {
  /** What is drawn in the button. */
  readonly text: string;
  /** What a screen reader says, when that is not the text. */
  readonly label?: string;
  readonly weight?: number | 'normal' | 'bold';
  readonly icon?: never;
}

export type ToolbarItem = ToolbarIconItem | ToolbarTextItem;

export interface ToolbarProps {
  readonly items: readonly ToolbarItem[];
  readonly label?: string;
  readonly ref?: (node: UiNode | null) => void;
}

export function Toolbar(inputs: Inputs<ToolbarProps>, ctx: ComponentContext) {
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
      {items.map((item, index) => button(item, index, at, ctx))}
    </row>
  );
}

function button(item: ToolbarItem, index: number, at: BehaviorSubject<number>, ctx: ComponentContext) {
  const pressed = item.pressed ?? of(false);
  const enabled = item.enabled ?? of(true);
  /** What the tooltip says: what it was given, or the button's own name. */
  const tip = item.tip ?? (item.icon === undefined ? item.label ?? item.text : item.label);
  const hovered = new BehaviorSubject(false);
  /**
   * Three reasons a button is lit, and one colour for all of them.
   *
   * It is already on, the roving highlight is resting on it, or the
   * pointer is over it. The roving highlight has to be *visible*
   * rather than shown as focus, because focus is on the toolbar
   * itself — without it the arrows would move a selection nobody can
   * see.
   */
  const carried = combineLatest([pressed, at, hovered]).pipe(
    map(([is, current, over]) =>
      is || current === index || over ? 'controlBackgroundHovered' : 'controlBackground'
    )
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
      /**
       * The label, after a pause, under the button.
       *
       * Under rather than over, which is the one thing this has to say
       * about placement: the toolbar is the second row on the screen,
       * so a tooltip above it would be flipped back down by the
       * overlay at the viewport edge every single time, and asking for
       * the answer is cheaper than being corrected into it.
       *
       * A modifier rather than the `Tooltip` component because the
       * component wraps its trigger in a box to listen to it, and this
       * row already has thirteen nodes it needs and none it does not.
       */
      modifiers={[tooltip(ctx, { text: tip, placement: 'bottom' })]}
      onPointerEnter={() => hovered.next(true)}
      onPointerLeave={() => hovered.next(false)}
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
      {item.icon === undefined ? (
        <text
          text={item.text}
          fontSize={12}
          fontWeight={item.weight ?? 'normal'}
          textAlign="center"
          textWrap="none"
          color="controlForeground"
          selectable={false}
        />
      ) : (
        glyph(item.icon)
      )}
    </button>
  );
}

/**
 * One icon, at the size and colour every button here draws it.
 *
 * **One colour, on purpose.** `IconRasterizer` keys its cache on
 * (path, size, colour, style), so an icon that changed colour on hover
 * would be two rasters of the same glyph and a fresh
 * `createImageBitmap` the first time a pointer crossed it. The button
 * already says *lit* with its background and *unavailable* with its
 * opacity — neither of which touches the glyph — so the row of icons
 * costs one raster each for as long as the theme holds.
 *
 * No `label`: the button carries it, and an icon that repeated it
 * would have a screen reader say the name twice.
 */
function glyph(icon: Glyph) {
  return createComponent(Icon, {
    path: icon.path,
    viewBox: icon.viewBox,
    size: 16,
    color: 'controlForeground',
    style: icon.style,
    strokeWidth: icon.strokeWidth ?? 1.5,
    fillRule: icon.fillRule ?? 'nonzero'
  });
}
