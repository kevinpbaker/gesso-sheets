import { combineLatest, map, type Observable } from 'rxjs';

import {
  percent,
  type UiKeyboardEvent,
  type UiNode,
  type UiSemanticState,
  type UiTextChangeEvent
} from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import type { ColourScale, ConditionalPaint, ConditionalTest } from '../sheet/Conditional';
import type { ValidationRule } from '../sheet/Validation';
import { Sheet, type SheetConditionalRule } from './SheetContract';
import type { SheetEditing } from './SheetEditing';

/**
 * Formats that think, and what a cell is allowed to hold, as one bar.
 *
 * A row in the flow rather than a floating panel, for the reason the
 * find bar is one: opening it must not cover the range the rule is
 * about. It costs a row of height, which is the cheaper of the two
 * mistakes.
 *
 * One bar for both kinds because they are the same gesture — pick a
 * range, say what about it, press Enter — and two bars would be two
 * places to learn it. The half that is showing is a tab, which is
 * also what the find bar does with Replace.
 *
 * The rule crosses as data. `addConditional` carries a tagged union
 * and not a predicate, because a `postMessage` carries no closures;
 * the range is the selection's and is filled in on the other side,
 * so this says what the rule is and never where.
 */

export type RulesTab = 'format' | 'validation';

export interface RulesBarProps {
  readonly editing: SheetEditing;
  readonly tab: RulesTab;
  readonly onTab: (tab: RulesTab) => void;
  readonly onClose: () => void;
  readonly ref?: (node: UiNode | null) => void;
}

/** The fills a rule can paint, named rather than free; see `TAB_COLOURS`. */
const FILLS: readonly { readonly name: string; readonly fill: string; readonly color: string }[] = [
  { name: 'Red', fill: '#fce8e6', color: '#c5221f' },
  { name: 'Yellow', fill: '#fef7e0', color: '#b06000' },
  { name: 'Green', fill: '#e6f4ea', color: '#137333' },
  { name: 'Blue', fill: '#e8f0fe', color: '#1967d2' },
  { name: 'Grey', fill: '#f1f3f4', color: '#5f6368' }
];

/** The `fill` that means "the colours the rule already had"; see `current`. */
const CURRENT = -1;

/** What a rule can ask, as the bar offers it. */
const TESTS: readonly { readonly id: string; readonly label: string; readonly values: number }[] = [
  { id: 'greaterThan', label: 'Greater than', values: 1 },
  { id: 'lessThan', label: 'Less than', values: 1 },
  { id: 'between', label: 'Between', values: 2 },
  { id: 'equalTo', label: 'Equal to', values: 1 },
  { id: 'textContains', label: 'Text contains', values: 1 },
  { id: 'notEmpty', label: 'Not empty', values: 0 },
  { id: 'formula', label: 'Custom formula', values: 1 },
  { id: 'scale', label: 'Colour scale', values: 0 }
];

const CHECKS: readonly { readonly id: ValidationRule['kind']; readonly label: string }[] = [
  { id: 'list', label: 'One of a list' },
  { id: 'number', label: 'A number' },
  { id: 'date', label: 'A date' },
  { id: 'text', label: 'Text' }
];

export function RulesBar(inputs: Inputs<RulesBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;

  const test = internalState('greaterThan');
  const first = internalState('');
  const second = internalState('');
  /** An index into `FILLS`, or `CURRENT` for the loaded rule's own colours. */
  const fill = internalState(0);
  /**
   * The colours of the rule the bar was opened on, when they are not
   * one of `FILLS` — the seeded pink-to-green scale, or bold red text
   * with no fill. Kept whole so that Apply puts back what was there
   * rather than the nearest swatch, which would quietly repaint it.
   */
  const current = internalState<Pick<SheetConditionalRule, 'paint' | 'scale'> | null>(null);

  const check = internalState<ValidationRule['kind']>('list');
  const allowed = internalState('');
  const strict = internalState(false);

  /**
   * The bar opens on the rule the active cell is already under.
   *
   * Only when that rule *changes*, which is the guard: the key moves
   * on every arrow key, and refilling the fields on each one would
   * throw away a value somebody typed before selecting the range it
   * was for. Moving onto a cell with no rule clears what an earlier
   * cell's rule put there, and leaves a draft of somebody's own alone.
   */
  let loadedFormat = 'null';
  let loadedCheck = 'null';
  ctx.effect(sheet.view.activeRules, active => {
    const format = JSON.stringify(active.conditional);
    if (format !== loadedFormat) {
      const had = loadedFormat !== 'null';
      loadedFormat = format;
      if (active.conditional !== null) {
        loadFormat(active.conditional);
      } else if (had) {
        test.value = 'greaterThan';
        first.value = '';
        second.value = '';
        fill.value = 0;
        current.value = null;
      }
    }
    const validation = JSON.stringify(active.validation);
    if (validation !== loadedCheck) {
      const had = loadedCheck !== 'null';
      loadedCheck = validation;
      if (active.validation !== null) {
        loadCheck(active.validation.rule, active.validation.strict);
      } else if (had) {
        check.value = 'list';
        allowed.value = '';
        strict.value = false;
      }
    }
  });

  function loadFormat(rule: SheetConditionalRule): void {
    if (rule.scale !== undefined) {
      const scale = rule.scale;
      const at = FILLS.findIndex(
        entry => scale.from.toLowerCase() === '#ffffff' && scale.middle === undefined && entry.color === scale.to.toLowerCase()
      );
      test.value = 'scale';
      current.value = at === -1 ? { scale } : null;
      fill.value = at === -1 ? CURRENT : at;
      return;
    }
    const held = rule.test;
    if (held === null || !TESTS.some(entry => entry.id === held.kind)) {
      return;
    }
    test.value = held.kind;
    first.value =
      held.kind === 'between'
        ? String(held.low)
        : held.kind === 'textContains'
          ? held.text
          : held.kind === 'formula'
            ? held.input
            : 'value' in held
              ? String(held.value)
              : '';
    second.value = held.kind === 'between' ? String(held.high) : '';
    const paint = rule.paint ?? {};
    const at = FILLS.findIndex(
      entry =>
        entry.fill === paint.fill?.toLowerCase() &&
        entry.color === paint.color?.toLowerCase() &&
        paint.bold !== true &&
        paint.italic !== true
    );
    current.value = at === -1 ? { paint } : null;
    fill.value = at === -1 ? CURRENT : at;
  }

  function loadCheck(rule: ValidationRule, refuses: boolean): void {
    check.value = rule.kind;
    strict.value = refuses;
    switch (rule.kind) {
      case 'list':
        allowed.value = rule.values.join(', ');
        return;
      case 'number':
        allowed.value =
          rule.min === undefined && rule.max === undefined
            ? ''
            : `${rule.min ?? ''}, ${rule.max ?? ''}`.replace(/, $/, '');
        return;
      case 'text':
        allowed.value = rule.maxLength === undefined ? '' : String(rule.maxLength);
        return;
      default:
        allowed.value = '';
    }
  }

  const close = (): void => {
    inputs.onClose.value();
    edit.focusSheet();
  };

  /**
   * A number, or null when what was typed is not one.
   *
   * Null rather than zero, because zero is a threshold somebody may
   * mean and an empty box is not.
   */
  const numberOf = (text: string): number | null => {
    const trimmed = text.trim();
    if (trimmed === '' || !/^[-+]?(\d+\.?\d*|\.\d+)$/.test(trimmed)) {
      return null;
    }
    return Number(trimmed);
  };

  const paintOf = (): ConditionalPaint => {
    const kept = current.value?.paint;
    if (fill.value === CURRENT && kept !== undefined) {
      return kept;
    }
    const chosen = FILLS[fill.value] ?? FILLS[0];
    return { fill: chosen.fill, color: chosen.color };
  };

  const testOf = (): ConditionalTest | null => {
    const one = numberOf(first.value);
    switch (test.value) {
      case 'greaterThan':
        return one === null ? null : { kind: 'greaterThan', value: one };
      case 'lessThan':
        return one === null ? null : { kind: 'lessThan', value: one };
      case 'between': {
        const other = numberOf(second.value);
        return one === null || other === null
          ? null
          : { kind: 'between', low: Math.min(one, other), high: Math.max(one, other) };
      }
      case 'equalTo':
        return { kind: 'equalTo', value: one ?? first.value };
      case 'textContains':
        return first.value === '' ? null : { kind: 'textContains', text: first.value };
      case 'notEmpty':
        return { kind: 'notEmpty' };
      case 'formula':
        return first.value === '' ? null : { kind: 'formula', input: first.value };
      default:
        return null;
    }
  };

  const applyFormat = (): void => {
    if (test.value === 'scale') {
      const kept = current.value?.scale;
      const chosen = FILLS[fill.value] ?? FILLS[0];
      sheet.send.addConditional({
        test: null,
        scale: fill.value === CURRENT && kept !== undefined ? kept : { from: '#ffffff', to: chosen.color }
      });
      close();
      return;
    }
    const rule = testOf();
    if (rule === null) {
      return;
    }
    sheet.send.addConditional({ test: rule, paint: paintOf() });
    close();
  };

  const ruleOf = (): ValidationRule | null => {
    // Kept by position for a number: `, 10` is "at most ten", and
    // closing up the empty part would read the ten as the minimum.
    const positions = allowed.value.split(',').map(part => part.trim());
    const parts = positions.filter(part => part !== '');
    switch (check.value) {
      case 'list':
        return parts.length === 0 ? null : { kind: 'list', values: parts };
      case 'number': {
        const low = numberOf(positions[0] ?? '');
        const high = numberOf(positions[1] ?? '');
        return {
          kind: 'number',
          ...(low === null ? {} : { min: low }),
          ...(high === null ? {} : { max: high })
        };
      }
      case 'date':
        return { kind: 'date' };
      case 'text': {
        const length = numberOf(parts[0] ?? '');
        return length === null ? { kind: 'text' } : { kind: 'text', maxLength: length };
      }
    }
  };

  const applyValidation = (): void => {
    const rule = ruleOf();
    if (rule === null) {
      return;
    }
    sheet.send.addValidation(rule, strict.value, '');
    close();
  };

  const onFieldKey = (run: () => void) => (event: UiKeyboardEvent) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      run();
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      close();
    }
  };

  const field = (
    label: string,
    value: { value: string } & Observable<string>,
    run: () => void,
    width: number,
    ref?: (node: UiNode | null) => void
  ) => (
    <editabletext
      key={`field-${label}`}
      ref={ref}
      value={value as never}
      width={width}
      flexShrink={0}
      fontSize={12}
      color="text"
      textWrap="none"
      verticalAlign="middle"
      backgroundColor="background"
      borderColor="border"
      borderWidth={1}
      padding={4}
      role="textbox"
      label={label}
      onInput={(event: UiTextChangeEvent) => (value.value = event.value)}
      onKeyDown={onFieldKey(run)}
    />
  );

  /**
   * A choice among a few, drawn as a row of radios.
   *
   * A dropdown would be one node instead of eight and a popup to
   * manage; a row is readable at a glance and every option is a tab
   * stop's worth of arrow away, which is what the toolbar decided for
   * the same trade.
   */
  const choice = <T extends string>(
    label: string,
    options: readonly { readonly id: T; readonly label: string }[],
    held: { value: T } & Observable<T>,
    onPick?: () => void
  ) => (
    <row key={`choice-${label}`} gap={2} y="center" role="radiogroup" label={label}>
      {options.map(option => (
        <button
          key={option.id}
          focusable={false}
          paddingLeft={7}
          paddingRight={7}
          paddingTop={3}
          paddingBottom={3}
          borderRadius={5}
          cursor="pointer"
          backgroundColor={held.pipe(map(is => (is === option.id ? 'controlBackgroundHovered' : 'transparent')))}
          borderColor="controlBorder"
          borderWidth={1}
          role="radio"
          label={option.label}
          states={held.pipe(map((is): readonly UiSemanticState[] => (is === option.id ? ['checked'] : [])))}
          onClick={() => {
            held.value = option.id;
            onPick?.();
          }}>
          <text text={option.label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
        </button>
      ))}
    </row>
  );

  /**
   * A scale drawn as its stops side by side, inside the swatch's
   * border.
   *
   * A scale is two or three colours, and a square of any one of them
   * says the wrong thing: the pale fill the swatches show for a plain
   * rule is not what a scale made from it runs to, and the middle of
   * a pink-yellow-green scale is only yellow.
   */
  const stops = (colours: readonly string[]) => (
    <row width={percent(100)} height={percent(100)}>
      {colours.map((colour, at) => (
        <box key={`stop-${at}`} flex={1} height={percent(100)} backgroundColor={colour} />
      ))}
    </row>
  );

  const stopsOf = (scale: ColourScale): readonly string[] =>
    scale.middle === undefined ? [scale.from, scale.to] : [scale.from, scale.middle, scale.to];

  /**
   * The swatch for colours none of `FILLS` has, drawn in them.
   *
   * A scale shows its stops; a paint with no fill shows its text
   * colour as a letter, because red bold text on white is a thing
   * somebody has to be able to see is selected.
   */
  const currentSwatch = (kept: Pick<SheetConditionalRule, 'paint' | 'scale'>) => (
    <button
      key="current"
      focusable={false}
      width={20}
      height={18}
      borderRadius={4}
      cursor="pointer"
      padding={kept.scale === undefined ? 0 : 2}
      backgroundColor={kept.scale === undefined ? (kept.paint?.fill ?? '#ffffff') : '#ffffff'}
      borderColor={fill.pipe(map(is => (is === CURRENT ? 'focusRing' : 'controlBorder')))}
      borderWidth={fill.pipe(map(is => (is === CURRENT ? 2 : 1)))}
      x="center"
      y="center"
      role="radio"
      label="Current colours"
      states={fill.pipe(map((is): readonly UiSemanticState[] => (is === CURRENT ? ['checked'] : [])))}
      onClick={() => (fill.value = CURRENT)}>
      {kept.scale !== undefined ? (
        stops(stopsOf(kept.scale))
      ) : kept.paint !== undefined && kept.paint.fill === undefined ? (
        <text
          text="A"
          fontSize={11}
          fontWeight={kept.paint.bold === true ? 'bold' : 'normal'}
          color={kept.paint.color ?? 'text'}
          selectable={false}
        />
      ) : (
        <box width={0} height={0} />
      )}
    </button>
  );

  /**
   * The five colours, drawn as what choosing one will paint: its fill
   * for a rule, and white to its colour for a scale, which is what
   * `applyFormat` sends.
   */
  const swatches = (kept: Pick<SheetConditionalRule, 'paint' | 'scale'> | null, scale: boolean) => (
    <row key="swatches" gap={3} y="center" role="radiogroup" label="Colour">
      {kept === null ? [] : [currentSwatch(kept)]}
      {FILLS.map((entry, index) => (
        <button
          key={entry.name}
          focusable={false}
          width={20}
          height={18}
          borderRadius={4}
          cursor="pointer"
          padding={scale ? 2 : 0}
          backgroundColor={scale ? '#ffffff' : entry.fill}
          borderColor={fill.pipe(map(is => (is === index ? 'focusRing' : 'controlBorder')))}
          borderWidth={fill.pipe(map(is => (is === index ? 2 : 1)))}
          role="radio"
          label={entry.name}
          states={fill.pipe(map((is): readonly UiSemanticState[] => (is === index ? ['checked'] : [])))}
          onClick={() => (fill.value = index)}
        >
          {scale ? stops(['#ffffff', entry.color]) : <box width={0} height={0} />}
        </button>
      ))}
    </row>
  );

  const toggle = (label: string, held: { value: boolean } & Observable<boolean>) => (
    <button
      key={`toggle-${label}`}
      focusable={false}
      paddingLeft={8}
      paddingRight={8}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={5}
      cursor="pointer"
      backgroundColor={held.pipe(map(is => (is ? 'controlBackgroundHovered' : 'transparent')))}
      borderColor="controlBorder"
      borderWidth={1}
      role="checkbox"
      label={label}
      states={held.pipe(map((is): readonly UiSemanticState[] => (is ? ['checked'] : [])))}
      onClick={() => (held.value = !held.value)}>
      <text text={label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );

  const button = (label: string, onClick: () => void) => (
    <button
      key={`button-${label}`}
      onClick={onClick}
      label={label}
      paddingLeft={9}
      paddingRight={9}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={5}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );

  /**
   * The middle of the bar, from one observable rather than two
   * nested ones.
   *
   * Written as `{tab.pipe(map(… {values.pipe(map(…))} …))}` the bar
   * mounted, reached the accessibility tree with every role in place,
   * and laid out at zero by zero: an observable of elements inside an
   * observable of elements is a subtree the engine never measures.
   * Every spec passed, because a spec asks what a node's properties
   * are and the properties were right; a browser showed a bar one
   * pixel tall.
   */
  const middle = combineLatest([inputs.tab, test, current]).pipe(
    map(([which, id, kept]) => {
      if (which !== 'format') {
        return [
          choice('Allow', CHECKS, check),
          field('Allowed values', allowed, applyValidation, 200, inputs.ref.value ?? undefined),
          toggle('Refuse anything else', strict),
          button('Apply', applyValidation)
        ];
      }
      const count = TESTS.find(entry => entry.id === id)?.values ?? 0;
      return [
        choice('Condition', TESTS, test),
        ...(count === 0 ? [] : [field('Value', first, applyFormat, 110, inputs.ref.value ?? undefined)]),
        ...(count > 1 ? [field('And', second, applyFormat, 80)] : []),
        swatches(kept, id === 'scale'),
        button('Apply', applyFormat)
      ];
    })
  );

  return (
    <row
      width={percent(100)}
      flexShrink={0}
      y="center"
      gap={8}
      flexWrap="wrap"
      rowGap={4}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      backgroundColor="surface"
      role="form"
      label="Rules for the selection">
      <row gap={2} y="center" role="radiogroup" label="Rule kind">
        {(['format', 'validation'] as const).map(which => (
          <button
            key={which}
            focusable={false}
            paddingLeft={7}
            paddingRight={7}
            paddingTop={3}
            paddingBottom={3}
            borderRadius={5}
            cursor="pointer"
            backgroundColor={inputs.tab.pipe(map(is => (is === which ? 'controlBackgroundHovered' : 'transparent')))}
            borderColor="controlBorder"
            borderWidth={1}
            role="radio"
            label={which === 'format' ? 'Format' : 'Allowed'}
            states={inputs.tab.pipe(map((is): readonly UiSemanticState[] => (is === which ? ['checked'] : [])))}
            onClick={() => inputs.onTab.value(which)}>
            <text
              text={which === 'format' ? 'Format' : 'Allowed'}
              fontSize={11}
              textWrap="none"
              color="controlForeground"
              selectable={false}
            />
          </button>
        ))}
      </row>
      {middle}
      <box flex={1} minWidth={0} />
      {button('Clear rules', () => {
        sheet.send.clearRules();
        close();
      })}
      {button('Close', close)}
    </row>
  );
}
