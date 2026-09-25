import { map, type Observable } from 'rxjs';

import {
  percent,
  type UiKeyboardEvent,
  type UiNode,
  type UiSemanticState,
  type UiTextChangeEvent
} from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import type { ConditionalPaint, ConditionalTest } from '../sheet/Conditional';
import type { ValidationRule } from '../sheet/Validation';
import { Sheet } from './SheetContract';
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
  const fill = internalState(0);

  const check = internalState<ValidationRule['kind']>('list');
  const allowed = internalState('');
  const strict = internalState(false);

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
      const chosen = FILLS[fill.value] ?? FILLS[0];
      sheet.send.addConditional({ test: null, scale: { from: '#ffffff', to: chosen.color } });
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
    const parts = allowed.value
      .split(',')
      .map(part => part.trim())
      .filter(part => part !== '');
    switch (check.value) {
      case 'list':
        return parts.length === 0 ? null : { kind: 'list', values: parts };
      case 'number': {
        const low = numberOf(parts[0] ?? '');
        const high = numberOf(parts[1] ?? '');
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
      ref={ref}
      value={value as never}
      width={width}
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
    <row gap={2} y="center" role="radiogroup" label={label}>
      {options.map(option => (
        <row
          key={option.id}
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
        </row>
      ))}
    </row>
  );

  const swatches = () => (
    <row gap={3} y="center" role="radiogroup" label="Colour">
      {FILLS.map((entry, index) => (
        <row
          key={entry.name}
          width={20}
          height={18}
          borderRadius={4}
          cursor="pointer"
          backgroundColor={entry.fill}
          borderColor={fill.pipe(map(is => (is === index ? 'focusRing' : 'controlBorder')))}
          borderWidth={fill.pipe(map(is => (is === index ? 2 : 1)))}
          role="radio"
          label={entry.name}
          states={fill.pipe(map((is): readonly UiSemanticState[] => (is === index ? ['checked'] : [])))}
          onClick={() => (fill.value = index)}
        />
      ))}
    </row>
  );

  const toggle = (label: string, held: { value: boolean } & Observable<boolean>) => (
    <row
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
    </row>
  );

  const button = (label: string, onClick: () => void) => (
    <button
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

  const values = test.pipe(map(id => TESTS.find(entry => entry.id === id)?.values ?? 0));

  return (
    <row
      width={percent(100)}
      flexShrink={0}
      y="center"
      gap={8}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      backgroundColor="surface"
      role="form"
      label="Rules for the selection">
      <row gap={2} y="center" role="radiogroup" label="Rule kind">
        {(['format', 'validation'] as const).map(which => (
          <row
            key={which}
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
          </row>
        ))}
      </row>
      {inputs.tab.pipe(
        map(which =>
          which === 'format'
            ? [
                <row key="format" gap={8} y="center">
                  {choice('Condition', TESTS, test)}
                  {values.pipe(
                    map(count =>
                      count === 0
                        ? []
                        : [
                            field('Value', first, applyFormat, 110, inputs.ref.value ?? undefined),
                            ...(count > 1 ? [field('And', second, applyFormat, 80)] : [])
                          ]
                    )
                  )}
                  {swatches()}
                  {button('Apply', applyFormat)}
                </row>
              ]
            : [
                <row key="validation" gap={8} y="center">
                  {choice('Allow', CHECKS, check)}
                  {field('Allowed values', allowed, applyValidation, 200, inputs.ref.value ?? undefined)}
                  {toggle('Refuse anything else', strict)}
                  {button('Apply', applyValidation)}
                </row>
              ]
        )
      )}
      <box flex={1} minWidth={0} />
      {button('Clear rules', () => {
        sheet.send.clearRules();
        close();
      })}
      {button('Close', close)}
    </row>
  );
}
