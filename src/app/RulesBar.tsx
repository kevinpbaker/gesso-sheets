import { combineLatest, map, type Observable } from 'rxjs';

import { percent, type UiKeyboardEvent, type UiNode, type UiSemanticState, type UiTextChangeEvent } from 'gesso-core';
import { Select } from 'gesso-components';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName, parseAddress } from '../sheet/A1';
import type { ColourScale, ConditionalPaint, ConditionalTest } from '../sheet/Conditional';
import type { ValidationRule } from '../sheet/Validation';
import {
  Sheet,
  type SheetConditionalRule,
  type SheetListedConditional,
  type SheetListedValidation,
  type SheetRules
} from './SheetContract';
import type { SheetEditing } from './SheetEditing';
import { colourName } from './colourNames';
import { rectOf } from './SheetRanges';

export { colourName };

/**
 * Formats that think, and what a cell is allowed to hold, as one bar
 * that reads as a sentence.
 *
 * It used to be a row of controls — eight buttons for the test, five
 * unlabelled swatches, a field, Apply — that never said which cells it
 * was about, and it had no way to see the rules a sheet already had.
 * So it is laid out as the rule reads aloud:
 *
 *   Highlight  cells E4:E27  where the value  [is greater than]  300  in  [Red fill]   7 of 24 cells match   Add rule
 *
 * with the cells written in the bar and editable, the test and colour
 * as two dropdowns with their words, a count of what the rule would
 * colour as it is written, and under it a line that says what the
 * choice means when it needs saying. The sheet's rules are listed on
 * the line below, each as a sentence, each to change or remove on its
 * own.
 *
 * A row in the flow rather than a floating panel, for the reason the
 * find bar is one: opening it must not cover the cells the rule is
 * about. The rule crosses as data, because a `postMessage` carries no
 * closures.
 */

export type RulesTab = 'format' | 'validation';

export interface RulesBarProps {
  readonly editing: SheetEditing;
  readonly tab: RulesTab;
  readonly onTab: (tab: RulesTab) => void;
  readonly onClose: () => void;
  readonly ref?: (node: UiNode | null) => void;
}

/** What the bar makes: a highlight, a colour scale, or a limit on what a cell may hold. */
type Kind = 'highlight' | 'scale' | 'limit';

/** The colours a highlight can paint, named for what they do. */
export const FILLS: readonly { readonly id: string; readonly name: string; readonly fill: string; readonly color: string }[] = [
  { id: 'red', name: 'Red', fill: '#fce8e6', color: '#c5221f' },
  { id: 'yellow', name: 'Yellow', fill: '#fef7e0', color: '#b06000' },
  { id: 'green', name: 'Green', fill: '#e6f4ea', color: '#137333' },
  { id: 'blue', name: 'Blue', fill: '#e8f0fe', color: '#1967d2' },
  { id: 'grey', name: 'Grey', fill: '#f1f3f4', color: '#5f6368' }
];

/** The colour choice that keeps the rule's own colours, when they are none of `FILLS`. */
const AS_IT_IS = 'current';

/** What a highlight can ask of a cell, as the sentence says it. */
const TESTS: readonly { readonly id: ConditionalTest['kind']; readonly label: string; readonly values: number }[] = [
  { id: 'greaterThan', label: 'is greater than', values: 1 },
  { id: 'lessThan', label: 'is less than', values: 1 },
  { id: 'between', label: 'is between', values: 2 },
  { id: 'equalTo', label: 'is equal to', values: 1 },
  { id: 'textContains', label: 'contains the text', values: 1 },
  { id: 'notEmpty', label: 'is not empty', values: 0 },
  { id: 'isEmpty', label: 'is empty', values: 0 },
  { id: 'formula', label: 'makes a formula true', values: 1 }
];

const CHECKS: readonly { readonly id: ValidationRule['kind']; readonly label: string }[] = [
  { id: 'list', label: 'one of a list' },
  { id: 'number', label: 'a number' },
  { id: 'date', label: 'a date' },
  { id: 'text', label: 'text' }
];

/** The name of a paint, in the words the bar lists it with: "red fill", "bold red text". */
export function paintName(paint: ConditionalPaint | undefined): string {
  if (paint === undefined) {
    return 'no colour';
  }
  const style = [paint.bold === true ? 'bold' : '', paint.italic === true ? 'italic' : ''].filter(Boolean).join(' ');
  const styled = (text: string) => (style === '' ? text : `${style} ${text}`);
  const named = FILLS.find(entry => entry.fill === paint.fill?.toLowerCase() && entry.color === paint.color?.toLowerCase());
  if (named !== undefined) {
    return styled(`${named.name.toLowerCase()} fill`);
  }
  if (paint.fill !== undefined && paint.color !== undefined) {
    return styled(`${colourName(paint.fill)} fill with ${colourName(paint.color)} text`);
  }
  if (paint.fill !== undefined) {
    return styled(`${colourName(paint.fill)} fill`);
  }
  if (paint.color !== undefined) {
    return styled(`${colourName(paint.color)} text`);
  }
  return style === '' ? 'no colour' : `${style} text`;
}

/** A test as the list says it: "greater than 300". */
export function testName(test: ConditionalTest | null): string {
  if (test === null) {
    return 'no test';
  }
  switch (test.kind) {
    case 'greaterThan':
      return `greater than ${test.value}`;
    case 'lessThan':
      return `less than ${test.value}`;
    case 'between':
      return `between ${test.low} and ${test.high}`;
    case 'equalTo':
      return typeof test.value === 'number' ? `equal to ${test.value}` : `equal to “${test.value}”`;
    case 'textContains':
      return `containing “${test.text}”`;
    case 'notEmpty':
      return 'not empty';
    case 'isEmpty':
      return 'empty';
    case 'formula':
      return `where ${test.input}`;
  }
}

/** A scale as the list says it: "white to green", or its own stops. */
function scaleName(scale: ColourScale): string {
  return [scale.from, scale.middle, scale.to]
    .filter((colour): colour is string => colour !== undefined)
    .map(colourName)
    .join(' to ');
}

/** One conditional rule, as a sentence for the list. */
export function describeConditional(rule: SheetListedConditional): string {
  if (rule.scale !== undefined) {
    return `Colour scale · ${rule.range}: shaded ${scaleName(rule.scale)}, lowest to highest`;
  }
  return `Highlight · ${rule.range}: ${testName(rule.test)}, in ${paintName(rule.paint)}`;
}

/** One validation, as a sentence for the list. */
export function describeValidation(rule: SheetListedValidation): string {
  const what = ((): string => {
    switch (rule.rule.kind) {
      case 'list':
        return `one of ${rule.rule.values.join(', ')}`;
      case 'number': {
        const { min, max } = rule.rule;
        return min !== undefined && max !== undefined
          ? `a number from ${min} to ${max}`
          : min !== undefined
            ? `a number of at least ${min}`
            : max !== undefined
              ? `a number of at most ${max}`
              : 'a number';
      }
      case 'date':
        return 'a date';
      case 'text':
        return rule.rule.maxLength === undefined ? 'text' : `text of at most ${rule.rule.maxLength} characters`;
    }
  })();
  return `Limit · ${rule.range}: ${what}${rule.strict ? ', and nothing else' : ', marked when it is not'}`;
}

/** The top left cell of a range as written, for the formula hint. */
function firstCellOf(range: string): string {
  const parsed = parseAddress(range.trim().replace(/\$/g, ''));
  if (parsed === null) {
    return 'the first cell';
  }
  return `${columnName(Math.min(parsed.start.column, parsed.end.column))}${Math.min(parsed.start.row, parsed.end.row) + 1}`;
}

export function RulesBar(inputs: Inputs<RulesBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;

  const kind = internalState<Kind>('highlight');
  /** The cells, as written in the bar; follows the selection until somebody types in it. */
  const range = internalState('');
  let rangeTyped = false;
  const test = internalState<ConditionalTest['kind']>('greaterThan');
  const first = internalState('');
  const second = internalState('');
  const colour = internalState(FILLS[0].id);
  /**
   * The colours of the rule being changed, when they are none of the
   * named ones — the seed's pink-to-green scale, bold red text. Kept
   * whole so saving it unchanged puts back what was there.
   */
  const kept = internalState<Pick<SheetConditionalRule, 'paint' | 'scale'> | null>(null);
  const check = internalState<ValidationRule['kind']>('list');
  const allowed = internalState('');
  const strict = internalState(false);
  /** Which rule the bar is changing, as an index into the sheet's list; null for a new one. */
  const editing = internalState<number | null>(null);
  const listOpen = internalState(false);

  // The tab the chrome asked for is where the bar opens.
  ctx.effect(inputs.tab, tab => {
    if (tab === 'validation') {
      kind.value = 'limit';
    } else if (kind.value === 'limit') {
      kind.value = 'highlight';
    }
  });

  const rangeOfSelection = (): string => {
    const rect = rectOf(edit.selectionNow());
    const top = `${columnName(rect.firstColumn)}${rect.firstRow + 1}`;
    const bottom = `${columnName(rect.lastColumn)}${rect.lastRow + 1}`;
    return top === bottom ? top : `${top}:${bottom}`;
  };
  ctx.effect(edit.selection, () => {
    if (!rangeTyped && editing.value === null) {
      range.value = rangeOfSelection();
    }
  });

  /**
   * The bar opens on the rule the active cell is already under, to be
   * changed rather than added to — and only when that rule changes: the
   * key moves on every arrow key, and refilling the fields each time
   * would throw away what somebody typed.
   */
  let loadedFormat = 'null';
  let loadedCheck = 'null';
  ctx.effect(sheet.view.activeRules, active => {
    const format = JSON.stringify([active.conditional, active.conditionalAt]);
    if (format !== loadedFormat && kind.value !== 'limit') {
      const had = loadedFormat !== 'null' && loadedFormat !== JSON.stringify([null, -1]) && loadedFormat !== JSON.stringify([null, undefined]);
      loadedFormat = format;
      const listed = active.conditionalAt === undefined ? undefined : sheet.view.rules.value.conditional[active.conditionalAt];
      if (active.conditional !== null && listed !== undefined) {
        loadConditional(listed, active.conditionalAt ?? null);
      } else if (had) {
        startNew();
      }
    }
    const validation = JSON.stringify([active.validation, active.validationAt]);
    if (validation !== loadedCheck && kind.value === 'limit') {
      const had = loadedCheck !== 'null' && loadedCheck !== JSON.stringify([null, -1]) && loadedCheck !== JSON.stringify([null, undefined]);
      loadedCheck = validation;
      const listed = active.validationAt === undefined ? undefined : sheet.view.rules.value.validations[active.validationAt];
      if (listed !== undefined) {
        loadValidation(listed, active.validationAt ?? null);
      } else if (had) {
        startNew();
      }
    }
  });

  /**
   * Another kind of rule: the active cell's rule of that kind if it has
   * one, to change, and otherwise a new one over the same cells — never
   * the last kind's rule wearing the new kind's controls, which would
   * save a highlight as a scale.
   */
  function switchTo(which: Kind): void {
    const cells = range.value;
    const typed = rangeTyped;
    const active = sheet.view.activeRules.value;
    const rules = sheet.view.rules.value;
    startNew();
    kind.value = which;
    if (which === 'limit') {
      const listed = active.validationAt === undefined ? undefined : rules.validations[active.validationAt];
      if (listed !== undefined) {
        loadValidation(listed, active.validationAt ?? null);
        return;
      }
    } else {
      const listed = active.conditionalAt === undefined ? undefined : rules.conditional[active.conditionalAt];
      if (listed !== undefined && (listed.scale !== undefined) === (which === 'scale')) {
        loadConditional(listed, active.conditionalAt ?? null);
        return;
      }
      if (which === 'scale') {
        colour.value = 'green';
      }
    }
    range.value = cells;
    rangeTyped = typed;
  }

  function startNew(): void {
    editing.value = null;
    rangeTyped = false;
    range.value = rangeOfSelection();
    test.value = 'greaterThan';
    first.value = '';
    second.value = '';
    colour.value = FILLS[0].id;
    kept.value = null;
    check.value = 'list';
    allowed.value = '';
    strict.value = false;
  }

  function loadConditional(rule: SheetListedConditional, at: number | null): void {
    editing.value = at;
    range.value = rule.range;
    rangeTyped = false;
    if (rule.scale !== undefined) {
      const scale = rule.scale;
      kind.value = 'scale';
      const named = FILLS.find(entry => scale.from.toLowerCase() === '#ffffff' && scale.middle === undefined && entry.color === scale.to.toLowerCase());
      kept.value = named === undefined ? { scale } : null;
      colour.value = named?.id ?? AS_IT_IS;
      return;
    }
    kind.value = 'highlight';
    const held = rule.test;
    if (held !== null) {
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
    }
    const paint = rule.paint ?? {};
    const named = FILLS.find(
      entry => entry.fill === paint.fill?.toLowerCase() && entry.color === paint.color?.toLowerCase() && paint.bold !== true && paint.italic !== true
    );
    kept.value = named === undefined ? { paint } : null;
    colour.value = named?.id ?? AS_IT_IS;
  }

  function loadValidation(rule: SheetListedValidation, at: number | null): void {
    editing.value = at;
    kind.value = 'limit';
    range.value = rule.range;
    rangeTyped = false;
    check.value = rule.rule.kind;
    strict.value = rule.strict;
    switch (rule.rule.kind) {
      case 'list':
        allowed.value = rule.rule.values.join(', ');
        return;
      case 'number':
        allowed.value =
          rule.rule.min === undefined && rule.rule.max === undefined ? '' : `${rule.rule.min ?? ''}, ${rule.rule.max ?? ''}`.replace(/, $/, '');
        return;
      case 'text':
        allowed.value = rule.rule.maxLength === undefined ? '' : String(rule.rule.maxLength);
        return;
      default:
        allowed.value = '';
    }
  }

  const close = (): void => {
    inputs.onClose.value();
    edit.focusSheet();
  };

  /** A number, or null when what was typed is not one — zero is a threshold, an empty box is not. */
  const numberOf = (text: string): number | null => {
    const trimmed = text.trim();
    if (trimmed === '' || !/^[-+]?(\d+\.?\d*|\.\d+)$/.test(trimmed)) {
      return null;
    }
    return Number(trimmed);
  };

  const paintOf = (): ConditionalPaint => {
    const own = kept.value?.paint;
    if (colour.value === AS_IT_IS && own !== undefined) {
      return own;
    }
    const chosen = FILLS.find(entry => entry.id === colour.value) ?? FILLS[0];
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
        return one === null || other === null ? null : { kind: 'between', low: Math.min(one, other), high: Math.max(one, other) };
      }
      case 'equalTo':
        return first.value === '' ? null : { kind: 'equalTo', value: one ?? first.value };
      case 'textContains':
        return first.value === '' ? null : { kind: 'textContains', text: first.value };
      case 'notEmpty':
        return { kind: 'notEmpty' };
      case 'isEmpty':
        return { kind: 'isEmpty' };
      case 'formula': {
        const text = first.value.trim();
        return text === '' || text === '=' ? null : { kind: 'formula', input: text.startsWith('=') ? text : `=${text}` };
      }
    }
  };

  /** The rule as the bar would add it, or null while it is not a whole one yet. */
  const draftRule = (): SheetConditionalRule | null => {
    if (kind.value === 'scale') {
      const own = kept.value?.scale;
      const chosen = FILLS.find(entry => entry.id === colour.value) ?? FILLS[2];
      return { test: null, scale: colour.value === AS_IT_IS && own !== undefined ? own : { from: '#ffffff', to: chosen.color } };
    }
    const rule = testOf();
    return rule === null ? null : { test: rule, paint: paintOf() };
  };

  const saveFormat = (): void => {
    const rule = draftRule();
    if (rule === null) {
      return;
    }
    if (editing.value === null) {
      sheet.send.addConditional(rule, range.value);
    } else {
      sheet.send.replaceConditional(editing.value, rule, range.value);
    }
  };

  const ruleOf = (): ValidationRule | null => {
    // Kept by position for a number: `, 10` is "at most ten".
    const positions = allowed.value.split(',').map(part => part.trim());
    const parts = positions.filter(part => part !== '');
    switch (check.value) {
      case 'list':
        return parts.length === 0 ? null : { kind: 'list', values: parts };
      case 'number': {
        const low = numberOf(positions[0] ?? '');
        const high = numberOf(positions[1] ?? '');
        return { kind: 'number', ...(low === null ? {} : { min: low }), ...(high === null ? {} : { max: high }) };
      }
      case 'date':
        return { kind: 'date' };
      case 'text': {
        const length = numberOf(parts[0] ?? '');
        return length === null ? { kind: 'text' } : { kind: 'text', maxLength: length };
      }
    }
  };

  const saveValidation = (): void => {
    const rule = ruleOf();
    if (rule === null) {
      return;
    }
    if (editing.value === null) {
      sheet.send.addValidation(rule, strict.value, '', range.value);
    } else {
      sheet.send.replaceValidation(editing.value, rule, strict.value, range.value);
    }
  };

  const save = (): void => (kind.value === 'limit' ? saveValidation() : saveFormat());

  // A saved rule is the one being edited from then on, so a second Save
  // changes it again rather than adding a copy — and it is found in the
  // list by being the last added, or the one that was replaced.
  let listed = { conditional: 0, validations: 0 };
  ctx.effect(sheet.view.rules, rules => {
    const grew = { conditional: rules.conditional.length > listed.conditional, validations: rules.validations.length > listed.validations };
    listed = { conditional: rules.conditional.length, validations: rules.validations.length };
    if (editing.value === null && grew.conditional && kind.value !== 'limit') {
      editing.value = rules.conditional.length - 1;
    } else if (editing.value === null && grew.validations && kind.value === 'limit') {
      editing.value = rules.validations.length - 1;
    }
    const inList = kind.value === 'limit' ? rules.validations.length : rules.conditional.length;
    if (editing.value !== null && editing.value >= inList) {
      editing.value = null;
    }
  });

  /**
   * How many cells the rule as written would colour, asked of the sheet
   * each time the draft changes; the serial tells the current answer
   * from one to an older draft.
   */
  let serial = 0;
  const askCount = (): void => {
    if (kind.value === 'limit') {
      return;
    }
    const rule = draftRule();
    serial++;
    if (rule !== null) {
      sheet.send.countMatches(serial, rule, range.value);
    }
  };
  ctx.effect(combineLatest([kind, range, test, first, second]), askCount);

  const onFieldKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      save();
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
    width: number,
    onInput?: () => void,
    ref?: (node: UiNode | null) => void,
    monospace = false
  ) => (
    <editabletext
      key={`field-${label}`}
      ref={ref}
      value={value as never}
      width={width}
      flexShrink={0}
      fontSize={12}
      fontFamily={monospace ? 'monospace' : undefined}
      color="text"
      textWrap="none"
      verticalAlign="middle"
      backgroundColor="background"
      borderColor="border"
      borderWidth={1}
      padding={4}
      role="textbox"
      label={label}
      onInput={(event: UiTextChangeEvent) => {
        value.value = event.value;
        onInput?.();
      }}
      onKeyDown={onFieldKey}
    />
  );

  const words = (text: string | Observable<string>, key: string, color = 'textMuted') => (
    <text key={key} text={text} fontSize={12} textWrap="none" color={color} selectable={false} />
  );

  const dropdown = <T extends string>(
    label: string,
    options: readonly { readonly value: T; readonly label: string }[],
    held: { value: T } & Observable<T>,
    width: number
  ) => (
    <Select
      key={`select-${label}`}
      label={label}
      compact={true}
      labelHidden={true}
      width={width}
      options={options}
      value={held}
      onChange={(next: string) => (held.value = next as T)}
    />
  );

  const button = (label: string, onClick: () => void, primary = false, shown = label) => (
    <button
      key={`button-${label}`}
      onClick={onClick}
      label={label}
      paddingLeft={9}
      paddingRight={9}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={5}
      backgroundColor={primary ? 'primary' : 'controlBackground'}
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={shown} fontSize={11} textWrap="none" color={primary ? 'primaryForeground' : 'controlForeground'} selectable={false} />
    </button>
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

  /** The three kinds, as tabs; each is a different sentence. */
  const kinds = (
    <row key="kinds" gap={2} y="center" role="radiogroup" label="Rule kind">
      {(
        [
          ['highlight', 'Highlight'],
          ['scale', 'Colour scale'],
          ['limit', 'Limit entries']
        ] as const
      ).map(([which, label]) => (
        <button
          key={which}
          focusable={false}
          paddingLeft={7}
          paddingRight={7}
          paddingTop={3}
          paddingBottom={3}
          borderRadius={5}
          cursor="pointer"
          backgroundColor={kind.pipe(map(is => (is === which ? 'controlBackgroundHovered' : 'transparent')))}
          borderColor="controlBorder"
          borderWidth={1}
          role="radio"
          label={label}
          states={kind.pipe(map((is): readonly UiSemanticState[] => (is === which ? ['checked'] : [])))}
          onClick={() => {
            if (kind.value === which) {
              return;
            }
            const wasLimit = kind.value === 'limit';
            switchTo(which);
            if (wasLimit !== (which === 'limit')) {
              inputs.onTab.value(which === 'limit' ? 'validation' : 'format');
            }
          }}>
          <text text={label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
        </button>
      ))}
    </row>
  );

  const colourOptions = combineLatest([kind, kept]).pipe(
    map(([which, own]) => [
      ...(own === null
        ? []
        : [{ value: AS_IT_IS, label: own.scale !== undefined ? scaleName(own.scale) : paintName(own.paint) }]),
      ...FILLS.map(entry => ({ value: entry.id, label: which === 'scale' ? `white to ${entry.name.toLowerCase()}` : `${entry.name.toLowerCase()} fill` }))
    ])
  );

  /** "7 of 24 cells match", from the latest answer to this draft. */
  const count = combineLatest([sheet.view.rules, kind, test, first, second]).pipe(
    map(([rules, which]) => {
      const answer = rules.matches;
      if (which === 'limit' || answer === null || answer.serial !== serial || draftRule() === null || answer.cells === null) {
        return '';
      }
      if (which === 'scale') {
        return `${answer.matching} of ${answer.cells} cells shaded`;
      }
      return answer.matching === 0 ? `no cells of ${answer.cells} match` : `${answer.matching} of ${answer.cells} cells match`;
    })
  );

  /** The sentence itself: which cells, and what about them. */
  const sentence = combineLatest([kind, test, colourOptions]).pipe(
    map(([which, id, colours]) => {
      const cells = [words('cells', 'cells'), field('Cells', range, 110, () => (rangeTyped = true))];
      if (which === 'limit') {
        return [
          ...cells,
          words('may hold', 'may'),
          dropdown('Allow', CHECKS.map(entry => ({ value: entry.id, label: entry.label })), check, 120),
          field('Allowed values', allowed, 200, undefined, inputs.ref.value ?? undefined),
          toggle('Refuse anything else', strict)
        ];
      }
      if (which === 'scale') {
        return [...cells, words('shaded from lowest to highest,', 'shaded'), dropdown('Colour', colours, colour, 150)];
      }
      const values = TESTS.find(entry => entry.id === id)?.values ?? 0;
      const formula = id === 'formula';
      return [
        ...cells,
        words(formula ? 'where the cell' : 'where the value', 'where'),
        dropdown('Condition', TESTS.map(entry => ({ value: entry.id, label: entry.label })), test, 170),
        ...(values === 0 ? [] : [field(formula ? 'Rule formula' : 'Value', first, formula ? 220 : 110, undefined, inputs.ref.value ?? undefined, formula)]),
        ...(values > 1 ? [words('and', 'and'), field('And', second, 80)] : []),
        words('in', 'in'),
        dropdown('Colour', colours, colour, 150)
      ];
    })
  );

  /** What the choice means, when it needs saying; and why a save was refused. */
  const hint = combineLatest([kind, test, range, check, sheet.view.rules]).pipe(
    map(([which, id, cells, allow, rules]) => {
      if (rules.refused !== '') {
        return rules.refused;
      }
      if (which === 'highlight' && id === 'formula') {
        return `Write the formula for ${firstCellOf(cells)}, the first cell. It is checked for each cell of ${cells}, moved as a fill would move it — so =${firstCellOf(cells)}>100 highlights every cell over 100.`;
      }
      if (which === 'scale') {
        return 'The lowest number is the palest and the highest the strongest; text and empty cells are left alone.';
      }
      if (which === 'limit') {
        switch (allow) {
          case 'list':
            return 'The choices, separated by commas. The cell offers them as a list.';
          case 'number':
            return 'A lowest and a highest, separated by a comma — either may be left out: ", 10" is at most ten.';
          case 'text':
            return 'The most characters it may hold, or leave it empty for any text.';
          default:
            return 'Anything that reads as a date.';
        }
      }
      return '';
    })
  );

  const actions = combineLatest([editing, sheet.view.rules, kind]).pipe(
    map(([at, rules, which]) => {
      const total = rules.conditional.length + rules.validations.length;
      void total;
      void which;
      return [button(at === null ? 'Add rule' : 'Save changes', save, true), ...(at === null ? [] : [button('New rule', startNew)])];
    })
  );

  /** One rule of the list: its sentence, and what can be done to it. */
  const listRow = (key: string, text: string, onEdit: () => void, onRemove: () => void, chosen: boolean) => (
    <row key={key} gap={8} y="center" minWidth={0} role="listitem" label={text}>
      <text
        text={text}
        flex={1}
        minWidth={0}
        fontSize={12}
        color={chosen ? 'text' : 'textMuted'}
        fontWeight={chosen ? 'bold' : 'normal'}
        textWrap="none"
      />
      {button(`Edit ${text}`, onEdit, false, 'Edit')}
      {button(`Delete ${text}`, onRemove, false, 'Delete')}
    </row>
  );

  const list = combineLatest([listOpen, sheet.view.rules, editing, kind]).pipe(
    map(([open, rules, at, which]: [boolean, SheetRules, number | null, Kind]) => {
      if (!open) {
        return [];
      }
      if (rules.conditional.length + rules.validations.length === 0) {
        return [words('This sheet has no rules yet. Write one above and press Add rule.', 'none')];
      }
      return [
        ...rules.conditional.map((rule, index) =>
          listRow(
            `c-${index}`,
            describeConditional(rule),
            () => loadConditional(rule, index),
            () => {
              sheet.send.removeConditional(index);
              if (at === index && which !== 'limit') {
                startNew();
              }
            },
            which !== 'limit' && at === index
          )
        ),
        ...rules.validations.map((rule, index) =>
          listRow(
            `v-${index}`,
            describeValidation(rule),
            () => loadValidation(rule, index),
            () => {
              sheet.send.removeValidation(index);
              if (at === index && which === 'limit') {
                startNew();
              }
            },
            which === 'limit' && at === index
          )
        ),
        <row key="all" gap={8} y="center">
          <box flex={1} minWidth={0} />
          {button('Remove every rule', () => {
            sheet.send.clearRules();
            startNew();
          })}
        </row>
      ];
    })
  );

  return (
    <column
      width={percent(100)}
      flexShrink={0}
      gap={4}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      backgroundColor="surface"
      role="form"
      label="Rules for the selection">
      <row width={percent(100)} y="center" gap={8} flexWrap="wrap" rowGap={4}>
        {kinds}
        {sentence}
        {Text(count)}
        {actions}
      </row>
      <row width={percent(100)} y="center" gap={8}>
        <text text={hint} flex={1} minWidth={0} fontSize={11} color="textMuted" textWrap="word" />
        {sheet.view.rules.pipe(
          map(rules => {
            const total = rules.conditional.length + rules.validations.length;
            return [
              button(
                'Rules on this sheet',
                () => (listOpen.value = !listOpen.value),
                false,
                total === 0 ? 'No rules on this sheet' : `Rules on this sheet (${total})`
              )
            ];
          })
        )}
        {button('Close', close)}
      </row>
      <column gap={3} width={percent(100)}>
        {list}
      </column>
    </column>
  );

  function Text(text: Observable<string>) {
    return <text text={text} fontSize={11} color="textMuted" textWrap="word" minWidth={0} />;
  }
}
