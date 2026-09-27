import { combineLatest, map } from 'rxjs';

import { Column, Text, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { formulaSpans } from './FormulaColours';
import type { SheetNames } from './SheetContract';

/**
 * Insert ▸ Names: every name the workbook has, to read, change and
 * take away.
 *
 * The name box makes a name in one gesture and was, until this, the
 * only way to touch one: a name could be made and never seen again,
 * and a name that holds a `LAMBDA` could not be read at all. So the
 * list is the whole table, ranges and formulas together, with what
 * each one refers to beside it; choosing one puts it in the two fields
 * below, where it is edited as text — a range is written as a range,
 * `=Sales!$A$4:$A$27`, and anything else is a formula.
 */
export interface NamesDialogProps {
  readonly open: boolean;
  readonly names: SheetNames;
  readonly onSave: (was: string, name: string, refersTo: string) => void;
  readonly onRemove: (name: string) => void;
  readonly onClose: () => void;
}

interface Row {
  readonly name: string;
  readonly refersTo: string;
}

/** Ranges and formulas as one list, in the order somebody reads names. */
export function namesListed(names: SheetNames): Row[] {
  return [
    ...names.entries.map(entry => ({ name: entry.name, refersTo: entry.refersTo })),
    ...names.formulas.map(entry => ({ name: entry.name, refersTo: entry.formula }))
  ].sort((a, b) => a.name.localeCompare(b.name));
}

export function NamesDialog(inputs: Inputs<NamesDialogProps>, ctx: ComponentContext) {
  const focus = ctx.inject(FocusService);
  /** The name being edited, as it was saved; empty for a new one. */
  const selected = internalState('');
  const draftName = internalState('');
  const draftRefers = internalState('');
  /** Whether a save has been sent and not yet answered, so its refusal can be shown. */
  const saving = internalState(false);

  const rows = (): Row[] => namesListed(inputs.names.value);

  const choose = (name: string): void => {
    const row = rows().find(each => each.name === name);
    selected.value = row?.name ?? '';
    draftName.value = row?.name ?? '';
    draftRefers.value = row?.refersTo ?? '=';
    saving.value = false;
  };

  ctx.effect(inputs.open, open => {
    if (open) {
      choose(rows()[0]?.name ?? '');
    }
  });
  // A save answered: the name it was saved under is the one being edited.
  ctx.effect(inputs.names, next => {
    if (!saving.value || next.refused !== '') {
      return;
    }
    const saved = namesListed(next).find(row => row.name.toUpperCase() === draftName.value.trim().toUpperCase());
    if (saved !== undefined) {
      selected.value = saved.name;
      saving.value = false;
    }
  });

  const save = (): void => {
    saving.value = true;
    inputs.onSave.value(selected.value, draftName.value, draftRefers.value);
  };

  const onKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      save();
    }
  };

  const button = (key: string, label: string, onClick: () => void, chosen = false) => (
    <button
      key={key}
      onClick={onClick}
      label={label}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={6}
      backgroundColor={chosen ? 'primary' : 'controlBackground'}
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} textWrap="none" color={chosen ? 'primaryForeground' : 'controlForeground'} selectable={false} />
    </button>
  );

  /** One line per name: the name as a button to edit it, and what it holds. */
  const list = combineLatest([inputs.names, selected]).pipe(
    map(([next, chosen]) => {
      const listed = namesListed(next);
      if (listed.length === 0) {
        return [
          <text
            key="none"
            text="No names yet. Name a range by selecting it and typing into the name box, or write one below."
            fontSize={12}
            color="textMuted"
            textWrap="word"
          />
        ];
      }
      return listed.map(row => (
        <row key={`name-${row.name}`} gap={8} y="center" minWidth={0}>
          {button(`pick-${row.name}`, row.name, () => choose(row.name), row.name === chosen)}
          <text text={row.refersTo} flex={1} minWidth={0} fontSize={12} fontFamily="monospace" color="textMuted" textWrap="none" />
        </row>
      ));
    })
  );

  const refused = combineLatest([inputs.names, saving]).pipe(map(([next, waiting]) => (waiting ? next.refused : '')));

  const actions = selected.pipe(
    map(chosen => [
      ...(chosen === ''
        ? []
        : [
            button('remove', 'Delete', () => {
              inputs.onRemove.value(chosen);
              choose('');
            })
          ]),
      button('new', 'New name', () => choose('')),
      <box key="gap" flex={1} minWidth={0} />,
      button('close', 'Close', () => inputs.onClose.value()),
      button('save', 'Save', save)
    ])
  );

  const field = (label: string, value: typeof draftName, spans: boolean, onInput: (text: string) => void, first: boolean) => (
    <row gap={8} y="center">
      <text text={label} width={72} fontSize={12} color="textMuted" selectable={false} />
      <editabletext
        value={value}
        spans={spans ? value.pipe(map(text => formulaSpans(text, undefined))) : undefined}
        flex={1}
        minWidth={0}
        fontSize={12}
        fontFamily={spans ? 'monospace' : undefined}
        color="text"
        textWrap="none"
        backgroundColor="background"
        borderColor="border"
        borderWidth={1}
        padding={4}
        role="textbox"
        label={label}
        ref={
          first
            ? (node: UiNode | null) => {
                if (node !== null) {
                  queueMicrotask(() => focus.focus(node));
                }
              }
            : undefined
        }
        onInput={(event: UiTextChangeEvent) => onInput(event.value)}
        onKeyDown={onKey}
      />
    </row>
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Names"
      width={560}
      content={Column(
        { gap: 10, minWidth: 0 },
        <column gap={4} minWidth={0} maxHeight={240} overflow="auto">
          {list}
        </column>,
        field('Name', draftName, false, text => (draftName.value = text), true),
        field('Refers to', draftRefers, true, text => (draftRefers.value = text), false),
        Text({
          text: 'A range, such as =Sales!$A$4:$A$27, or a formula. =LAMBDA(x, x*2) makes a function called by the name.',
          fontSize: 11,
          color: 'textMuted',
          textWrap: 'word'
        }),
        Text({ text: refused, fontSize: 12, color: 'danger', textWrap: 'word' }),
        <row gap={8} y="center">
          {actions}
        </row>
      )}
    />
  );
}
