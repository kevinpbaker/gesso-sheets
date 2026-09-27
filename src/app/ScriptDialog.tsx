import { combineLatest, map } from 'rxjs';

import { Column, editorFor, percent, Text, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { scriptSpans } from './FormulaColours';
import type { SheetScripts } from './SheetContract';

/**
 * Data ▸ Scripts: the workbook's scripts, one at a time.
 *
 * A text field with the formula editor's colouring and nothing more,
 * as the roadmap has it — the references in a script's strings in the
 * colours a formula would give them — plus a name, the run's outcome
 * and what it logged. Ctrl+Enter runs, as it saves a note.
 *
 * A script that came with a file says so under the field, and Run
 * asks first, each time: the host refuses a file's script that nobody
 * confirmed, and this is where somebody confirms it. Running saves
 * first, so what runs is what is on the screen.
 */
export interface ScriptDialogProps {
  readonly open: boolean;
  readonly scripts: SheetScripts;
  readonly onSave: (was: string, name: string, source: string, kind: 'run' | 'functions') => void;
  readonly onRemove: (name: string) => void;
  readonly onRun: (name: string, confirmed: boolean) => void;
  readonly onStop: () => void;
  readonly onClose: () => void;
}

/** What a new script starts as: enough to show the shape of the API. */
export const TEMPLATE = `// Doubles A2:A10 into B2:B10, and makes the result bold.
const rows = sheet.range("A2:A10").values;
sheet.range("B2:B10").write(rows.map(([value]) => [typeof value === "number" ? value * 2 : null]));
sheet.range("B2:B10").format({ bold: true });
`;

/** What a new functions script starts as: one function, and how a formula calls it. */
export const FUNCTIONS_TEMPLATE = `// A formula calls these by name: =TAX(B2, 0.2)
// A function sees its arguments and nothing else. A range arrives as rows.
function TAX(amount, rate) {
  return Math.round(amount * rate * 100) / 100;
}
`;

export function ScriptDialog(inputs: Inputs<ScriptDialogProps>, ctx: ComponentContext) {
  const focus = ctx.inject(FocusService);
  /** The script being edited, by the name it was saved under; empty for one not saved yet. */
  const selected = internalState('');
  const draftName = internalState('');
  const draftSource = internalState('');
  /** Run was pressed on a file's script, and the question is showing. */
  const confirming = internalState(false);
  /** Whether the script being edited is run, or holds functions formulas call. */
  const draftKind = internalState<'run' | 'functions'>('run');

  const scripts = (): SheetScripts => inputs.scripts.value;
  const fromOf = (name: string): string => scripts().entries.find(entry => entry.name === name)?.from ?? '';

  const choose = (name: string, kind: 'run' | 'functions' = 'run'): void => {
    const entry = scripts().entries.find(each => each.name === name);
    const chosenKind = entry?.kind ?? kind;
    draftKind.value = chosenKind;
    selected.value = entry?.name ?? '';
    draftName.value = entry?.name ?? nextName(chosenKind === 'functions' ? 'Functions' : 'Script');
    draftSource.value = entry?.source ?? (chosenKind === 'functions' ? FUNCTIONS_TEMPLATE : TEMPLATE);
    confirming.value = false;
  };
  const nextName = (stem: string): string => {
    const taken = new Set(scripts().entries.map(entry => entry.name.toUpperCase()));
    let number = 1;
    while (taken.has(`${stem} ${number}`.toUpperCase())) {
      number++;
    }
    return `${stem} ${number}`;
  };

  // Each time it opens: the first script, or a new one when there are none.
  ctx.effect(inputs.open, open => {
    if (open) {
      choose(scripts().entries[0]?.name ?? '');
    }
  });
  // A save that renamed the script, or made it, is still the one being edited.
  ctx.effect(inputs.scripts, next => {
    if (next.entries.some(entry => entry.name === selected.value)) {
      return;
    }
    if (next.entries.some(entry => entry.name === draftName.value.trim())) {
      selected.value = draftName.value.trim();
    }
  });

  const save = (): void => inputs.onSave.value(selected.value, draftName.value, draftSource.value, draftKind.value);
  const run = (confirmed: boolean): void => {
    save();
    confirming.value = false;
    inputs.onRun.value(draftName.value.trim(), confirmed);
  };
  const askToRun = (): void => {
    // Functions are not run: Ctrl+Enter puts them to work by saving them.
    if (draftKind.value === 'functions') {
      save();
      return;
    }
    if (scripts().running !== '') {
      return;
    }
    if (fromOf(selected.value) !== '') {
      confirming.value = true;
      return;
    }
    run(false);
  };

  const onKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter' && (event.modifiers.ctrl === true || event.modifiers.meta === true)) {
      event.preventDefault();
      event.stopPropagation();
      askToRun();
    }
  };

  /**
   * Tab indents, as it does in any editor of code; Shift+Tab still
   * moves the keyboard back out, so the field is never a trap, and the
   * line under it says both.
   */
  const onSourceKey = (event: UiKeyboardEvent): void => {
    const plain = event.modifiers.ctrl !== true && event.modifiers.meta !== true && event.modifiers.alt !== true;
    if (event.key === 'Tab' && plain && event.modifiers.shift !== true && sourceNode !== null) {
      event.preventDefault();
      event.stopPropagation();
      const model = editorFor(sourceNode);
      model.insertText('  ');
      draftSource.value = model.text;
      return;
    }
    onKey(event);
  };
  let sourceNode: UiNode | null = null;

  const button = (key: string, label: string, onClick: () => void, chosen = false) => (
    <button
      key={key}
      onClick={onClick}
      label={label}
      paddingLeft={12}
      paddingRight={12}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor={chosen ? 'primary' : 'controlBackground'}
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} textWrap="none" color={chosen ? 'primaryForeground' : 'controlForeground'} selectable={false} />
    </button>
  );

  const list = combineLatest([inputs.scripts, selected, draftKind]).pipe(
    map(([next, chosen, kind]) => [
      ...next.entries.map(entry => button(`script-${entry.name}`, entry.name, () => choose(entry.name), entry.name === chosen)),
      button('new', 'New script', () => choose(''), chosen === '' && kind === 'run'),
      button('new-functions', 'New functions', () => choose('', 'functions'), chosen === '' && kind === 'functions')
    ])
  );

  const origin = combineLatest([inputs.scripts, selected, draftKind]).pipe(
    map(([, chosen, kind]) => {
      const from = fromOf(chosen);
      if (kind === 'functions') {
        return from === ''
          ? 'Functions a formula calls by name. Each sees only its arguments, and reaches nothing else.'
          : `Came with ${from}. Its functions are off, so a formula that calls one says #NAME?.`;
      }
      return from === ''
        ? 'Written here. It runs when you run it, on a worker of its own, and reaches this workbook and nothing else.'
        : `Came with ${from}. It runs only when you choose to, and asks each time.`;
    })
  );

  const question = combineLatest([confirming, selected]).pipe(
    map(([asking, chosen]) =>
      asking
        ? [
            <row key="confirm" gap={8} y="center" padding={8} borderRadius={6} backgroundColor="surface">
              <text
                text={`This script came with ${fromOf(chosen)}. Running it lets it change this workbook. Run it?`}
                flex={1}
                minWidth={0}
                fontSize={12}
                color="text"
                textWrap="word"
              />
              {button('no', 'Cancel', () => (confirming.value = false))}
              {button('yes', 'Run it', () => run(true))}
            </row>
          ]
        : []
    )
  );

  /** How the last run of this script went, and what it logged; or why a save was refused. */
  const outcome = combineLatest([inputs.scripts, selected, draftKind]).pipe(
    map(([next, chosen, kind]) => {
      const lines: string[] = [];
      const entry = next.entries.find(each => each.name === chosen);
      if (next.refused !== '') {
        lines.push(next.refused);
      } else if (kind === 'functions') {
        if (entry !== undefined && entry.from === '') {
          lines.push(entry.defines.length === 0 ? 'Defines no function a formula can call yet.' : `Defines ${entry.defines.join(', ')}.`);
        }
        if (entry !== undefined && entry.problem !== '') {
          lines.push(entry.problem);
        }
      } else if (next.running !== '') {
        lines.push(`${next.running} is running…`);
      } else if (next.last !== null && next.last.name === (chosen === '' ? draftName.value.trim() : chosen)) {
        lines.push(next.last.text);
      }
      const log = next.running === '' && next.last !== null && next.last.name === chosen ? next.last.log.slice(-8) : [];
      return [
        ...lines.map((line, index) => (
          <text key={`line-${index}`} text={line} fontSize={12} color="text" textWrap="word" selectable={true} />
        )),
        ...log.map((line, index) => (
          <text
            key={`log-${index}`}
            text={line}
            fontSize={11}
            fontFamily="monospace"
            color="textMuted"
            textWrap="word"
            selectable={true}
          />
        ))
      ];
    })
  );

  const actions = combineLatest([inputs.scripts, selected, draftKind]).pipe(
    map(([next, chosen]) => [
      ...(chosen === '' ? [] : [button('remove', 'Delete', () => {
        inputs.onRemove.value(chosen);
        choose(scripts().entries.find(entry => entry.name !== chosen)?.name ?? '');
      })]),
      <box key="gap" flex={1} minWidth={0} />,
      button('close', 'Close', () => inputs.onClose.value()),
      button('save', 'Save', save),
      ...(draftKind.value === 'functions'
        ? []
        : [next.running === '' ? button('run', 'Run', askToRun) : button('stop', 'Stop', () => inputs.onStop.value())])
    ])
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Scripts"
      width={680}
      content={Column(
        { gap: 10, minWidth: 0 },
        <row gap={6} y="center">
          {list}
        </row>,
        <row gap={8} y="center">
          <text text="Name" fontSize={12} color="textMuted" selectable={false} />
          <editabletext
            value={draftName}
            width={240}
            fontSize={12}
            color="text"
            textWrap="none"
            backgroundColor="background"
            borderColor="border"
            borderWidth={1}
            padding={4}
            role="textbox"
            label="Script name"
            onInput={(event: UiTextChangeEvent) => (draftName.value = event.value)}
            onKeyDown={onKey}
          />
        </row>,
        <editabletext
          value={draftSource}
          spans={draftSource.pipe(map(scriptSpans))}
          multiline={true}
          ref={(node: UiNode | null) => {
            // After the dialog has placed the keyboard on the first
            // thing it holds, which here is a script's name in the list.
            sourceNode = node;
            if (node !== null) {
              queueMicrotask(() => focus.focus(node));
            }
          }}
          width={percent(100)}
          minWidth={0}
          minHeight={260}
          fontSize={12}
          fontFamily="monospace"
          color="text"
          // Wrapped rather than scrolled sideways: a line longer than the
          // dialog is still read whole, and the dialog stays its width.
          textWrap="word"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={6}
          role="textbox"
          label="Script"
          onInput={(event: UiTextChangeEvent) => (draftSource.value = event.value)}
          onKeyDown={onSourceKey}
        />,
        Text({
          text: draftKind.pipe(map(kind => `Tab indents and Shift+Tab leaves. Ctrl+Enter ${kind === 'functions' ? 'saves' : 'runs'}.`)),
          fontSize: 11,
          color: 'textMuted'
        }),
        Text({ text: origin, fontSize: 11, color: 'textMuted', textWrap: 'word' }),
        <column gap={4}>{question}</column>,
        <column gap={4}>{outcome}</column>,
        <row gap={8} y="center">
          {actions}
        </row>
      )}
    />
  );
}
