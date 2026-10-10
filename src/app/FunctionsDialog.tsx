import { combineLatest, map } from 'rxjs';

import { Column, Text, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { SIGNATURES, type Signature } from '../sheet/Signatures';

/**
 * Help ▸ Functions and Insert ▸ Function: every function the sheet
 * has, with what it takes and what it does — Phase 41.
 *
 * The formula editor already knew all of this and told it one function
 * at a time, once somebody had typed the name. A person who does not
 * know the name had nowhere to look. The table is the editor's own,
 * `SIGNATURES`, so the list and the hint cannot disagree.
 *
 * From the Insert menu each row has an Insert button that opens the
 * active cell with `=NAME(` — or adds `NAME(` to a formula being
 * typed — so the hint takes over from there.
 */
export interface FunctionsDialogProps {
  readonly open: boolean;
  /** Whether choosing a function puts it in the cell. */
  readonly inserting: boolean;
  readonly onInsert: (name: string) => void;
  readonly onClose: () => void;
}

export interface FunctionEntry {
  readonly name: string;
  /** `SUMIF(range, criteria, [sum_range])`, as the hint writes it. */
  readonly call: string;
  readonly summary: string;
}

/** Every function, as the reference lists it: alphabetical. */
export const FUNCTION_ENTRIES: readonly FunctionEntry[] = Object.entries(SIGNATURES)
  .map(([name, signature]: [string, Signature]) => ({ name, call: `${name}(${signature.args.join(', ')})`, summary: signature.summary }))
  .sort((a, b) => a.name.localeCompare(b.name));

/**
 * The functions a search finds: names that start with it first, then
 * names that hold it, then the ones whose description does.
 */
export function functionsMatching(search: string): readonly FunctionEntry[] {
  const words = search.trim().toLowerCase();
  if (words === '') {
    return FUNCTION_ENTRIES;
  }
  const upper = words.toUpperCase();
  const starts = FUNCTION_ENTRIES.filter(entry => entry.name.startsWith(upper));
  const holds = FUNCTION_ENTRIES.filter(entry => !entry.name.startsWith(upper) && entry.name.includes(upper));
  const says = FUNCTION_ENTRIES.filter(entry => !entry.name.includes(upper) && entry.summary.toLowerCase().includes(words));
  return [...starts, ...holds, ...says];
}

export function FunctionsDialog(inputs: Inputs<FunctionsDialogProps>, ctx: ComponentContext) {
  const focus = ctx.inject(FocusService);
  const search = internalState('');

  ctx.effect(inputs.open, open => {
    if (open) {
      search.value = '';
    }
  });

  const button = (key: string, label: string, name: string, onClick: () => void) => (
    <button
      key={key}
      onClick={onClick}
      label={name}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={3}
      paddingBottom={3}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );

  const list = combineLatest([search, inputs.inserting]).pipe(
    map(([words, inserting]) => {
      const found = functionsMatching(words);
      if (found.length === 0) {
        return [<text key="none" text={`No function is called or described as “${words}”.`} fontSize={12} color="textMuted" textWrap="word" />];
      }
      return found.map(entry => (
        <row key={`f-${entry.name}`} gap={8} y="center" minWidth={0} padding={3} role="listitem" label={`${entry.call}: ${entry.summary}`}>
          <column gap={1} flex={1} minWidth={0}>
            <text text={entry.call} fontSize={12} fontFamily="monospace" fontWeight="bold" color="text" textWrap="none" />
            <text text={entry.summary} fontSize={11} color="textMuted" textWrap="word" />
          </column>
          {inserting ? button(`insert-${entry.name}`, 'Insert', `Insert ${entry.name}`, () => inputs.onInsert.value(entry.name)) : null}
        </row>
      ));
    })
  );

  const count = search.pipe(
    map(words => {
      const found = functionsMatching(words).length;
      return words.trim() === '' ? `${found} functions` : `${found} of ${FUNCTION_ENTRIES.length} functions`;
    })
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title={inputs.inserting.pipe(map(inserting => (inserting ? 'Insert a function' : 'Functions')))}
      width={600}
      content={Column(
        { gap: 10, minWidth: 0 },
        <row gap={8} y="center">
          <editabletext
            value={search}
            ref={(node: UiNode | null) => {
              if (node !== null && inputs.open.value) {
                queueMicrotask(() => focus.focus(node));
              }
            }}
            flex={1}
            minWidth={0}
            fontSize={13}
            color="text"
            textWrap="none"
            backgroundColor="background"
            borderColor="border"
            borderWidth={1}
            padding={5}
            role="searchbox"
            label="Search functions"
            placeholder="Search by name or what it does"
            onInput={(event: UiTextChangeEvent) => (search.value = event.value)}
          />
          <text text={count} fontSize={11} color="textMuted" textWrap="none" selectable={false} />
        </row>,
        <column gap={2} minWidth={0} maxHeight={360} overflow="auto" role="list" label="Functions">
          {list}
        </column>,
        Text({
          text: 'Square brackets mark an argument that can be left out, and … one that can repeat.',
          fontSize: 11,
          color: 'textMuted',
          textWrap: 'word'
        }),
        <row gap={8} x="end">
          {button('close', 'Close', 'Close', () => inputs.onClose.value())}
        </row>
      )}
    />
  );
}
