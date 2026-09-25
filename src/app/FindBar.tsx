import { map, type Observable } from 'rxjs';

import {
  percent,
  type UiKeyboardEvent,
  type UiNode,
  type UiSemanticState,
  type UiTextChangeEvent
} from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { Sheet } from './SheetContract';
import type { SheetEditing } from './SheetEditing';

/**
 * Find, and replace when it is asked for.
 *
 * A row in the flow rather than a floating panel, so that opening it
 * never covers the cell somebody is looking for. It costs a row of
 * height, which is the cheaper of the two mistakes.
 *
 * Every question it asks is answered on the other thread — this file
 * holds the query text and nothing else. `SheetFind` searches the
 * store, the application worker moves the selection, and what comes
 * back over the barrier is a count and a position. The matches
 * themselves never cross: there can be tens of thousands of them and
 * the render worker's whole use for the list is to draw "3 of 412".
 */
export interface FindBarProps {
  readonly editing: SheetEditing;
  /** True when the bar should show the replace half as well. */
  readonly replacing: Observable<boolean>;
  readonly onClose: () => void;
  readonly ref?: (node: UiNode | null) => void;
}

export function FindBar(inputs: Inputs<FindBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;
  const query = internalState('');
  const replacement = internalState('');
  const matchCase = internalState(false);
  const wholeCell = internalState(false);

  /** Re-runs the search. Every control that changes an option calls it. */
  const search = (): void => {
    if (query.value === '') {
      sheet.send.clearFind();
      return;
    }
    sheet.send.find(query.value, matchCase.value, wholeCell.value, true);
  };

  const close = (): void => {
    sheet.send.clearFind();
    inputs.onClose.value();
    edit.focusSheet();
  };

  const onQueryKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      // Enter searches the first time and steps every time after,
      // which is what the key does in every find bar ever built.
      if (sheet.view.find.value.query === query.value && query.value !== '') {
        sheet.send.findStep(event.modifiers.shift !== true);
      } else {
        search();
      }
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      close();
    }
  };

  const count: Observable<string> = sheet.view.find.pipe(
    map(find => {
      if (find.query === '') {
        return '';
      }
      if (find.matches === 0) {
        return 'None';
      }
      return find.active === 0 ? `${find.matches}` : `${find.active}/${find.matches}`;
    })
  );

  const field = (
    label: string,
    value: { value: string },
    onKeyDown: (event: UiKeyboardEvent) => void,
    node?: (n: UiNode | null) => void
  ) => (
    <editabletext
      ref={node}
      value={value as never}
      width={170}
      fontSize={12}
      color="text"
      textWrap="none"
      verticalAlign="middle"
      backgroundColor="background"
      borderColor="border"
      borderWidth={1}
      padding={4}
      role="searchbox"
      label={label}
      onInput={(event: UiTextChangeEvent) => (value.value = event.value)}
      onKeyDown={onKeyDown}
    />
  );

  const toggle = (label: string, state: { value: boolean }, on: Observable<boolean>) => (
    <row
      paddingLeft={8}
      paddingRight={8}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={5}
      cursor="pointer"
      backgroundColor={on.pipe(map(is => (is ? 'controlBackgroundHovered' : 'transparent')))}
      borderColor="controlBorder"
      borderWidth={1}
      role="checkbox"
      label={label}
      states={on.pipe(map((is): readonly UiSemanticState[] => (is ? ['checked'] : [])))}
      onClick={() => {
        state.value = !state.value;
        search();
      }}>
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
      role="search"
      label="Find in sheet">
      {field('Find', query, onQueryKey, inputs.ref.value ?? undefined)}
      <text text={count} width={54} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} />
      {button('Previous', () => sheet.send.findStep(false))}
      {button('Next', () => sheet.send.findStep(true))}
      {toggle('Match case', matchCase, matchCase)}
      {toggle('Whole cell', wholeCell, wholeCell)}
      {inputs.replacing.pipe(
        map(on =>
          on
            ? [
                field('Replace with', replacement, event => {
                  if (event.key === 'Escape') {
                    event.preventDefault();
                    event.stopPropagation();
                    close();
                  }
                }),
                button('Replace', () => sheet.send.replaceOne(replacement.value)),
                button('Replace all', () => sheet.send.replaceAll(replacement.value))
              ]
            : []
        )
      )}
      <box flex={1} minWidth={0} />
      {button('Close', close)}
    </row>
  );
}
