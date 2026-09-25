import { combineLatest, map, type Observable } from 'rxjs';

import { type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName, parseAddress } from '../sheet/A1';
import type { SheetEditing } from './SheetEditing';

/**
 * The box to the left of the formula bar: where you are, and where
 * you would like to be.
 *
 * It reads as a label and behaves as a field, which is the whole
 * trick. Most of the time it shows `B7` and is ignored; typed into
 * and committed, it jumps — and typing an address into it is how
 * people reach `A9999` without scrolling to it, which is the one
 * navigation a keyboard cannot otherwise do.
 *
 * Nothing it does crosses the barrier as a command of its own: a jump
 * is a selection, and the selection already has a path. Sending an
 * address to the application worker and waiting to be told where it
 * put the selection would make a jump a round trip, and Phase 4
 * settled that the selection leads on this side.
 */
export interface NameBoxProps {
  readonly editing: SheetEditing;
  /** Receives the node, so a command can put the keyboard here. */
  readonly ref?: (node: UiNode | null) => void;
}

export function NameBox(inputs: Inputs<NameBoxProps>, _ctx: ComponentContext) {
  const edit = inputs.editing.value;
  /**
   * What is being typed, or null when nobody is typing.
   *
   * The same shape as the cell editor's draft and for the same
   * reason: while somebody is typing an address the box must show
   * what they typed and not where they still are, and Escape has to
   * put back the address rather than a half-typed one.
   */
  const draft = internalState<string | null>(null);

  /** Where the selection is, as a person would write it. */
  const address: Observable<string> = edit.selection.pipe(
    map(at =>
      at.row === at.anchorRow && at.column === at.anchorColumn
        ? `${columnName(at.column)}${at.row + 1}`
        : `${columnName(Math.min(at.column, at.anchorColumn))}${Math.min(at.row, at.anchorRow) + 1}:${columnName(
            Math.max(at.column, at.anchorColumn)
          )}${Math.max(at.row, at.anchorRow) + 1}`
    )
  );

  const shown = combineLatest([draft, address]).pipe(map(([typed, where]) => typed ?? where));

  const commit = (): void => {
    const typed = draft.value;
    draft.value = null;
    if (typed === null) {
      return;
    }
    const range = parseAddress(typed);
    if (range === null) {
      // Not an address. Put the selection's own address back rather
      // than jumping somewhere invented, and say nothing: a name box
      // that argued would be a name box people stop using.
      return;
    }
    // The anchor goes on the *far* corner and the active cell is
    // extended back to the near one, because `extendTo` moves the
    // active corner and leaves the anchor. Somebody who asked for
    // `A1:C9` expects to start typing in A1; done the other way round
    // the same nine cells are selected with the cursor in the wrong
    // corner of them.
    edit.moveTo(range.end.row, range.end.column);
    if (range.end.row !== range.start.row || range.end.column !== range.start.column) {
      edit.extendTo(range.start.row, range.start.column);
    }
    edit.focusSheet();
  };

  const onKeyDown = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      commit();
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      draft.value = null;
      edit.focusSheet();
    }
  };

  return (
    <editabletext
      ref={inputs.ref.value ?? undefined}
      value={shown}
      width={92}
      fontSize={12}
      fontWeight={600}
      color="text"
      textWrap="none"
      textAlign="center"
      verticalAlign="middle"
      backgroundColor="background"
      borderColor="border"
      borderWidth={1}
      padding={4}
      role="textbox"
      label="Name box"
      onInput={(event: UiTextChangeEvent) => (draft.value = event.value)}
      onKeyDown={onKeyDown}
    />
  );
}
