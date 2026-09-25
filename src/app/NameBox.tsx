import { combineLatest, map, type Observable } from 'rxjs';

import { editorFor, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName, parseAddress, relativeRef, type RangeRef } from '../sheet/A1';
import type { SheetName } from './SheetContract';
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

  /** The names, kept to hand so a commit can resolve one. */
  let known: readonly SheetName[] = [];
  _ctx.effect(edit.names, view => (known = view.entries));

  /**
   * The range a name stands for, as an address would give one.
   *
   * Case-insensitive, because the sheet's names are: somebody who
   * typed `sales` meant `Sales`, and a name box that disagreed would
   * define a second name rather than going to the first.
   */
  const namedRange = (typed: string): RangeRef | null => {
    const found = known.find(entry => entry.name.toUpperCase() === typed.trim().toUpperCase());
    if (found === undefined) {
      return null;
    }
    return {
      start: relativeRef(found.firstRow, found.firstColumn),
      end: relativeRef(found.lastRow, found.lastColumn)
    };
  };

  const commit = (): void => {
    const typed = draft.value;
    draft.value = null;
    if (typed === null) {
      return;
    }
    const range = parseAddress(typed) ?? namedRange(typed);
    if (range === null) {
      /**
       * Not an address and not a name the sheet knows — so it is
       * somebody naming the selection.
       *
       * This is the whole of `Insert ▸ Name` in one gesture: pick a
       * range, type what it is, press Enter. The rules live on the
       * other thread and the answer comes back on the names view, so
       * nothing here has to know what a legal name looks like.
       *
       * A single cell is not named. Naming one is legal and almost
       * always a slip — the selection was a range a moment ago — and
       * a name box that silently defined `Sales` as `B7` would be
       * worse than one that did nothing.
       */
      const at = edit.selectionNow();
      if (at.row !== at.anchorRow || at.column !== at.anchorColumn) {
        edit.defineName(typed);
      }
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

  /**
   * The node, so focus can select what is in it.
   *
   * The parent's own focus path already selects all when it *asks*
   * for the box — a command that jumps here should not make somebody
   * delete `B7` before typing. Tabbing in went the other way and left
   * the caret at the start, so typing `C9` gave `C9B7`, which is not
   * an address and not a name and does nothing at all.
   */
  let node: UiNode | null = null;

  const onFocus = (): void => {
    if (node !== null) {
      editorFor(node).selectAll();
    }
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
      ref={found => {
        node = found;
        inputs.ref.value?.(found);
      }}
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
      onFocus={onFocus}
    />
  );
}
