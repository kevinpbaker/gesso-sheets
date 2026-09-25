import { BehaviorSubject, combineLatest, map, type Observable } from 'rxjs';

import {
  editorFor,
  percent,
  type UiKeyboardEvent,
  type UiNode,
  type UiSemanticState,
  type UiTextChangeEvent
} from 'gesso-core';
import { Column, Row, Text } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { Sheet, type SheetTab } from './SheetContract';
import type { SheetEditing } from './SheetEditing';

/**
 * The tabs along the bottom: which sheets there are, and which one
 * you are on.
 *
 * **One tab stop, with the arrows moving inside it**, which is the
 * decision the menu bar and the toolbar already made and is made
 * again here for a third reason: a workbook can hold two hundred and
 * fifty-six sheets, and as tab stops that is a grid two hundred and
 * fifty-six presses from the keyboard.
 *
 * Arrowing onto a tab **shows** it rather than merely highlighting
 * it, which is ARIA's automatic-activation pattern and is what
 * anybody who has used Ctrl+PageDown expects. It costs nothing that a
 * click would not: showing a sheet is a viewport move, not an edit.
 *
 * What the strip does *not* own is the complete set of operations.
 * Renaming, duplicating, deleting, reordering and colouring are
 * commands in the `Sheet` menu, because Phase 4's standard is that
 * anything the mouse can do the keyboard can, and a colour chosen
 * from a swatch nobody can tab to is a colour only half the people
 * using this can set. The strip is the surface a pointer wants; the
 * menu is the one that is complete.
 */

/**
 * The colours a tab can be.
 *
 * Five and a plain one, named rather than free, because a tab colour
 * is a label people sort by — *the red one* — and a picker offering
 * sixteen million of them produces workbooks where nobody can say
 * which tab they mean.
 */
export const TAB_COLOURS: Readonly<Record<string, string>> = {
  blue: '#4285f4',
  red: '#ea4335',
  green: '#34a853',
  purple: '#b061f5',
  orange: '#fa7b17'
};

export interface SheetTabsProps {
  readonly editing: SheetEditing;
}

export function SheetTabs(inputs: Inputs<SheetTabsProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const edit = inputs.editing.value;
  const tabs = sheet.view.sheets;

  /**
   * Where the roving highlight is.
   *
   * It follows the active sheet rather than being a second position
   * to keep in step, except when it is resting on the add button —
   * which is not a sheet and has no active-ness to follow. `null`
   * means "on a tab, wherever the active one is".
   */
  const onAdd = internalState(false);

  /** The sheet being renamed, or null. Its index, not its name. */
  const renaming = internalState<number | null>(null);
  const draft = new BehaviorSubject('');
  let box: UiNode | null = null;
  let boxWanted = false;

  const deleting = internalState(false);

  const active = (): number => tabs.value.active;
  const count = (): number => tabs.value.entries.length;

  const show = (index: number): void => {
    if (index >= 0 && index < count() && index !== active()) {
      sheet.send.activateSheet(index);
    }
  };

  const startRename = (index = active()): void => {
    onAdd.value = false;
    show(index);
    draft.next(tabs.value.entries[index]?.name ?? '');
    renaming.value = index;
    boxWanted = true;
    take();
  };

  /**
   * The rename box, focused the moment it exists.
   *
   * The same request-and-honour the find bar needs: the box is not a
   * node until the frame after `renaming` changes, so the want is
   * remembered and the `ref` callback answers it when the node
   * arrives — at the exact moment rather than a plausible one.
   */
  const take = (): void => {
    if (!boxWanted || box === null) {
      return;
    }
    boxWanted = false;
    focus.focus(box);
    editorFor(box)?.selectAll();
  };

  const commitRename = (): void => {
    const index = renaming.value;
    renaming.value = null;
    if (index === null) {
      return;
    }
    const typed = draft.value.trim();
    if (typed !== '' && typed !== tabs.value.entries[index]?.name) {
      sheet.send.renameSheet(index, typed);
    }
    edit.focusSheet();
  };

  const cancelRename = (): void => {
    renaming.value = null;
    edit.focusSheet();
  };

  edit.provideRename(() => startRename());

  /**
   * Alt+F10 lands here, and it is the only way in from the sheet.
   *
   * Tab cannot be: the grid uses it to move one cell right, so a
   * keyboard standing in the sheet never leaves by that route. See
   * `COMMANDS.sheetTabs`.
   */
  let strip: UiNode | null = null;
  edit.provideTabs(() => {
    if (strip !== null) {
      onAdd.value = false;
      focus.focus(strip);
    }
  });

  const step = (by: number): void => {
    const last = count() - 1;
    if (onAdd.value) {
      // Leaving the add button goes back onto the last tab; there is
      // nothing to its right.
      if (by < 0) {
        onAdd.value = false;
        show(last);
      }
      return;
    }
    const next = active() + by;
    if (next > last) {
      onAdd.value = true;
      return;
    }
    show(Math.max(next, 0));
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
        onAdd.value = false;
        show(0);
        break;
      case 'End':
        onAdd.value = false;
        show(count() - 1);
        break;
      case 'Enter':
      case ' ':
        if (onAdd.value) {
          sheet.send.addSheet();
          onAdd.value = false;
        } else {
          startRename();
        }
        break;
      case 'F2':
        if (!onAdd.value) {
          startRename();
        }
        break;
      case 'Delete':
        if (!onAdd.value) {
          deleting.value = true;
        }
        break;
      default:
        // Everything else belongs to whatever is underneath, which is
        // what lets Tab out of the strip keep working.
        return;
    }
    event.preventDefault();
    event.stopPropagation();
  };

  const closeDelete = (): void => {
    deleting.value = false;
    edit.focusSheet();
  };

  const remove = (): void => {
    deleting.value = false;
    sheet.send.removeSheet(active());
    edit.focusSheet();
  };

  return (
    <column width={percent(100)} flexShrink={0}>
      <box width={percent(100)} height={1} backgroundColor="border" />
      <row
        ref={(node: UiNode | null) => (strip = node)}
        width={percent(100)}
        y="center"
        gap={2}
        paddingLeft={6}
        paddingRight={6}
        paddingTop={3}
        paddingBottom={3}
        backgroundColor="surface"
        focusable={true}
        role="tablist"
        label="Sheets"
        onKeyDown={onKeyDown}>
        {combineLatest([tabs, renaming]).pipe(
          map(([view, being]) =>
            view.entries.map((entry, index) =>
              index === being
                ? renameBox(entry, index, draft, commitRename, cancelRename, found => {
                    box = found;
                    take();
                  })
                : tab(entry, index, view.active, () => show(index))
            )
          )
        )}
        {addButton(() => {
          onAdd.value = false;
          sheet.send.addSheet();
        }, onAdd)}
        <box flex={1} minWidth={0} />
        {tabs.pipe(
          map(view => [
            <text
              key="count"
              text={view.entries.length === 1 ? '1 sheet' : `${view.entries.length} sheets`}
              fontSize={11}
              color="textMuted"
              verticalAlign="middle"
              selectable={false}
            />
          ])
        )}
      </row>
      {/**
       * The one destructive thing in this phase, so the one thing
       * that asks.
       *
       * Deleting a sheet takes every cell, format and merge on it and
       * is **not undoable** — the stack is dropped, because the
       * entries on it name their sheet by an index that removing one
       * renumbers. A question is what that costs, and it is the right
       * price: every other command here is a keystroke away from
       * being taken back.
       *
       * Built here rather than as a component because its text is an
       * Observable of the sheet being deleted, and a prop is read
       * through `Inputs`, which hands a component the current value
       * rather than the stream.
       */}
      <Dialog
        open={deleting}
        onClose={closeDelete}
        title="Delete sheet"
        width={360}
        content={Column(
          { gap: 12, minWidth: 0 },
          Text({
            text: tabs.pipe(
              map(view =>
                view.entries.length <= 1
                  ? 'A workbook keeps at least one sheet, so this one cannot go.'
                  : `Everything on ${view.entries[view.active]?.name ?? ''} goes with it, and this cannot be undone.`
              )
            ),
            fontSize: 12,
            color: 'text',
            textWrap: 'word'
          }),
          Row(
            { x: 'end', gap: 8 },
            plainButton('Cancel', closeDelete),
            plainButton('Delete sheet', remove, 'danger')
          )
        )}
      />
    </column>
  );
}

/**
 * One tab.
 *
 * The colour is a bar under the name rather than the tab's fill,
 * which is what a desktop spreadsheet does and for a reason worth
 * keeping: a fill dark enough to tell six colours apart is a fill the
 * name cannot be read on, and the name is what the tab is for.
 */
function tab(entry: SheetTab, index: number, active: number, onShow: () => void) {
  const is = index === active;
  return (
    <column key={`tab-${index}`} minWidth={0}>
      <button
        /**
         * Not a tab stop. The strip is the stop and the arrows move
         * inside it, so every tab is reachable and none of them
         * stands between the keyboard and the sheet.
         */
        focusable={false}
        onClick={onShow}
        label={entry.name}
        role="tab"
        states={(is ? ['selected'] : []) as readonly UiSemanticState[]}
        paddingLeft={10}
        paddingRight={10}
        paddingTop={4}
        paddingBottom={3}
        borderRadius={4}
        backgroundColor={is ? 'background' : 'surface'}
        borderColor={is ? 'focusRing' : 'surface'}
        borderWidth={1}
        cursor="pointer">
        <text
          text={entry.name}
          fontSize={12}
          fontWeight={is ? 600 : 'normal'}
          textWrap="none"
          color="text"
          selectable={false}
        />
      </button>
      <box
        width={percent(100)}
        height={3}
        backgroundColor={entry.colour ?? 'transparent'}
        borderRadius={2}
      />
    </column>
  );
}

/** The tab, while its name is being typed over. */
function renameBox(
  entry: SheetTab,
  index: number,
  draft: BehaviorSubject<string>,
  onCommit: () => void,
  onCancel: () => void,
  ref: (node: UiNode | null) => void
) {
  return (
    <column key={`tab-${index}`} minWidth={0}>
      <editabletext
        ref={ref}
        value={draft}
        width={96}
        fontSize={12}
        fontWeight={600}
        color="text"
        textWrap="none"
        textAlign="center"
        verticalAlign="middle"
        backgroundColor="background"
        borderColor="focusRing"
        borderWidth={1}
        paddingLeft={6}
        paddingRight={6}
        paddingTop={3}
        paddingBottom={2}
        role="textbox"
        label="Sheet name"
        onInput={(event: UiTextChangeEvent) => draft.next(event.value)}
        onKeyDown={(event: UiKeyboardEvent) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            event.stopPropagation();
            onCommit();
            return;
          }
          if (event.key === 'Escape') {
            event.preventDefault();
            event.stopPropagation();
            onCancel();
          }
        }}
      />
      <box width={percent(100)} height={3} backgroundColor={entry.colour ?? 'transparent'} borderRadius={2} />
    </column>
  );
}

function addButton(onAdd: () => void, resting: { value: boolean } & Observable<boolean>) {
  return (
    <button
      key="add"
      focusable={false}
      onClick={onAdd}
      label="Insert sheet"
      paddingLeft={9}
      paddingRight={9}
      paddingTop={4}
      paddingBottom={3}
      marginLeft={6}
      borderRadius={4}
      backgroundColor="surface"
      borderColor={resting.pipe(map(is => (is ? 'focusRing' : 'surface')))}
      borderWidth={1}
      cursor="pointer">
      <text text="+" fontSize={14} fontWeight={600} color="textMuted" selectable={false} />
    </button>
  );
}

function plainButton(label: string, onClick: () => void, colour = 'controlForeground') {
  return (
    <button
      onClick={onClick}
      label={label}
      paddingLeft={12}
      paddingRight={12}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} color={colour} selectable={false} />
    </button>
  );
}
