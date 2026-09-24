import { combineLatest, of, type Observable } from 'rxjs';
import { map } from 'rxjs/operators';

import { percent, type UiKeyboardEvent, type UiTextChangeEvent } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName } from '../sheet/A1';
import { Grid } from './Grid';
import { Sheet } from './SheetContract';
import { editing } from './SheetEditing';
import { keyAction } from './SheetKeys';

/**
 * How long a chain the proof button builds.
 *
 * Two hundred thousand cells, every one of them depending on the one
 * before it, so the recalculation cannot be parallelised or skipped —
 * it is the longest possible critical path through the sheet. None of
 * them are on screen.
 */
const STRESS_CELLS = 200_000;

/**
 * The status line: what the application thread still owes, and how
 * much it has ever done.
 *
 * The second half is what makes the first half worth reading. "Ready"
 * on its own is also what a sheet that did nothing would say.
 */
function describe(status: { pending: number; evaluated: number }): string {
  const done = status.evaluated === 0 ? '' : ` \u00b7 ${status.evaluated.toLocaleString()} evaluated`;
  return `${status.pending === 0 ? 'Ready' : `${status.pending.toLocaleString()} to do`}${done}`;
}

/**
 * The screen: a formula bar, the grid, and a line saying what the
 * application thread is doing.
 *
 * Nothing here holds application state. The address, the cell's text,
 * and whether there is anything to undo are view keys off the channel
 * or the shared editing handle, and the buttons send commands back.
 * The sheet itself — the store, the parser, the dependency graph, the
 * recalc — is on the other thread and shares nothing with this file
 * but the token.
 */
export function SheetApp(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = editing(ctx, sheet);
  const status = sheet.view.status;

  const address = edit.selection.pipe(map(at => `${columnName(at.column)}${at.row + 1}`));

  /**
   * What the formula bar shows: the draft while a cell is open, and
   * what the application worker says the cell holds otherwise.
   *
   * The same buffer as the cell, not a copy of it. Two buffers kept in
   * step would be two answers to what Escape puts back, and the bar
   * and the cell would disagree for exactly as long as it took a
   * keystroke to cross between them.
   */
  const formula: Observable<string> = combineLatest([edit.draft, sheet.view.editor]).pipe(
    map(([draft, current]) => draft ?? current.input)
  );

  /**
   * Keys in the formula bar.
   *
   * Always read as "a cell is open", whatever the draft says, because
   * the caret is in a text field and the keys belong to the text.
   * Read the other way — as "a cell is selected and nothing is open" —
   * Backspace meant *empty this cell* and a digit meant *replace this
   * cell*, so deleting one character wiped the lot, and the next
   * character arrived twice: once from the key table seeding a draft
   * and once from the field inserting it.
   *
   * Enter and Escape are still the sheet's, which is what makes this a
   * formula bar rather than a text box that happens to sit above a
   * grid.
   */
  const onFormulaKey = (event: UiKeyboardEvent): void => {
    if (edit.apply(keyAction(event.key, event.modifiers, true))) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  return (
    <column width={percent(100)} height={percent(100)} backgroundColor="background">
      <row
        width={percent(100)}
        flexShrink={0}
        y="center"
        gap={10}
        padding={8}
        backgroundColor="surface"
        borderColor="border"
        borderWidth={1}>
        <text
          text={address}
          width={70}
          fontSize={12}
          fontWeight={600}
          color="text"
          textAlign="center"
          verticalAlign="middle"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={4}
          role="status"
          label="Active cell"
        />
        <editabletext
          value={formula}
          flex={1}
          minWidth={0}
          fontSize={12}
          color="text"
          textWrap="none"
          verticalAlign="middle"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={4}
          role="textbox"
          label="Formula"
          onInput={(event: UiTextChangeEvent) => edit.write(event.value)}
          onKeyDown={onFormulaKey}
        />
        {barButton('Undo', status.pipe(map(current => current.canUndo)), () => sheet.send.undo())}
        {barButton('Redo', status.pipe(map(current => current.canRedo)), () => sheet.send.redo())}
        {barButton(`Recalculate ${STRESS_CELLS.toLocaleString()}`, of(true), () => sheet.send.stress(STRESS_CELLS))}
        <text
          text={status.pipe(map(describe))}
          width={190}
          fontSize={11}
          color="textMuted"
          verticalAlign="middle"
          textAlign="end"
          live="polite"
        />
      </row>
      <Grid editing={edit} />
    </column>
  );
}

function barButton(label: string, enabled: Observable<boolean>, onClick: () => void) {
  return (
    <button
      onClick={onClick}
      label={label}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer"
      opacity={enabled.pipe(map(on => (on ? 1 : 0.4)))}>
      <text text={label} fontSize={12} color="controlForeground" />
    </button>
  );
}
