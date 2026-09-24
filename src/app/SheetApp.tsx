import { combineLatest, type Observable } from 'rxjs';
import { map } from 'rxjs/operators';

import { percent, type UiKeyboardEvent, type UiTextChangeEvent } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { columnName } from '../sheet/A1';
import { Grid } from './Grid';
import { Sheet } from './SheetContract';
import { editing } from './SheetEditing';
import { keyAction } from './SheetKeys';

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

  const onFormulaKey = (event: UiKeyboardEvent): void => {
    if (edit.apply(keyAction(event.key, event.modifiers, edit.openNow()))) {
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
        {historyButton('Undo', status.pipe(map(current => current.canUndo)), () => sheet.send.undo())}
        {historyButton('Redo', status.pipe(map(current => current.canRedo)), () => sheet.send.redo())}
        <text
          text={status.pipe(map(current => (current.pending === 0 ? 'Ready' : `${current.pending} to do`)))}
          width={90}
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

function historyButton(label: string, enabled: Observable<boolean>, onClick: () => void) {
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
