import { map } from 'rxjs/operators';

import { percent } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { Grid } from './Grid';
import { columnName } from '../sheet/A1';
import { Sheet } from './SheetContract';

/**
 * The screen: a formula bar, the grid, and a line saying what the
 * application thread is doing.
 *
 * Nothing here holds state. The address, the formula text, and whether
 * there is anything to undo are all view keys off the channel, and the
 * buttons send commands back. The sheet itself — the store, the
 * parser, the dependency graph, the recalc — is on the other thread
 * and shares nothing with this file but the token.
 */
export function SheetApp(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const editor = sheet.view.editor;
  const status = sheet.view.status;

  const address = editor.pipe(map(current => `${columnName(current.column)}${current.row + 1}`));
  const formula = editor.pipe(map(current => (current.input === '' ? '—' : current.input)));

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
          role="textbox"
          label="Active cell"
        />
        <text
          text={formula}
          flex={1}
          minWidth={0}
          fontSize={12}
          color="text"
          textWrap="none"
          textOverflow="clip"
          verticalAlign="middle"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={4}
          role="textbox"
          label="Formula"
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
      <Grid />
    </column>
  );
}

function historyButton(label: string, enabled: import('rxjs').Observable<boolean>, onClick: () => void) {
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
