import { dropTarget, EXTERNAL_FILES, percent, type UiDroppedFile } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { Grid } from './Grid';
import { Sheet } from './SheetContract';
import { editing } from './SheetEditing';
import { fileActions, sendFile } from './Files';
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import { TopBar } from './TopBar';

/**
 * The screen: the chrome, the grid, and the line along the bottom.
 *
 * Nothing here holds application state. The address, the cell's text,
 * the selection's total and whether there is anything to undo are
 * view keys off the channel or the shared editing handle, and the
 * controls send commands back. The sheet itself — the store, the
 * parser, the dependency graph, the recalc — is on the other thread
 * and shares nothing with this file but the token.
 *
 * Four pieces rather than one since Phase 13, and the split is by who
 * owns the keyboard: `TopBar` owns the menus, the fields and command
 * dispatch, `Grid` owns the sheet, `SheetTabs` owns the strip along
 * the bottom, and `StatusBar` owns nothing at all and is the only one
 * of the four that cannot be focused.
 *
 * The same screen on both routes, and deliberately the same one: the
 * proof route is not a different application with a spreadsheet in
 * it, it is this spreadsheet with a strip of instruments bolted to
 * the page around it. All `proof` changes down here is whether the
 * chrome offers the one command that exists to feed those
 * instruments — see `Routes.tsx` for where the boolean comes from and
 * `PROOF_ONLY` for what it covers.
 */
export interface SheetAppProps {
  readonly proof?: boolean;
}

export function SheetApp(inputs: Inputs<SheetAppProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = editing(ctx, sheet);
  const files = fileActions(ctx, sheet);

  /**
   * A file dragged in from the desktop, opened.
   *
   * The whole window is the zone rather than the grid, because a
   * person dropping a file is aiming at the application and not at a
   * cell — and the file lands on a sheet of its own either way, so
   * where it was let go says nothing about where it goes.
   *
   * Sent by kind through `sendFile`, the same path a picked file
   * takes: a workbook or an Excel file opens as a document of its own,
   * and anything else as a sheet in this one. What each file *is* is
   * the application worker's question, and it turns down one it cannot
   * read and says so.
   */
  const drop = dropTarget({
    accepts: EXTERNAL_FILES,
    onDrop: payload => {
      for (const file of payload.data as readonly UiDroppedFile[]) {
        sendFile(sheet, file.name, file.bytes, null);
      }
      return 'copy';
    },
    over: { borderColor: 'focusRing', borderWidth: 2 }
  });

  return (
    <column
      width={percent(100)}
      height={percent(100)}
      backgroundColor="background"
      borderColor="transparent"
      borderWidth={0}
      modifiers={[drop]}>
      <TopBar editing={edit} files={files} proof={inputs.proof.value === true} />
      <Grid editing={edit} />
      <SheetTabs editing={edit} />
      <StatusBar />
    </column>
  );
}
