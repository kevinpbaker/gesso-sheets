import { percent } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { Grid } from './Grid';
import { Sheet } from './SheetContract';
import { editing } from './SheetEditing';
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
 * Three pieces rather than one since Phase 8, and the split is by who
 * owns the keyboard: `TopBar` owns the menus, the fields and command
 * dispatch, `Grid` owns the sheet, and `StatusBar` owns nothing at
 * all and is the only one of the three that cannot be focused.
 */
export function SheetApp(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = editing(ctx, sheet);

  return (
    <column width={percent(100)} height={percent(100)} backgroundColor="background">
      <TopBar editing={edit} />
      <Grid editing={edit} />
      <StatusBar />
    </column>
  );
}
