/**
 * The render worker: everything the person sees.
 *
 * A component cannot cross `postMessage`, so the root is named here
 * rather than passed in from `main.ts` — and so are the routes, for
 * the same reason: a route names a component too. `useChannel(Sheet)`
 * with no worker named resolves to whichever application worker the
 * shell spawned, so this file never learns where the cells come from.
 *
 * `notFound` is the plain sheet rather than an apology. A url this
 * application does not know is somebody's mistyped `/proof`, and the
 * right thing to hand them is the spreadsheet.
 */
import { renderRoot } from 'gesso-framework';

import { AppRoot, ROUTES, SHEET } from './app/Routes';
import { Sheet } from './app/SheetContract';

renderRoot(AppRoot).useChannel(Sheet).useRoutes({ routes: ROUTES, notFound: SHEET });
