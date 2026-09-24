/**
 * The render worker: everything the person sees.
 *
 * A component cannot cross `postMessage`, so the root is named here
 * rather than passed in from `main.ts`. `useChannel(Sheet)` with no
 * worker named resolves to whichever application worker the shell
 * spawned, so this file never learns where the cells come from.
 */
import { renderRoot } from 'gesso-framework';

import { Sheet } from './app/SheetContract';
import { SheetApp } from './app/SheetApp';

renderRoot(SheetApp).useChannel(Sheet);
