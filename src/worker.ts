/**
 * The render worker: everything the person sees.
 *
 * `useChannel(Sheet)` with no worker named resolves to whichever
 * application worker the shell spawned, so this file never learns
 * where the cells come from.
 */
import { renderRoot } from 'gesso-framework';
import { App } from './App';
import { Sheet } from './spike/SheetContract';

renderRoot(App).useChannel(Sheet);
