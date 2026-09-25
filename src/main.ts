/**
 * The main thread's entire job.
 *
 * It finds a host element, creates the app and mounts. Both workers
 * are written by `gesso-vite-plugin` from `vite.config.ts`: it finds
 * `worker.ts` beside this file for the render worker and `AppWorker.ts`
 * for the application worker, and writes the `new Worker(new URL(...))`
 * construction for each. Everything a person sees is built, laid out,
 * painted and hit-tested in the render worker, and the sheet itself
 * lives on a third thread, so work on this one cannot delay a frame.
 *
 * `history: { mode: 'path' }` is the shell's whole half of routing:
 * it reports the url the window is at and performs the pushes the
 * worker asks for. Which screen that url means is decided over there,
 * in `Routes.tsx`, where the components are.
 *
 * The proof strip is the exception that says so out loud — the only
 * DOM in the application, and the only thing here that does any work
 * — and it exists on one route and not the other. `/` is a
 * spreadsheet and gets a spreadsheet; `/proof` gets the instruments
 * as well. The url is read twice, once here and once by the router,
 * and that is not a duplication that can be removed: this half is DOM
 * on the page, the other half is a component in a worker, and the
 * thread that owns each is the thread that has to ask. See
 * `ProofPanel`.
 */
import { createApp } from 'gesso-framework';

import { isProofPath } from './route';
import { proofPanel } from './shell/ProofPanel';

const host = document.querySelector<HTMLElement>('#app');
if (host === null) {
  throw new Error('index.html has no #app element to mount into.');
}

const panel = isProofPath(location.pathname) ? proofPanel(host) : null;
const app = createApp({ ...(panel?.options ?? {}), history: { mode: 'path' } });
panel?.attach(app);
app.mount(host);
