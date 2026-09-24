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
 */
import { createApp } from 'gesso-framework';

const host = document.querySelector<HTMLElement>('#app');
if (host === null) {
  throw new Error('index.html has no #app element to mount into.');
}

createApp().mount(host);
