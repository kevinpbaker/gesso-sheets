/**
 * The main thread's entire job.
 *
 * It finds a host element, creates the app and mounts. Both workers
 * are written by `gesso-vite-plugin` from `vite.config.ts`: it finds
 * `worker.ts` beside this file for the render worker and `AppWorker.ts`
 * for the application worker, and writes the `new Worker(new URL(...))`
 * construction for each.
 *
 * The one thing said out loud is `workerName`. A worker has no page
 * url, so a flag that only the url carries has to be handed to it as
 * its name, and `?bench` is such a flag: it makes the render worker
 * run the Phase 0 measurement unattended instead of waiting for
 * someone to press a button. The idiom is Gesso's own `?still`.
 */
import { createApp } from 'gesso-framework';

const host = document.querySelector<HTMLElement>('#app');
if (host === null) {
  throw new Error('index.html has no #app element to mount into.');
}

const app = createApp({
  workerName: new URLSearchParams(location.search).has('bench') ? 'bench' : undefined
});

app.mount(host);
