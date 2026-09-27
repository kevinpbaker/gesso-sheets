import type { ScriptWorker } from './ScriptHost';

/**
 * A script worker in a browser: a module worker started by the
 * application worker, which holds the workbook and the `ScriptHost`.
 *
 * Nested rather than started by the shell, so a run never crosses the
 * main thread or the render worker: the workbook's values go straight
 * from the thread that has them to the worker that reads them, and the
 * changes come straight back. `new URL(…, import.meta.url)` is the form
 * Vite finds and bundles as a worker of its own.
 */
export function spawnScriptWorker(): ScriptWorker {
  const worker = new Worker(new URL('./scriptWorker.ts', import.meta.url), { type: 'module', name: 'script' });
  return {
    post: message => worker.postMessage(message),
    onMessage: listener => worker.addEventListener('message', event => listener(event.data)),
    terminate: () => worker.terminate()
  };
}
