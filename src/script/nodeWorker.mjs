// The browser's worker scope, as far as `scriptWorker.ts` uses it, on
// a Node worker thread: `postMessage` out and `onmessage` in. For the
// specs, which run the real worker file rather than a copy of it.
import { parentPort } from 'node:worker_threads';

let handler = null;
globalThis.postMessage = message => parentPort.postMessage(message);
Object.defineProperty(globalThis, 'onmessage', {
  get: () => handler,
  set: value => {
    handler = value;
  },
  configurable: true
});
parentPort.on('message', data => handler?.({ data }));
await import('./scriptWorker.ts');
