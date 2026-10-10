import { serveChannels } from 'gesso-framework';

import { COLUMN_COUNT } from './dimensions';
import { OpfsSheetLibrary } from './OpfsSheetLibrary';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { seed } from './SheetSeed';
import type { PrintJob } from './SheetPrint';
import { SheetService } from './SheetService';
import { spawnScriptWorker } from '../script/browserWorker';

/**
 * The sheet, on its own thread.
 *
 * Six lines of wiring over four plain classes and one call that
 * publishes the result. Above `serveChannels` there is no framework
 * import in this file's dependency graph at all — `SheetService` takes
 * RxJS and nothing else, `SheetDocument`, the repository and
 * everything under `src/sheet` take nothing — which is why every layer
 * below is testable with bare vitest in node, and `boundaries.spec.ts`
 * fails the build if that stops being true.
 *
 * `serveChannels` is called synchronously, before any await, or the
 * handshake is missed and the render worker binds to a channel nobody
 * is serving. Nothing is read here at all: which workbook to show is
 * the route's, and the route is in the render worker, so the first
 * thing it sends is `openDocument`. The screen shows an empty grid for
 * the frame or two that takes — which is honest, and better than a
 * seeded sheet that is about to be replaced by the real one.
 *
 * `gesso-vite-plugin` finds this file by name and writes the
 * `appLogicWorker` construction into `createApp`, so `main.ts` names
 * neither worker.
 */
const library = new OpfsSheetLibrary(COLUMN_COUNT);
const service = new SheetService(new SheetDocument(), { library, seed, scripts: spawnScriptWorker, printer: printWindow() });

serveChannels([sheetChannel(service)]);

// A tab being closed does not wait for a debounce. The write is
// synchronous once it starts, which is the other half of why this
// belongs on a worker: on the shell it would be a stall at exactly
// the moment a person is leaving.
self.addEventListener('beforeunload', () => void service.flush());

/**
 * The print window's end of Print — Phase 37.
 *
 * The render worker opens `public/print.html` in a popup while the
 * click is fresh, and this worker builds the page; the two meet on a
 * `BroadcastChannel`, which reaches a window of the same origin from a
 * worker with no window of its own. The page may not have loaded when
 * the job is ready, so the last job is kept and sent again when the
 * page says it is listening.
 */
function printWindow(): ((job: PrintJob) => void) | undefined {
  if (typeof BroadcastChannel === 'undefined') {
    return undefined;
  }
  const channel = new BroadcastChannel('gessosheet-print');
  let last: PrintJob | null = null;
  channel.onmessage = event => {
    if ((event.data as { ready?: boolean } | null)?.ready === true && last !== null) {
      channel.postMessage(last);
    }
  };
  return job => {
    last = job;
    channel.postMessage(job);
  };
}
