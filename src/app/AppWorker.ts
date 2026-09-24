import { serveChannels } from 'gesso-framework';

import { COLUMN_COUNT } from './dimensions';
import { OpfsSheetRepository } from './OpfsSheetRepository';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { seed } from './SheetSeed';
import { SheetService } from './SheetService';

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
 * The order of the last two lines matters. `serveChannels` is called
 * synchronously, before any await, or the handshake is missed and the
 * render worker binds to a channel nobody is serving. Reading the file
 * happens after, and the screen shows an empty grid for the frame or
 * two it takes — which is honest, and better than a seeded sheet that
 * is about to be replaced by the real one.
 *
 * `gesso-vite-plugin` finds this file by name and writes the
 * `appLogicWorker` construction into `createApp`, so `main.ts` names
 * neither worker.
 */
const document = new SheetDocument();
const repository = new OpfsSheetRepository('gessosheet.json', COLUMN_COUNT);
const service = new SheetService(document, { repository });

serveChannels([sheetChannel(service)]);

void service.restore(seed);

// A tab being closed does not wait for a debounce. The write is
// synchronous once it starts, which is the other half of why this
// belongs on a worker: on the shell it would be a stall at exactly
// the moment a person is leaving.
self.addEventListener('beforeunload', () => void service.flush());
