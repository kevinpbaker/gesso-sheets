import { serveChannels } from 'gesso-framework';

import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { seed } from './SheetSeed';
import { SheetService } from './SheetService';

/**
 * The sheet, on its own thread.
 *
 * Five lines of wiring over three plain classes and one call that
 * publishes the result. Above `serveChannels` there is no framework
 * import in this file's dependency graph at all — `SheetService` takes
 * RxJS and nothing else, `SheetDocument` and everything under
 * `src/sheet` take nothing — which is why every layer below is
 * testable with bare vitest in node, and `boundaries.spec.ts` fails
 * the build if that stops being true.
 *
 * `gesso-vite-plugin` finds this file by name and writes the
 * `appLogicWorker` construction into `createApp`, so `main.ts` names
 * neither worker.
 */
const document = new SheetDocument();
seed(document);
// Settled before the service exists, so that the first status it
// publishes is the true one. Recalculating afterwards would leave the
// service holding the count it was built with — the sheet would be
// right and the readout would say there was work outstanding forever.
document.sheet.recalculate();

const service = new SheetService(document);

serveChannels([sheetChannel(service)]);
