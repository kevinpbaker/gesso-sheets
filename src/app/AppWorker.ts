import { serve, serveChannels } from 'gesso-framework';

import { Sheet } from './SheetContract';
import { SheetDocument } from './SheetDocument';
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

const service = new SheetService(document);
// The seed queued a few dozen formulas; settle them before anyone
// looks, so the first window is the sheet and not its skeleton.
document.sheet.recalculate();

serveChannels([
  serve(Sheet, {
    view: {
      window: service.window,
      geometry: service.geometry,
      selection: service.selection,
      editor: service.editor,
      status: service.status
    },
    commands: {
      setViewport: (firstRow, lastRow, firstColumn, lastColumn) =>
        service.setViewport(firstRow, lastRow, firstColumn, lastColumn),
      setCell: (row, column, input) => service.setCell(row, column, input),
      setSelection: (row, column, anchorRow, anchorColumn) =>
        service.setSelection(row, column, anchorRow, anchorColumn),
      undo: () => service.undo(),
      redo: () => service.redo()
    }
  })
]);
