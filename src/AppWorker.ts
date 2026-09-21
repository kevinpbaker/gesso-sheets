import { serve, serveChannels } from 'gesso-framework';

import { Sheet } from './spike/SheetContract';
import { SheetSource } from './spike/SheetSource';

/**
 * The sheet, on its own thread.
 *
 * `gesso-vite-plugin` finds this file by name and writes the
 * `appLogicWorker` construction into `createApp`, the same way it
 * writes the render worker, so `main.ts` names neither.
 */
const source = new SheetSource();

serveChannels([
  serve(Sheet, {
    view: { block: source.block, stats: source.stats },
    commands: {
      setViewport: (firstRow, lastRow, firstColumn, lastColumn) =>
        source.setViewport(firstRow, lastRow, firstColumn, lastColumn),
      setFetchBand: (rows, columns) => source.setFetchBand(rows, columns),
      setShape: shape => source.setShape(shape),
      setBusy: ms => source.setBusy(ms)
    }
  })
]);
