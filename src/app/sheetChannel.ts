import { serve, type ServedChannel } from 'gesso-framework';

import { Sheet } from './SheetContract';
import type { SheetService } from './SheetService';

/**
 * The channel, wired once.
 *
 * Every view key and every command in one place, so that the worker
 * and the four specs that mount a sheet cannot drift from each other.
 * They did: adding a command meant editing four literals, and the
 * compiler only said so because the token's type demanded all of them.
 */
export function sheetChannel(service: SheetService): ServedChannel {
  return serve(Sheet, {
    view: {
      window: service.window,
      geometry: service.geometry,
      selection: service.selection,
      editor: service.editor,
      status: service.status,
      clipboard: service.clipboard,
      stats: service.selectionStats,
      find: service.findView,
      formats: service.formats,
      palette: service.palette,
      activeFormat: service.activeFormat
    },
    commands: {
      setViewport: (firstRow, lastRow, firstColumn, lastColumn) =>
        service.setViewport(firstRow, lastRow, firstColumn, lastColumn),
      setCell: (row, column, input) => service.setCell(row, column, input),
      setSelection: (row, column, anchorRow, anchorColumn) =>
        service.setSelection(row, column, anchorRow, anchorColumn),
      undo: () => service.undo(),
      redo: () => service.redo(),
      copy: cut => service.copy(cut),
      paste: text => service.paste(text),
      clearRange: () => service.clearRange(),
      fill: (toRow, toColumn) => service.fill(toRow, toColumn),
      setColumnWidth: (column, width) => service.setColumnWidth(column, width),
      stress: cells => service.stress(cells),
      fillDown: () => service.fillDown(),
      fillRight: () => service.fillRight(),
      find: (query, matchCase, wholeCell, inFormulas) => service.find(query, matchCase, wholeCell, inFormulas),
      findStep: forward => service.findStep(forward),
      replaceOne: replacement => service.replaceOne(replacement),
      replaceAll: replacement => service.replaceAll(replacement),
      clearFind: () => service.clearFind(),
      format: change => service.format(change),
      clearFormat: () => service.clearFormat(),
      insertRows: (at, count) => service.insertRows(at, count),
      deleteRows: (at, count) => service.deleteRows(at, count),
      insertColumns: (at, count) => service.insertColumns(at, count),
      deleteColumns: (at, count) => service.deleteColumns(at, count)
    }
  });
}
