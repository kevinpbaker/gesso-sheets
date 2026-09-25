/**
 * The sheet's geometry, in one place because two threads need it.
 *
 * The render worker sizes its window from these and the application
 * worker publishes them, so they are neither's to own. Phase 3 makes
 * the column width per-column and this becomes the default.
 */
export const ROW_HEIGHT = 24;
export const COLUMN_WIDTH = 104;
/** The frozen strip of row numbers down the left. */
export const GUTTER_WIDTH = 52;
/** The frozen strip of column letters across the top. */
export const HEADER_HEIGHT = 24;
/** A column may be dragged narrow, but not to nothing. */
export const MIN_COLUMN_WIDTH = 32;
/**
 * And autofit will not make one wider than this.
 *
 * A column holding one four-hundred-character note would otherwise
 * become a column nothing else fits beside, which is not what "fit
 * the contents" means to the person who asked for it.
 */
export const MAX_COLUMN_WIDTH = 480;
/** What a cell draws with, which autofit has to leave room for. */
export const CELL_PADDING = 6;
export const CELL_FONT_SIZE = 12;

/** The sheet's extent. A million cells, as the roadmap's headline says. */
export const ROW_COUNT = 10_000;
export const COLUMN_COUNT = 100;
