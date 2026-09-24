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

/** The sheet's extent. A million cells, as the roadmap's headline says. */
export const ROW_COUNT = 10_000;
export const COLUMN_COUNT = 100;
