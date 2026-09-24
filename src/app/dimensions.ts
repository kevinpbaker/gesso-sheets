/**
 * The sheet's geometry, in one place because two threads need it.
 *
 * The render worker sizes its window from these and the application
 * worker publishes them, so they are neither's to own. Phase 3 makes
 * the column width per-column and this becomes the default.
 */
export const ROW_HEIGHT = 24;
export const COLUMN_WIDTH = 104;
