/**
 * Whether a display string reads as a number, which is how the grid
 * places a cell whose alignment is `auto`.
 *
 * Shared by both workers, and that is the point of it being a module.
 * The grid decides from the string because the string is all it has;
 * the application worker has the value too, and uses this to find the
 * cells where the string misleads — `$4.50`, `9.6%`, `2026-09-24`,
 * which are numbers, and a Text cell holding `007`, which is not — so
 * that it can send an alignment for those and for nothing else. Two
 * copies of the rule would be two answers to what the grid will do.
 */
export function looksNumeric(text: string | null): boolean {
  if (text === null || text === '') {
    return false;
  }
  return !Number.isNaN(Number(text.replace(/,/g, '')));
}
