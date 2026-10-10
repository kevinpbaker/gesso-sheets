/**
 * The fonts the toolbar offers, and what each one is drawn with —
 * Phase 40.
 *
 * A cell keeps a font by its *name*, `Georgia`, because that is what a
 * person chose and what an `.xlsx` calls it. What the canvas is handed
 * is a stack, because a worker draws with the fonts the machine has and
 * a Linux desktop has no Georgia: the stack falls back to the metric
 * twin a distribution ships, and then to the generic family, so the
 * cell is at least the right *kind* of type everywhere.
 */

/** What a cell with no font of its own is drawn in: the theme's own stack. */
export const DEFAULT_FONT_STACK = 'system-ui, sans-serif';

export const FONTS: readonly { readonly name: string; readonly stack: string }[] = [
  { name: 'Arial', stack: 'Arial, Helvetica, "Liberation Sans", sans-serif' },
  { name: 'Calibri', stack: 'Calibri, Carlito, "Segoe UI", sans-serif' },
  { name: 'Courier New', stack: '"Courier New", "Liberation Mono", "DejaVu Sans Mono", monospace' },
  { name: 'Georgia', stack: 'Georgia, "DejaVu Serif", serif' },
  { name: 'Times New Roman', stack: '"Times New Roman", "Liberation Serif", Times, serif' },
  { name: 'Trebuchet MS', stack: '"Trebuchet MS", "DejaVu Sans", sans-serif' },
  { name: 'Verdana', stack: 'Verdana, "DejaVu Sans", sans-serif' }
];

/** The sizes the toolbar offers, in points; 12 is a cell's own. */
export const FONT_SIZES: readonly number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 30, 36];

/** The stack a font name is drawn with; a name not in the list is tried first, then the default. */
export function fontStack(name: string | undefined): string {
  if (name === undefined || name === '') {
    return DEFAULT_FONT_STACK;
  }
  return FONTS.find(font => font.name === name)?.stack ?? `"${name.replace(/"/g, '')}", ${DEFAULT_FONT_STACK}`;
}

/** Whether a stored name is one a file may carry: letters, digits, spaces and a little punctuation. */
export function isFontName(name: unknown): name is string {
  return typeof name === 'string' && /^[\p{L}\p{N} .'-]{1,40}$/u.test(name);
}
