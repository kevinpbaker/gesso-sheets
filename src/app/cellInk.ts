import { darkColors, lightColors, parseColor, relativeLuminance, type UiColorValue } from 'gesso-core';

/**
 * The colour a cell's text is drawn in when the cell does not say.
 *
 * Without a fill that is the theme's `text`, which is what makes the
 * sheet follow light and dark. With one it cannot be: a fill is the
 * person's colour and does not change with the theme, so a pale header
 * written in the dark theme's white ink is a blank header. The ink is
 * chosen against the fill instead — the light palette's ink on a light
 * fill and the dark palette's on a dark one — which is the same answer
 * in both themes, because the fill is the same in both.
 *
 * The threshold is where black and white text have equal contrast
 * against the fill, so whichever is chosen is the more legible of the
 * two.
 */
const EQUAL_CONTRAST = 0.179;

const inks = new Map<string, UiColorValue>();

export function inkFor(fill: string): UiColorValue {
  if (fill === '') {
    return 'text';
  }
  let ink = inks.get(fill);
  if (ink === undefined) {
    const parsed = parseColor(fill);
    ink = parsed === undefined || parsed.a < 0.5 ? 'text' : relativeLuminance(parsed) > EQUAL_CONTRAST ? lightColors.text : darkColors.text;
    inks.set(fill, ink);
  }
  return ink;
}
