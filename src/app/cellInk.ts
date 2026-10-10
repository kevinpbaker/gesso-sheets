import { darkColors, lightColors, parseColor, relativeLuminance, type UiColorValue } from 'gesso-core';

import type { CellPaint } from '../sheet/Format';

/**
 * The colour a cell's text is drawn in when the cell does not say.
 *
 * Without a fill that is the theme's `text`, which is what makes the
 * sheet follow light and dark. With one it cannot be: a fill is the
 * person's colour, and a pale header written in the dark theme's white
 * ink is a blank header. The ink is chosen against the fill instead —
 * the light palette's ink on a light fill and the dark palette's on a
 * dark one. It is asked of the fill as *shown*, which in the dark
 * theme is not always the fill as written: see `shownFill`.
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

/**
 * How light a fill has to be before the dark theme darkens it, as an
 * HSL lightness: the pastels and pale greys a sheet is mostly filled
 * with, and not the navy of a title bar or the full colour of a
 * highlighter yellow.
 */
const LIGHT_FILL = 0.7;

const shown = new Map<string, string>();

/**
 * The fill a cell is drawn with, in the theme it is drawn in.
 *
 * The light theme draws the author's colour, always. The dark theme
 * did too, and a header row of pale grey and a colour scale of pastels
 * were bright stripes across a dark sheet, the loudest thing on the
 * screen and saying the least. So a *light* fill is drawn in the dark
 * theme as its dark equivalent: the same hue, its lightness turned
 * over — the palest fill darkest, closest to the sheet — and its
 * saturation eased so a pastel does not come back as a neon. A fill
 * that is not light — the navy title, a saturated mid-tone — is
 * already at home on a dark sheet and is drawn as written.
 *
 * Turned over rather than set to one value, so fills that differed in
 * the light theme still differ: a colour scale keeps its steps, and
 * the scenario's two violets (`#c4b5fd` where it typed, `#ede9fe`
 * where it moved) are still two.
 */
export function shownFill(fill: string, dark: boolean): string {
  if (!dark || fill === '') {
    return fill;
  }
  let drawn = shown.get(fill);
  if (drawn === undefined) {
    drawn = darkened(fill);
    shown.set(fill, drawn);
  }
  return drawn;
}

function darkened(fill: string): string {
  const colour = hslOf(fill);
  if (colour === null || colour.lightness <= LIGHT_FILL) {
    return fill;
  }
  // The palest (lightness 1) lands at 0.13, a shade above the dark
  // theme's background; the least pale still darkened (0.7) at 0.355.
  return hsl(colour.hue, colour.saturation * 0.6, 0.13 + (1 - colour.lightness) * 0.75);
}

/**
 * The text colour an author chose for a light fill, on that fill
 * darkened.
 *
 * A highlight is a pair — pale red with dark red text, pale green with
 * dark green — and darkening the fill alone left dark red on dark
 * red. A dark ink on a fill the theme darkened is turned over too, to
 * the light shade of its own hue; a light one already reads and is
 * left alone.
 */
function lightened(color: string): string {
  const parsed = parseColor(color);
  const colour = hslOf(color);
  if (parsed === undefined || colour === null || relativeLuminance(parsed) > EQUAL_CONTRAST) {
    return color;
  }
  return hsl(colour.hue, colour.saturation, 0.6 + (1 - colour.lightness) * 0.3);
}

/** A colour as a hue in sixths of a turn, a saturation and a lightness; null when it is not one, or mostly see-through. */
function hslOf(value: string): { hue: number; saturation: number; lightness: number } | null {
  const parsed = parseColor(value);
  if (parsed === undefined || parsed.a < 0.5) {
    return null;
  }
  const { r, g, b } = parsed;
  const high = Math.max(r, g, b);
  const low = Math.min(r, g, b);
  const lightness = (high + low) / 2;
  const chroma = high - low;
  const saturation = chroma === 0 ? 0 : chroma / (1 - Math.abs(2 * lightness - 1));
  const hue =
    chroma === 0
      ? 0
      : high === r
        ? ((g - b) / chroma + 6) % 6
        : high === g
          ? (b - r) / chroma + 2
          : (r - g) / chroma + 4;
  return { hue, saturation, lightness };
}

/** `#rrggbb` from a hue in sixths of a turn, a saturation and a lightness. */
function hsl(hue: number, saturation: number, lightness: number): string {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const second = chroma * (1 - Math.abs((hue % 2) - 1));
  const [r, g, b] =
    hue < 1
      ? [chroma, second, 0]
      : hue < 2
        ? [second, chroma, 0]
        : hue < 3
          ? [0, chroma, second]
          : hue < 4
            ? [0, second, chroma]
            : hue < 5
              ? [second, 0, chroma]
              : [chroma, 0, second];
  const base = lightness - chroma / 2;
  return `#${[r, g, b].map(channel => Math.round((channel + base) * 255).toString(16).padStart(2, '0')).join('')}`;
}

/**
 * A palette as the theme draws it: each entry's fill through
 * `shownFill`, with an author's dark text on a darkened fill turned
 * light (see `lightened`), and the entry itself kept wherever that
 * changes nothing, so the light theme's palette is the one it was given and a
 * cell's paint is pushed again only when it will look different.
 */
export function shownPalette(entries: readonly CellPaint[], dark: boolean): readonly CellPaint[] {
  if (!dark) {
    return entries;
  }
  return entries.map(paint => {
    const fill = shownFill(paint.fill, true);
    if (fill === paint.fill) {
      return paint;
    }
    return { ...paint, fill, color: paint.color === '' ? '' : lightened(paint.color) };
  });
}
