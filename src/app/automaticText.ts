/**
 * The colour "automatic" text is drawn in, over a cell's fill.
 *
 * With no fill it is the theme's own `text`, which is right on the
 * theme's own background and nowhere else. A fill is a person's choice
 * and stays the colour they chose in light and dark alike, so the text
 * on it has to follow the fill rather than the theme: a pale yellow
 * cell under a dark theme's near-white text is a cell nobody can read.
 *
 * So a filled cell takes whichever of black and white stands out from
 * the fill — the two themes' own text colours — which is what Excel
 * means by automatic as well. Measured by relative luminance, with
 * the crossover where the two contrasts are equal, so neither side is
 * favoured.
 *
 * A fill that is not `#rgb` or `#rrggbb` has no luminance to measure
 * and keeps the theme's colour.
 */
export function automaticText(fill: string): string {
  if (fill === '') {
    return 'text';
  }
  const value = luminance(fill);
  if (value === null) {
    return 'text';
  }
  return value > CROSSOVER ? '#000000' : '#ffffff';
}

/** Where black and white text have the same contrast against a fill. */
const CROSSOVER = Math.sqrt(1.05 * 0.05) - 0.05;

function luminance(hex: string): number | null {
  const match = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(hex.trim());
  if (match === null) {
    return null;
  }
  const digits = match[1].length === 3 ? [...match[1]].map(d => d + d).join('') : match[1];
  const value = Number.parseInt(digits, 16);
  const [r, g, b] = [(value >> 16) & 255, (value >> 8) & 255, value & 255].map(channel => {
    const c = channel / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
