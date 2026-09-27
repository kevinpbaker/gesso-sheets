/**
 * A colour's name, for a person: "red", "pale green", "dark blue".
 *
 * Worked out from the colour rather than looked up, because a rule from
 * a file or the seed carries any colour at all, and `#fde2e2` in a list
 * of rules tells nobody anything.
 */
export function colourName(hex: string): string {
  const match = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (match === null) {
    return hex;
  }
  const value = Number.parseInt(match[1], 16);
  const r = ((value >> 16) & 255) / 255;
  const g = ((value >> 8) & 255) / 255;
  const b = (value & 255) / 255;
  const high = Math.max(r, g, b);
  const low = Math.min(r, g, b);
  const lightness = (high + low) / 2;
  const chroma = high - low;
  if (lightness > 0.97) {
    return 'white';
  }
  if (lightness < 0.08) {
    return 'black';
  }
  const shade = lightness > 0.82 ? 'pale ' : lightness < 0.2 ? 'dark ' : '';
  if (chroma < 0.06) {
    return `${shade}grey`;
  }
  const hue =
    high === r ? ((g - b) / chroma + 6) % 6 : high === g ? (b - r) / chroma + 2 : (r - g) / chroma + 4;
  const degrees = hue * 60;
  const name =
    degrees < 15 || degrees >= 345
      ? 'red'
      : degrees < 40
        ? 'orange'
        : degrees < 70
          ? 'yellow'
          : degrees < 165
            ? 'green'
            : degrees < 200
              ? 'teal'
              : degrees < 255
                ? 'blue'
                : degrees < 290
                  ? 'purple'
                  : 'pink';
  return `${shade}${name}`;
}
