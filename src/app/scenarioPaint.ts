import type { CellPaint } from '../sheet/Format';

/**
 * The tints a scenario's cells are drawn in.
 *
 * Two, because a scenario changes a cell in one of two ways and they
 * answer different questions. A cell the scenario *types* is an
 * assumption — where to look to see what the scenario says. A cell it
 * only *changes* is a consequence — how far that assumption reaches.
 *
 * One hue, violet, at two strengths: violet means "the scenario", the
 * stronger one where it typed. Not amber, which was the first choice and
 * is the colour a workbook marks its own input cells in — the example's
 * are `#fff8d6`, and a changed cell beside an input read as another
 * input. Nor blue, the selection's wash. Both are pale enough that the
 * text over them is drawn in the light palette's ink in either theme
 * (see `inkFor`).
 */
export const SCENARIO_TYPED = '#c4b5fd';
export const SCENARIO_CHANGED = '#ede9fe';

/**
 * A cell's paint with a scenario's tint over its fill.
 *
 * The tint replaces the fill, rule or not: while a scenario is shown the
 * question on screen is which cells it moved. A text colour chosen for a
 * dark fill — white on a navy heading — is dropped, since it would be
 * white on pale yellow; any other is kept, so a rule's red still says
 * red.
 */
export function scenarioPaint(paint: CellPaint, standing: 'typed' | 'changed'): CellPaint {
  const fill = standing === 'typed' ? SCENARIO_TYPED : SCENARIO_CHANGED;
  return { ...paint, fill, color: isLight(paint.color) ? '' : paint.color };
}

/** Whether a `#rrggbb` colour is light enough to vanish on a pale tint. */
function isLight(color: string): boolean {
  const match = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color);
  if (match === null) {
    return false;
  }
  const [r, g, b] = [match[1], match[2], match[3]].map(hex => parseInt(hex, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6;
}
