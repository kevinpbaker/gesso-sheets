/**
 * The icon set, and what this application calls each glyph.
 *
 * **Heroicons is the set.** One set rather than a drawer of glyphs
 * found one at a time, because a toolbar's icons are read as a group:
 * the stroke weight, the corner radius and the optical size have to
 * agree across all of them or the row looks assembled rather than
 * designed. Heroicons agrees with itself, is MIT licensed, and is
 * drawn on the 24-unit grid with round caps that Gesso's rasteriser
 * already strokes with — see `IconRasterizer` in `gesso-core`, whose
 * own docblock names it.
 *
 * The path data is lifted into `heroicons.ts` by `pnpm icons`; this
 * file is the layer above it, where a glyph gets an application's name
 * for it. `undo` rather than `arrowUturnLeft` is the point: what the
 * button means is stable, and which glyph says it is a decision that
 * can be revisited without touching a call site.
 *
 * **Where the set runs out.** Heroicons has no centred-text glyph, so
 * `alignCenter` below is drawn here, in Heroicons' own geometry, and
 * is the only path in this application that did not come from the
 * package. It is marked, and it is the thing to delete if the set ever
 * ships one.
 */
import {
  arrowPath,
  arrowUturnLeft,
  arrowUturnRight,
  bars3BottomLeft,
  bars3BottomRight,
  bold,
  currencyDollar,
  italic,
  percentBadge,
  underline
} from './heroicons';

/**
 * A glyph: a path in a square, and how to paint it.
 *
 * The same six fields `IconProps` takes, so an icon is passed along
 * rather than unpacked and reassembled at each place that draws one.
 */
export interface Glyph {
  /** SVG path data, in the coordinates of `viewBox`. */
  readonly path: string;
  /** The side of the square the path is drawn in. */
  readonly viewBox: number;
  readonly style: 'fill' | 'stroke';
  /** Line width for a stroked glyph, in viewBox units. */
  readonly strokeWidth?: number;
  /** How a filled path decides what is inside it. */
  readonly fillRule?: 'nonzero' | 'evenodd';
}

/**
 * Three lines with the last one centred: align centre.
 *
 * Heroicons ships `bars-3-bottom-left` and `bars-3-bottom-right` and
 * nothing between them, which leaves the middle button of an alignment
 * group with no glyph. Rather than borrow an unrelated one — a set's
 * icon used for something it does not mean is worse than a drawn one —
 * this is `bars-3-bottom-left` with its short line centred: the same
 * three rules at 6.75, 12 and 17.25, the same 8.25-unit length, moved
 * to sit symmetrically about the grid's centre at 12. Stroked at 1.5
 * like the rest of the outline set, so it weighs the same in the row.
 */
const barsThreeBottomCentre: Glyph = {
  path: 'M3.75 6.75h16.5M3.75 12h16.5M7.875 17.25h8.25',
  viewBox: 24,
  style: 'stroke',
  strokeWidth: 1.5
};

/**
 * What this application's buttons mean, and the glyph that says it.
 *
 * Keyed by intent, not by picture. A reader looking for the icon on
 * the Bold button looks up `bold`; a reader wondering what
 * `percentBadge` is doing in the bundle finds the one line that uses
 * it.
 */
export const ICONS = {
  undo: arrowUturnLeft,
  redo: arrowUturnRight,
  bold,
  italic,
  underline,
  alignLeft: bars3BottomLeft,
  alignCenter: barsThreeBottomCentre,
  alignRight: bars3BottomRight,
  currency: currencyDollar,
  percent: percentBadge,
  recalculate: arrowPath
} as const satisfies Record<string, Glyph>;

export type IconName = keyof typeof ICONS;
