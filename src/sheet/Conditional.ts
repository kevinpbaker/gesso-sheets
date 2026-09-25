import type { RangeRef } from './A1';
import { isError, type CellValue } from './Values';

/**
 * Formats that think: a rule over a range, resolved where it is seen.
 *
 * **A conditional format is resolved at the window, not in the
 * graph.** A rule covering a million cells does not need a million
 * graph nodes, because a format nobody can see does not exist. The
 * application worker asks each rule about the cells in the window at
 * publish time and folds the answer into the palette index Phase 9 is
 * already sending, so a rule over the whole sheet costs the viewport.
 *
 * Nothing here imports the framework or knows what a viewport is —
 * it is the same constraint every file in this directory is under.
 * What it knows is: given a rule, a value and the extent of the range
 * the rule covers, what paint does this cell take.
 *
 * ## Why the extent is passed in
 *
 * A colour scale needs the smallest and largest number in its range,
 * and that is not a question about one cell. Working it out here
 * would mean this file reading the sheet; working it out per cell
 * would mean reading the range once per cell. So the caller computes
 * it once — see `extentOf` — and hands it over.
 */

/** What a rule paints when it matches. Every field is an override. */
export interface ConditionalPaint {
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

/**
 * What a rule asks of a cell.
 *
 * `formula` is the general case and the others are the shorthands
 * everybody actually uses. They are kept as their own kinds rather
 * than being sugar over `formula` because they can be answered
 * without parsing or evaluating anything — which is the difference
 * between a rule costing a comparison per visible cell and costing a
 * formula evaluation per visible cell, on every publish.
 */
export type ConditionalTest =
  | { readonly kind: 'greaterThan'; readonly value: number }
  | { readonly kind: 'lessThan'; readonly value: number }
  | { readonly kind: 'between'; readonly low: number; readonly high: number }
  | { readonly kind: 'equalTo'; readonly value: number | string }
  | { readonly kind: 'textContains'; readonly text: string }
  | { readonly kind: 'isEmpty' }
  | { readonly kind: 'notEmpty' }
  /**
   * A formula, written as it would be for the range's first cell and
   * moved to each cell it is asked about.
   *
   * The expensive kind, and the reason the others exist. Held as
   * text here; the application worker parses it once and shifts the
   * tree per cell, because parsing per cell per publish is the thing
   * that would make this unaffordable.
   */
  | { readonly kind: 'formula'; readonly input: string };

/**
 * A colour scale: the cell's colour from where its value sits
 * between the smallest and the largest in the range.
 *
 * **Quantised into steps**, and that is a decision rather than a
 * shortcut. Continuous, a scale over a window of a hundred and fifty
 * cells is a hundred and fifty distinct paints, every one of which
 * has to be interned and sent; quantised it is at most as many as
 * there are steps, they are the same ones after a scroll, and nobody
 * can see the difference between a gradient and twenty-five of it.
 */
export interface ColourScale {
  readonly from: string;
  /** The middle stop, for a three-colour scale. */
  readonly middle?: string;
  readonly to: string;
}

/**
 * An **odd** number, which is the whole of why it is 25 and not 24.
 *
 * The steps run from 0 to `SCALE_STEPS - 1`, so an even count has no
 * step exactly halfway — and halfway is where a three-colour scale
 * keeps its middle stop. At 24 steps the middle of a red-white-blue
 * scale came out faintly blue, which is the sort of thing nobody
 * reports and everybody notices.
 */
export const SCALE_STEPS = 25;

export interface ConditionalRule {
  readonly range: RangeRef;
  readonly test: ConditionalTest | null;
  /** The paint when `test` matches. Ignored for a scale. */
  readonly paint?: ConditionalPaint;
  /** A colour scale instead of a test, for the range as a whole. */
  readonly scale?: ColourScale;
}

/** The smallest and largest numbers a scale has to spread between. */
export interface Extent {
  readonly low: number;
  readonly high: number;
}

/**
 * Whether a plain test matches a value.
 *
 * `formula` is not answered here: it needs an evaluator and a sheet,
 * neither of which belongs in a pure predicate over one value. The
 * caller handles that kind; this returns null to say so, which is
 * different from false.
 */
export function matches(test: ConditionalTest, value: CellValue): boolean | null {
  if (isError(value)) {
    // An error is not greater than five and not less than it either.
    // Only the emptiness tests have an honest answer about a cell
    // that is showing `#DIV/0!`.
    return test.kind === 'isEmpty' ? false : test.kind === 'notEmpty' ? true : false;
  }
  switch (test.kind) {
    case 'greaterThan':
      return typeof value === 'number' && value > test.value;
    case 'lessThan':
      return typeof value === 'number' && value < test.value;
    case 'between':
      return typeof value === 'number' && value >= test.low && value <= test.high;
    case 'equalTo':
      return typeof test.value === 'number'
        ? value === test.value
        : typeof value === 'string' && value.toUpperCase() === test.value.toUpperCase();
    case 'textContains':
      return typeof value === 'string' && value.toUpperCase().includes(test.text.toUpperCase());
    case 'isEmpty':
      return value === null || value === '';
    case 'notEmpty':
      return value !== null && value !== '';
    case 'formula':
      return null;
  }
}

/**
 * A scale's colour for a value, or null when the value is not a
 * number or the range has nothing to spread between.
 *
 * A range whose numbers are all the same has no scale — every cell
 * is both the smallest and the largest — and the honest answer is
 * the middle of it rather than an arbitrary end.
 */
export function scaleColour(scale: ColourScale, value: CellValue, extent: Extent): string | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return null;
  }
  const span = extent.high - extent.low;
  const raw = span === 0 ? 0.5 : (value - extent.low) / span;
  const clamped = Math.min(Math.max(raw, 0), 1);
  // Quantised before it is interpolated, so the same step always
  // produces the same colour and the palette can intern it.
  const step = Math.round(clamped * (SCALE_STEPS - 1)) / (SCALE_STEPS - 1);
  if (scale.middle === undefined) {
    return mix(scale.from, scale.to, step);
  }
  return step <= 0.5 ? mix(scale.from, scale.middle, step * 2) : mix(scale.middle, scale.to, (step - 0.5) * 2);
}

/**
 * Two colours, mixed.
 *
 * In sRGB, which is the wrong colour space and the right answer: a
 * spreadsheet's scales are matched against other spreadsheets', and
 * every one of those mixes in sRGB. A perceptually even gradient
 * would be a scale that looked different from the same rule in Excel.
 */
function mix(from: string, to: string, at: number): string {
  const a = rgbOf(from);
  const b = rgbOf(to);
  if (a === null || b === null) {
    return from;
  }
  const channel = (one: number, other: number): string =>
    Math.round(one + (other - one) * at)
      .toString(16)
      .padStart(2, '0');
  return `#${channel(a[0], b[0])}${channel(a[1], b[1])}${channel(a[2], b[2])}`;
}

function rgbOf(colour: string): [number, number, number] | null {
  const text = colour.trim();
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(text);
  if (short !== null) {
    return [
      parseInt(short[1] + short[1], 16),
      parseInt(short[2] + short[2], 16),
      parseInt(short[3] + short[3], 16)
    ];
  }
  const long = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(text);
  if (long === null) {
    return null;
  }
  return [parseInt(long[1], 16), parseInt(long[2], 16), parseInt(long[3], 16)];
}

/** Whether a cell is inside a rule's range. */
export function covers(range: RangeRef, row: number, column: number): boolean {
  const firstRow = Math.min(range.start.row, range.end.row);
  const lastRow = Math.max(range.start.row, range.end.row);
  const firstColumn = Math.min(range.start.column, range.end.column);
  const lastColumn = Math.max(range.start.column, range.end.column);
  return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
}

/**
 * The rules over a cell, folded into one set of overrides.
 *
 * **Later rules win, field by field.** A rule that only sets a fill
 * does not clear a colour an earlier rule set, which is what makes
 * "red text for negatives" and "grey fill for the weekend" two rules
 * rather than four. Excel stops at the first match unless a rule says
 * otherwise; this is the other convention and the one that composes,
 * and it is written down here because it is the kind of thing that is
 * otherwise discovered.
 */
export function overlay(paints: readonly ConditionalPaint[]): ConditionalPaint {
  const folded: {
    fill?: string;
    color?: string;
    bold?: boolean;
    italic?: boolean;
  } = {};
  for (const paint of paints) {
    if (paint.fill !== undefined) {
      folded.fill = paint.fill;
    }
    if (paint.color !== undefined) {
      folded.color = paint.color;
    }
    if (paint.bold !== undefined) {
      folded.bold = paint.bold;
    }
    if (paint.italic !== undefined) {
      folded.italic = paint.italic;
    }
  }
  return folded;
}

/** Whether an overlay would change anything at all. */
export function isEmptyPaint(paint: ConditionalPaint): boolean {
  return (
    paint.fill === undefined && paint.color === undefined && paint.bold === undefined && paint.italic === undefined
  );
}
