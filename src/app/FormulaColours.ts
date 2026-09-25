import type { UiTextSpan } from 'gesso-core';

import type { RangeRef } from '../sheet/A1';
import { scanFormula } from '../sheet/FormulaScan';
import type { Span } from '../sheet/Tokenizer';

/**
 * The colours a formula's references are drawn in, in the text and in
 * the grid.
 *
 * One list feeds both, which is the whole point: the `B2` in the text
 * and the box drawn round B2 on the sheet have to be the same colour
 * or the feature is worse than not having it. A second palette
 * computed somewhere else would drift the first time either changed.
 *
 * ## Why these are hex and not theme tokens
 *
 * The theme's colours are *roles* — `primary`, `danger`, `textMuted` —
 * and there is no role here to borrow. What is wanted is five hues
 * that are told apart at a glance and carry no meaning, which is
 * exactly what a role is not. Borrowing `danger` for the second
 * reference would say something false about it.
 *
 * They are chosen to stay legible on a light ground and a dark one,
 * which is a compromise: on a dark theme they are a little brighter
 * than ideal and on a light one a little darker. A theme that wanted
 * its own set would be a reasonable thing to add later.
 */
const PALETTE: readonly string[] = ['#4285f4', '#ea4335', '#34a853', '#b061f5', '#fa7b17'];

export interface ColouredReference {
  /** Where in the formula text it was written. */
  readonly span: Span;
  /** The cells it names, for drawing a box round them. */
  readonly range: RangeRef;
  readonly color: string;
}

/**
 * The references in a formula, each with the colour it is shown in.
 *
 * Two references to the same cells get the same colour however they
 * were written, so `=A1+$A$1` is one colour twice rather than two —
 * the dollars change how the reference *copies*, not what it points
 * at, and colouring them differently would suggest otherwise.
 *
 * The palette repeats after five. A formula with six distinct
 * references has two of them the same colour, which is worse than
 * six distinct hues and better than five hues plus one nobody can
 * distinguish from a neighbour.
 */
export function colouredReferences(text: string): readonly ColouredReference[] {
  const colours = new Map<string, string>();
  return scanFormula(text).references.map(reference => {
    const key = `${reference.from.row}:${reference.from.column}:${reference.to.row}:${reference.to.column}`;
    let color = colours.get(key);
    if (color === undefined) {
      color = PALETTE[colours.size % PALETTE.length];
      colours.set(key, color);
    }
    return {
      span: { start: reference.start, end: reference.end },
      range: { start: reference.from, end: reference.to },
      color
    };
  });
}

/**
 * A formula as runs of text, for an `EditableText` to draw.
 *
 * The runs **tile the whole string**, because runs supply a
 * paragraph's text: the gaps between the references are runs too,
 * carrying no colour and therefore inheriting the field's. The engine
 * checks that their concatenation equals the value and draws plainly
 * when it does not, so a mistake here is a colourless field rather
 * than a field whose caret is in the wrong place — but it would still
 * be a mistake, and this is the one function that must not make it.
 *
 * Undefined rather than a single plain run when there is nothing to
 * colour, so an ordinary cell being edited costs nothing at all.
 */
export function formulaSpans(text: string): readonly UiTextSpan[] | undefined {
  const references = colouredReferences(text);
  if (references.length === 0) {
    return undefined;
  }
  const runs: UiTextSpan[] = [];
  let at = 0;
  for (const reference of references) {
    if (reference.span.start > at) {
      runs.push({ text: text.slice(at, reference.span.start) });
    }
    runs.push({ text: text.slice(reference.span.start, reference.span.end), color: reference.color });
    at = reference.span.end;
  }
  if (at < text.length) {
    runs.push({ text: text.slice(at) });
  }
  return runs;
}
