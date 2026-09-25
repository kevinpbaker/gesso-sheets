import { formatRange, formatRef, relativeRef, type CellRef, type RangeRef } from '../sheet/A1';
import { referenceAt, scanFormula, type ScannedReference } from '../sheet/FormulaScan';
import type { Span } from '../sheet/Tokenizer';

/**
 * What a click in the grid means while a formula is open.
 *
 * **This is a mode, and modes are where spreadsheets keep their worst
 * bugs.** The same click either moves the selection or types an
 * address into somebody's formula, and which one it does depends on
 * where the caret is sitting. Get it wrong in one direction and a
 * click during an edit throws the formula away; get it wrong in the
 * other and the selection freezes and the sheet feels broken.
 *
 * So it is a function of (text, caret) with a table of cases beside
 * it in `FormulaEditing.spec.ts`, and not a series of conditions
 * inside a pointer handler. The handler asks this and does what it
 * says.
 *
 * Nothing here imports the framework. It is application logic — the
 * same shape `SheetKeys` has, and for the same reason.
 */

/** What a click should do, given where the caret is. */
export type PickDecision =
  /** Put an address in at the caret. */
  | { readonly kind: 'insert' }
  /** Put an address in *over* the reference the caret is on. */
  | { readonly kind: 'replace'; readonly span: Span }
  /** Not a picking position: the click means what it always means. */
  | { readonly kind: 'select' };

/** The text after a pick, and where the caret and the address ended up. */
export interface Picked {
  readonly text: string;
  readonly caret: number;
  /** Where the address was written, so a drag can rewrite it in place. */
  readonly span: Span;
}

/**
 * The characters a reference may follow.
 *
 * After any of these, the formula is expecting a value and an address
 * is a legal thing to type — so a click puts one in. After a number,
 * a closing bracket or a quoted string it is not: `=1` followed by a
 * click means the person is done and wants to go somewhere, and
 * inserting there would produce `=1A1`.
 */
const OPENS_A_VALUE = new Set(['=', '+', '-', '*', '/', '^', '&', '(', ',', ':', '<', '>', '<=', '>=', '<>']);

/**
 * Whether a click picks a reference, and what it would do to the text.
 *
 * Three answers, and the boundary between them is the whole design.
 *
 *   - The caret is **on a reference** — `=SUM(A1)` with the caret
 *     anywhere in or beside `A1`. Clicking replaces it, which is how
 *     somebody fixes a formula that points at the wrong cell.
 *   - The caret **follows something that wants a value** — after `=`,
 *     an operator, `(` or `,`. Clicking inserts.
 *   - Anything else. Clicking is a click: it commits the edit and
 *     moves the selection, as it would with nothing open.
 */
export function pickDecision(text: string, caret: number): PickDecision {
  if (!text.startsWith('=')) {
    return { kind: 'select' };
  }
  const scan = scanFormula(text);
  const reference = referenceAt(scan, caret);
  if (reference !== null && !endsARange(text, reference, caret)) {
    return { kind: 'replace', span: { start: reference.start, end: reference.end } };
  }
  return opensAValue(text, caret) ? { kind: 'insert' } : { kind: 'select' };
}

/**
 * The text before the caret, ignoring spaces, as the last thing typed.
 *
 * Two characters are read as well as one, so `<=` is not mistaken for
 * a bare `=` — which would be harmless here, both being in the set,
 * and is the kind of near-miss worth not relying on.
 */
function opensAValue(text: string, caret: number): boolean {
  let at = caret;
  while (at > 0 && text[at - 1] === ' ') {
    at--;
  }
  if (at === 0) {
    return false;
  }
  // `Math.max` rather than `at - 2`: a negative start makes `slice`
  // count from the end of the string, which happens to give the empty
  // string here and would be a silent wrong answer somewhere else.
  return OPENS_A_VALUE.has(text.slice(Math.max(0, at - 2), at)) || OPENS_A_VALUE.has(text[at - 1]);
}

/**
 * A caret just after `A1:` is not on the reference, it is past it.
 *
 * `=SUM(A1:` with the caret at the end is somebody halfway through
 * typing a range, and the next click should extend it rather than
 * replace the corner they have already chosen. Without this the colon
 * case reads as "on a reference" and eats what they typed.
 */
function endsARange(text: string, reference: ScannedReference, caret: number): boolean {
  return caret === reference.end && text[caret] === ':';
}

/** The address a cell or a range is written as when a click puts it in. */
export function addressOf(range: RangeRef): string {
  const single = range.start.row === range.end.row && range.start.column === range.end.column;
  // Relative, always. A click is not a statement about copying, and
  // every spreadsheet writes a picked reference without dollars.
  const start = relativeRef(range.start.row, range.start.column);
  const end = relativeRef(range.end.row, range.end.column);
  return single ? formatRef(start) : formatRange({ start, end });
}

/** Writes an address into the text, by whichever of the two ways applies. */
export function pick(text: string, caret: number, range: RangeRef): Picked | null {
  const decision = pickDecision(text, caret);
  if (decision.kind === 'select') {
    return null;
  }
  const span = decision.kind === 'replace' ? decision.span : { start: caret, end: caret };
  return replaceSpan(text, span, addressOf(range));
}

/**
 * Rewrites an address already written, which is what a drag does.
 *
 * The pointer going down inserts `B2`; every move after it replaces
 * that same span with `B2:D7`, so the formula grows rather than
 * accumulating corners. The span comes back from the previous call,
 * which is why `Picked` carries one.
 */
export function repick(text: string, span: Span, range: RangeRef): Picked {
  return replaceSpan(text, span, addressOf(range));
}

function replaceSpan(text: string, span: Span, address: string): Picked {
  const written = text.slice(0, span.start) + address + text.slice(span.end);
  const end = span.start + address.length;
  return { text: written, caret: end, span: { start: span.start, end } };
}

/**
 * F4: `A1 → $A$1 → A$1 → $A1 → A1`.
 *
 * The order is Excel's and is not the order anybody would choose —
 * "both, then the row, then the column" — but it is the one people's
 * fingers know, and a cycle that went the other way round would be
 * wrong in the only way that matters.
 *
 * A range cycles both corners together, because a range with one
 * pinned end is a thing almost nobody means and always a surprise
 * when it happens by accident.
 *
 * Null when the caret is not on a reference, which the caller shows
 * by doing nothing rather than by beeping.
 */
export function cycleAbsolute(text: string, caret: number): Picked | null {
  const reference = referenceAt(scanFormula(text), caret);
  if (reference === null) {
    return null;
  }
  const next = nextAnchoring(reference.from);
  const from = { ...reference.from, ...next };
  const to = { ...reference.to, ...next };
  const address = reference.isRange ? formatRange({ start: from, end: to }) : formatRef(from);
  return replaceSpan(text, { start: reference.start, end: reference.end }, address);
}

/** One step round the cycle, in terms of which halves are pinned. */
function nextAnchoring(ref: CellRef): { rowAbsolute: boolean; columnAbsolute: boolean } {
  if (!ref.rowAbsolute && !ref.columnAbsolute) {
    return { rowAbsolute: true, columnAbsolute: true };
  }
  if (ref.rowAbsolute && ref.columnAbsolute) {
    return { rowAbsolute: true, columnAbsolute: false };
  }
  if (ref.rowAbsolute) {
    return { rowAbsolute: false, columnAbsolute: true };
  }
  return { rowAbsolute: false, columnAbsolute: false };
}
