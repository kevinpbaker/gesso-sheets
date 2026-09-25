import { columnIndex, parseRef, type CellRef, type RangeRef } from './A1';
import { FormulaSyntaxError, tokenize, type Span, type Token } from './Tokenizer';

/**
 * A formula as the thing somebody is *typing*, rather than as a tree.
 *
 * The parser answers "what does this mean", and for a formula being
 * edited that is the wrong question: half the time it means nothing
 * yet, because the caret is sitting after `=SUM(` and the user has not
 * finished. Every part of the formula editor needs answers about
 * unfinished text — which reference to colour, which argument of which
 * call the caret is in, which bracket matches, whether a click should
 * insert an address — and an AST that refuses to exist until the
 * formula is well formed cannot give them.
 *
 * So this scans tokens, not trees. It never throws: a formula that
 * cannot even be tokenized scans as nothing, because a red squiggle is
 * not worth an exception on every keystroke.
 *
 * Headless, like everything else here. The decisions it supports are
 * the application's; the reading of the text is the engine's.
 */

/** A reference in the text, and where it was written. */
export interface ScannedReference extends Span {
  /**
   * The cells it names — one, or the two corners of a range.
   *
   * Named `from`/`to` rather than `start`/`end` because a scanned
   * reference is two things at once: a span of *text*, whose start and
   * end are character offsets, and a span of *cells*. Using one pair
   * of names for both is how an offset ends up compared against a row
   * number.
   */
  readonly from: CellRef;
  readonly to: CellRef;
  /** True when the text was `A1:B9` rather than `A1`. */
  readonly isRange: boolean;
}

/** Where the caret is, in terms of the call it sits inside. */
export interface CallContext {
  /** The function name as typed, upper-cased. */
  readonly name: string;
  /** Which argument the caret is in, counting from zero. */
  readonly argument: number;
  /** Where the name itself was written, for a popup to point at. */
  readonly nameSpan: Span;
}

export interface FormulaScan {
  readonly tokens: readonly Token[];
  readonly references: readonly ScannedReference[];
}

/**
 * The tokens and the references, or nothing when the text will not
 * tokenize at all.
 *
 * `=1+"` is unclosed and scans as nothing rather than throwing,
 * because this runs on every keystroke and half of every formula is
 * unfinished while it is being written.
 */
export function scanFormula(text: string): FormulaScan {
  if (!text.startsWith('=')) {
    return EMPTY;
  }
  let tokens: Token[];
  try {
    // Scanned from index 1 and shifted back, so every offset this file
    // hands out is an offset into the text the *editor* holds —
    // including the `=`. An off-by-one here would colour the wrong
    // characters, which is the kind of bug that looks like a font
    // problem.
    tokens = tokenize(text.slice(1)).map(token => ({ ...token, start: token.start + 1, end: token.end + 1 }));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    return EMPTY;
  }
  return { tokens, references: referencesIn(tokens) };
}

const EMPTY: FormulaScan = { tokens: [], references: [] };

/**
 * The references in a token list, ranges kept whole.
 *
 * A range is three tokens — word, colon, word — and has to be found
 * here rather than by looking at each word alone, or `A1:B9` colours
 * as two references with a gap in the middle.
 */
function referencesIn(tokens: readonly Token[]): ScannedReference[] {
  const found: ScannedReference[] = [];
  for (let at = 0; at < tokens.length; at++) {
    const token = tokens[at];
    if (token.kind !== 'word') {
      continue;
    }
    // A word followed by `(` is a call, not a cell: `LOG10(2)`.
    if (tokens[at + 1]?.kind === 'open') {
      continue;
    }
    const ref = parseRef(token.value);
    const colon = tokens[at + 1];
    const after = tokens[at + 2];

    if (colon?.kind === 'colon' && after?.kind === 'word') {
      const end = parseRef(after.value);
      if (ref !== null && end !== null) {
        found.push({ from: ref, to: end, isRange: true, start: token.start, end: after.end });
        at += 2;
        continue;
      }
      const columns = wholeColumns(token.value, after.value);
      if (columns !== null) {
        found.push({ ...columns, isRange: true, start: token.start, end: after.end });
        at += 2;
        continue;
      }
    }
    if (ref !== null) {
      found.push({ from: ref, to: ref, isRange: false, start: token.start, end: token.end });
    }
  }
  return found;
}

/** `A:C`, which names columns rather than cells. */
function wholeColumns(first: string, last: string): { from: CellRef; to: CellRef } | null {
  const from = columnIndex(first.replace(/\$/g, ''));
  const to = columnIndex(last.replace(/\$/g, ''));
  if (from === null || to === null) {
    return null;
  }
  return {
    from: { row: 0, column: from, rowAbsolute: true, columnAbsolute: false },
    to: { row: 0, column: to, rowAbsolute: true, columnAbsolute: false }
  };
}

/** The reference the caret is inside or touching, if any. */
export function referenceAt(scan: FormulaScan, caret: number): ScannedReference | null {
  for (const reference of scan.references) {
    if (caret >= reference.start && caret <= reference.end) {
      return reference;
    }
  }
  return null;
}

/**
 * The call the caret is inside, innermost first, and which argument.
 *
 * Counted by walking the tokens and keeping a stack: every `(`
 * preceded by a word opens a call, every `)` closes one, and a comma
 * at the top of the stack advances the argument. Unclosed brackets are
 * the normal case rather than an error — the caret is usually inside
 * one that has not been typed yet.
 */
export function callAt(scan: FormulaScan, caret: number): CallContext | null {
  const stack: { name: string; argument: number; nameSpan: Span }[] = [];
  for (let at = 0; at < scan.tokens.length; at++) {
    const token = scan.tokens[at];
    if (token.start >= caret) {
      break;
    }
    if (token.kind === 'open') {
      const before = scan.tokens[at - 1];
      stack.push(
        before !== undefined && before.kind === 'word'
          ? { name: before.value.toUpperCase(), argument: 0, nameSpan: { start: before.start, end: before.end } }
          : // A bare `(` for grouping. Pushed anyway, so that the
            // brackets stay balanced and a comma inside it is not
            // counted as an argument of the call outside it.
            { name: '', argument: 0, nameSpan: { start: token.start, end: token.start } }
      );
      continue;
    }
    if (token.kind === 'close') {
      stack.pop();
      continue;
    }
    if (token.kind === 'comma' && stack.length > 0) {
      stack[stack.length - 1].argument++;
    }
  }
  for (let at = stack.length - 1; at >= 0; at--) {
    if (stack[at].name !== '') {
      return stack[at];
    }
  }
  return null;
}

/**
 * The bracket matching the one beside the caret, or null.
 *
 * Looks at the character before the caret first and then the one
 * after, which is what every editor does and what makes the highlight
 * appear when you have just typed a `)` as well as when you have
 * arrowed onto one.
 */
export function matchingBracket(scan: FormulaScan, caret: number): { here: Span; there: Span } | null {
  const before = scan.tokens.find(token => token.end === caret && (token.kind === 'open' || token.kind === 'close'));
  const after = scan.tokens.find(token => token.start === caret && (token.kind === 'open' || token.kind === 'close'));
  const bracket = before ?? after;
  if (bracket === undefined) {
    return null;
  }
  const forwards = bracket.kind === 'open';
  const step = forwards ? 1 : -1;
  const from = scan.tokens.indexOf(bracket);
  let depth = 0;
  for (let at = from; at >= 0 && at < scan.tokens.length; at += step) {
    const token = scan.tokens[at];
    if (token.kind === 'open') {
      depth += forwards ? 1 : -1;
    } else if (token.kind === 'close') {
      depth += forwards ? -1 : 1;
    } else {
      continue;
    }
    if (depth === 0) {
      return {
        here: { start: bracket.start, end: bracket.end },
        there: { start: token.start, end: token.end }
      };
    }
  }
  return null;
}

/**
 * The bare word the caret is at the end of, if there is one.
 *
 * What a function name looks like while it is being typed: a word
 * with the caret against its right-hand edge and no `(` after it.
 * A word the caret merely sits inside does not count — somebody
 * editing the middle of `SUM` is not asking for a list of names —
 * and neither does one already turned into a call.
 *
 * Words that are cell references are skipped, or typing `A1` would
 * offer to complete it to nothing and `B` would offer every function
 * beginning with B while somebody is typing an address.
 */
export function wordAt(scan: FormulaScan, caret: number): Span | null {
  for (let at = 0; at < scan.tokens.length; at++) {
    const token = scan.tokens[at];
    if (token.kind !== 'word' || token.end !== caret) {
      continue;
    }
    if (scan.tokens[at + 1]?.kind === 'open') {
      return null;
    }
    return parseRef(token.value) === null ? { start: token.start, end: token.end } : null;
  }
  return null;
}

/** The range a scanned reference covers, for outlining it in the grid. */
export function rangeOf(reference: ScannedReference): RangeRef {
  return { start: reference.from, end: reference.to };
}
