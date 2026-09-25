import { callAt, scanFormula, wordAt } from './FormulaScan';
import { completionsFor, signatureOf, type Signature } from './Signatures';
import type { Span } from './Tokenizer';

/**
 * What to offer somebody in the middle of typing a formula.
 *
 * Two things, and only ever one at a time.
 *
 * **A list of names**, while a bare word is being typed. That is the
 * moment a completion helps and the only one: offering a list while
 * the caret sits inside a finished call would cover the formula with
 * something nobody asked for.
 *
 * **A signature**, once the `(` is typed and the arguments are being
 * filled in, with the argument the caret is in marked. That is the
 * question people actually have — *which* argument is this? — and it
 * is the one a list of names cannot answer.
 *
 * The list wins when both apply, because a word being typed is a more
 * specific state than being inside a call: `=SUM(RO` is somebody
 * typing `ROUND`, not somebody wondering about `SUM`'s first
 * argument.
 */
export type FormulaHint =
  | {
      readonly kind: 'completions';
      /** The word so far, and where it is, so accepting one can replace it. */
      readonly prefix: string;
      readonly span: Span;
      readonly names: readonly string[];
    }
  | {
      readonly kind: 'signature';
      readonly name: string;
      readonly signature: Signature;
      /** Which argument the caret is in, counting from zero. */
      readonly argument: number;
    }
  | null;

export function hintFor(text: string, caret: number): FormulaHint {
  if (!text.startsWith('=')) {
    return null;
  }
  const scan = scanFormula(text);

  const word = wordAt(scan, caret);
  if (word !== null) {
    const prefix = text.slice(word.start, word.end);
    const names = completionsFor(prefix);
    // A word that matches nothing is somebody typing something else —
    // a name the sheet does not know, or a word inside a string. An
    // empty popup is worse than none.
    if (names.length > 0) {
      return { kind: 'completions', prefix, span: word, names };
    }
    return null;
  }

  const call = callAt(scan, caret);
  if (call === null) {
    return null;
  }
  const signature = signatureOf(call.name);
  if (signature === null) {
    return null;
  }
  return { kind: 'signature', name: call.name, signature, argument: call.argument };
}

/**
 * The text and caret after accepting a completion.
 *
 * The `(` is typed for you, because a function name without one is
 * not a call and every completion is about to need it. The caret
 * lands inside the brackets, which is where the next thing typed
 * goes.
 */
export function acceptCompletion(text: string, span: Span, name: string): { text: string; caret: number } {
  const written = `${text.slice(0, span.start)}${name}(${text.slice(span.end)}`;
  return { text: written, caret: span.start + name.length + 1 };
}

/**
 * Which argument of a signature to mark, given how many commas have
 * been typed.
 *
 * A signature whose last argument repeats marks that last one for
 * every argument past it, rather than running out of things to point
 * at: the fourth argument of `SUM(number, …)` is still a number.
 */
export function markedArgument(signature: Signature, argument: number): number {
  if (signature.args.length === 0) {
    return -1;
  }
  if (argument < signature.args.length) {
    return argument;
  }
  return signature.repeats === true ? signature.args.length - 1 : -1;
}
