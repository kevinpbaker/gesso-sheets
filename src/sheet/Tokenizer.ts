import type { ErrorCode } from './Values';

/**
 * A formula's characters, as pieces the parser can read.
 *
 * One decision worth knowing: a bare word is a `word` token and
 * nothing more. `SUM` and `A1` and `LOG10` all arrive the same way,
 * because `LOG10` is a perfectly good cell address and a perfectly
 * good function name and only what follows it says which — a `(`
 * makes it a call, anything else makes it a reference. Deciding that
 * here would need the lexer to look ahead past whitespace, which is
 * the parser's job and where it already happens.
 */
/**
 * Where a token sat in the text it came from.
 *
 * Carried on every token because Phase 12 is built out of the
 * question "what is under the caret" — which reference to colour,
 * which argument of which call the caret is inside, which bracket
 * matches the one beside it. None of that can be answered from a
 * token list that has forgotten where the tokens were, and
 * re-deriving the positions by re-scanning the text is a second
 * scanner to disagree with this one.
 *
 * `start` is the first character, `end` is one past the last, so
 * `source.slice(start, end)` is the token.
 */
export interface Span {
  readonly start: number;
  readonly end: number;
}

export type Token = Span &
  (
    | { readonly kind: 'number'; readonly value: number }
    | { readonly kind: 'text'; readonly value: string }
    | { readonly kind: 'word'; readonly value: string }
    /**
     * A sheet named in front of a reference: the `Sheet2` of
     * `Sheet2!A1`, and the `Q3 Budget` of `'Q3 Budget'!A1:B9`.
     *
     * Read here rather than in the parser because the `!` is what
     * makes it one. Without the `!` a quoted name is not a token this
     * grammar has at all, and `Q3 Budget` is two words — so the
     * decision needs the character after the name, which is exactly
     * what a scanner has and a parser of already-scanned tokens does
     * not.
     */
    | { readonly kind: 'sheet'; readonly value: string }
    | { readonly kind: 'error'; readonly code: ErrorCode }
    | { readonly kind: 'operator'; readonly value: string }
    | { readonly kind: 'open' }
    | { readonly kind: 'close' }
    | { readonly kind: 'comma' }
    | { readonly kind: 'colon' }
    | { readonly kind: 'end' }
  );

export class FormulaSyntaxError extends Error {}

const ERROR_LITERALS: readonly ErrorCode[] = ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#CIRC!', '#N/A'];

const TWO_CHARACTER_OPERATORS = new Set(['<>', '<=', '>=']);
const ONE_CHARACTER_OPERATORS = new Set(['+', '-', '*', '/', '^', '&', '=', '<', '>']);

export function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let at = 0;

  while (at < source.length) {
    const character = source[at];

    if (character === ' ' || character === '\t' || character === '\n' || character === '\r') {
      at++;
      continue;
    }

    if (character === '(') {
      tokens.push({ kind: 'open', start: at, end: at + 1 });
      at++;
      continue;
    }
    if (character === ')') {
      tokens.push({ kind: 'close', start: at, end: at + 1 });
      at++;
      continue;
    }
    if (character === ',') {
      tokens.push({ kind: 'comma', start: at, end: at + 1 });
      at++;
      continue;
    }
    if (character === ':') {
      tokens.push({ kind: 'colon', start: at, end: at + 1 });
      at++;
      continue;
    }

    if (character === '"') {
      const { value, next } = readText(source, at);
      tokens.push({ kind: 'text', value, start: at, end: next });
      at = next;
      continue;
    }

    if (character === "'") {
      const { value, next } = readQuotedSheet(source, at);
      tokens.push({ kind: 'sheet', value, start: at, end: next });
      at = next;
      continue;
    }

    if (character === '#') {
      const literal = ERROR_LITERALS.find(code => source.startsWith(code, at));
      if (literal === undefined) {
        throw new FormulaSyntaxError(`Unknown error value at ${at}.`);
      }
      tokens.push({ kind: 'error', code: literal, start: at, end: at + literal.length });
      at += literal.length;
      continue;
    }

    if (isDigit(character) || (character === '.' && isDigit(source[at + 1] ?? ''))) {
      const { value, next } = readNumber(source, at);
      tokens.push({ kind: 'number', value, start: at, end: next });
      at = next;
      continue;
    }

    if (isWordStart(character)) {
      let end = at;
      while (end < source.length && isWordPart(source[end])) {
        end++;
      }
      // A word with a `!` against it names a sheet. Nothing else in
      // this grammar uses the character, so there is nothing for it
      // to be mistaken for.
      if (source[end] === '!') {
        tokens.push({ kind: 'sheet', value: source.slice(at, end), start: at, end: end + 1 });
        at = end + 1;
        continue;
      }
      tokens.push({ kind: 'word', value: source.slice(at, end), start: at, end });
      at = end;
      continue;
    }

    const two = source.slice(at, at + 2);
    if (TWO_CHARACTER_OPERATORS.has(two)) {
      tokens.push({ kind: 'operator', value: two, start: at, end: at + 2 });
      at += 2;
      continue;
    }
    if (ONE_CHARACTER_OPERATORS.has(character)) {
      tokens.push({ kind: 'operator', value: character, start: at, end: at + 1 });
      at++;
      continue;
    }

    throw new FormulaSyntaxError(`Unexpected character ${JSON.stringify(character)} at ${at}.`);
  }

  tokens.push({ kind: 'end', start: source.length, end: source.length });
  return tokens;
}

/** A quoted string, in which `""` is one quote. */
function readText(source: string, start: number): { value: string; next: number } {
  let at = start + 1;
  let value = '';
  while (at < source.length) {
    if (source[at] === '"') {
      if (source[at + 1] === '"') {
        value += '"';
        at += 2;
        continue;
      }
      return { value, next: at + 1 };
    }
    value += source[at];
    at++;
  }
  throw new FormulaSyntaxError('A quoted string was never closed.');
}

/**
 * `'Q3 Budget'!`, in which `''` is one quote.
 *
 * The `!` is required and consumed. A quoted name with nothing after
 * it is not a sheet and not anything else, and saying so here is
 * better than handing the parser a token it cannot place.
 */
function readQuotedSheet(source: string, start: number): { value: string; next: number } {
  let at = start + 1;
  let value = '';
  while (at < source.length) {
    if (source[at] === "'") {
      if (source[at + 1] === "'") {
        value += "'";
        at += 2;
        continue;
      }
      if (source[at + 1] !== '!') {
        throw new FormulaSyntaxError(`A sheet name has to be followed by '!' at ${at + 1}.`);
      }
      if (value === '') {
        throw new FormulaSyntaxError(`A sheet name cannot be empty at ${start}.`);
      }
      return { value, next: at + 2 };
    }
    value += source[at];
    at++;
  }
  throw new FormulaSyntaxError('A quoted sheet name was never closed.');
}

function readNumber(source: string, start: number): { value: number; next: number } {
  let at = start;
  while (at < source.length && isDigit(source[at])) {
    at++;
  }
  if (source[at] === '.') {
    at++;
    while (at < source.length && isDigit(source[at])) {
      at++;
    }
  }
  // An exponent, but only when it is really one: `1E5` is a number and
  // `A1E5` never reaches here, while `1E` followed by nothing is the
  // number 1 beside a word, which is a syntax error one level up
  // rather than a malformed number here.
  if (source[at] === 'e' || source[at] === 'E') {
    const afterSign = source[at + 1] === '+' || source[at + 1] === '-' ? at + 2 : at + 1;
    if (isDigit(source[afterSign] ?? '')) {
      at = afterSign;
      while (at < source.length && isDigit(source[at])) {
        at++;
      }
    }
  }
  const text = source.slice(start, at);
  const value = Number(text);
  if (Number.isNaN(value)) {
    throw new FormulaSyntaxError(`${text} is not a number.`);
  }
  return { value, next: at };
}

function isDigit(character: string): boolean {
  return character >= '0' && character <= '9';
}

function isWordStart(character: string): boolean {
  return /[A-Za-z_$]/.test(character);
}

function isWordPart(character: string): boolean {
  return /[A-Za-z0-9_$.]/.test(character);
}
