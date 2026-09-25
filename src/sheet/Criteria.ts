import { compareValues, isError, toNumber, type CellValue } from './Values';

/**
 * The little language inside `SUMIF(A:A, ">10", B:B)`.
 *
 * It is a language, however small, and it is shared by six functions —
 * which is the reason it is a file rather than a helper hidden inside
 * one of them. Six copies of "does this cell match" is six chances to
 * disagree about whether `"<>"` matches a blank.
 *
 * The grammar:
 *
 *   - A leading `>=`, `<=`, `<>`, `>`, `<` or `=` is a comparison
 *     against what follows it.
 *   - Anything else is a test for equality with the whole string.
 *   - The operand is read as a number when it looks like one, so
 *     `">10"` compares numerically and `">abc"` compares as text.
 *   - Text comparison is case-insensitive, as everything else in this
 *     engine's comparisons is.
 *   - `*` and `?` in an equality test are wildcards — any run of
 *     characters and any one character — because `COUNTIF(A:A, "N*")`
 *     is how people count the names beginning with N, and a literal
 *     reading of that would return zero and look like a broken sheet.
 *     `~` escapes them.
 *
 * A criterion that is not text at all — `COUNTIF(A:A, 5)` — is an
 * equality test against that value, which is the common case and does
 * not go near the parser.
 */

export interface Criterion {
  matches(value: CellValue): boolean;
}

type Comparison = '>=' | '<=' | '<>' | '>' | '<' | '=';

const COMPARISONS: readonly Comparison[] = ['>=', '<=', '<>', '>', '<', '='];

export function criterionOf(raw: CellValue): Criterion {
  if (typeof raw !== 'string') {
    // A number, a boolean or a blank: equality, and no grammar.
    return { matches: value => equal(value, raw) };
  }

  const operator = COMPARISONS.find(candidate => raw.startsWith(candidate)) ?? null;
  const operand = operator === null ? raw : raw.slice(operator.length);
  const wanted = operandValue(operand);

  if (operator === null || operator === '=' || operator === '<>') {
    const pattern = hasWildcards(operand) ? wildcardPattern(operand) : null;
    const test = (value: CellValue): boolean =>
      pattern === null ? equal(value, wanted) : pattern.test(textOf(value));
    return operator === '<>' ? { matches: value => !test(value) } : { matches: test };
  }

  return {
    matches(value) {
      const order = compareValues(value, wanted);
      if (isError(order)) {
        return false;
      }
      switch (operator) {
        case '>':
          return order > 0;
        case '>=':
          return order >= 0;
        case '<':
          return order < 0;
        default:
          return order <= 0;
      }
    }
  };
}

/**
 * The text after the operator, as the value it stands for.
 *
 * `">10"` has to compare numerically or every numeric criterion in
 * every sheet is a string comparison, in which "9" is greater than
 * "10".
 */
function operandValue(text: string): CellValue {
  const trimmed = text.trim();
  if (trimmed === '') {
    // `"<>"` on its own: not equal to blank, which is how people
    // count the cells that have anything in them.
    return null;
  }
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') {
    return true;
  }
  if (upper === 'FALSE') {
    return false;
  }
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }
  return text;
}

/**
 * Equality as a criterion means it, which is not the same as `=`.
 *
 * A blank cell equals a blank criterion and equals nothing else —
 * not zero, not the empty string. `SUMIF(A:A, 0)` over a column of
 * empty cells must be zero rather than the whole column, which is
 * what treating a blank as zero here would give.
 */
function equal(value: CellValue, wanted: CellValue): boolean {
  if (isError(wanted)) {
    // A criterion that is itself an error matches the same error,
    // which is how `COUNTIF(A:A, "#N/A")`'s cousin behaves and is the
    // only reading that is not simply "never".
    return isError(value) && value.code === wanted.code;
  }
  if (wanted === null) {
    return value === null;
  }
  if (value === null) {
    return false;
  }
  if (typeof wanted === 'number') {
    // A boolean is not a number here, whatever `toNumber` says of it:
    // `COUNTIF(A:A, 1)` must not count the TRUEs.
    if (typeof value === 'boolean') {
      return false;
    }
    const number = toNumber(value);
    return !isError(number) && number === wanted;
  }
  if (typeof wanted === 'boolean') {
    return value === wanted;
  }
  return typeof value === 'string' && value.toUpperCase() === wanted.toUpperCase();
}

function textOf(value: CellValue): string {
  if (value === null) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  return isError(value) ? value.code : String(value);
}

/**
 * Whether a criterion is a pattern rather than a literal.
 *
 * True for an unescaped `*` or `?`, and *also* for an escaped one —
 * `"N~*"` has no wildcard left once the tilde is read, but it still
 * has to go through the pattern builder, because that is what strips
 * the tilde. Answering false there compared the criterion to the
 * cell tilde and all, and matched nothing.
 */
function hasWildcards(text: string): boolean {
  for (let at = 0; at < text.length; at++) {
    if (text[at] === '~' && escapable(text[at + 1])) {
      return true;
    }
    if (text[at] === '*' || text[at] === '?') {
      return true;
    }
  }
  return false;
}

/** The three characters a tilde can escape; before anything else it is one. */
function escapable(character: string | undefined): boolean {
  return character === '*' || character === '?' || character === '~';
}

/** `N*` as a regular expression, with `~` escaping a literal star. */
function wildcardPattern(text: string): RegExp {
  let source = '';
  for (let at = 0; at < text.length; at++) {
    const character = text[at];
    if (character === '~' && escapable(text[at + 1])) {
      source += escapeLiteral(text[++at]);
      continue;
    }
    if (character === '*') {
      source += '.*';
      continue;
    }
    if (character === '?') {
      source += '.';
      continue;
    }
    source += escapeLiteral(character);
  }
  return new RegExp(`^${source}$`, 'i');
}

function escapeLiteral(character: string): string {
  return /[\\^$.*+?()[\]{}|]/.test(character) ? `\\${character}` : character;
}
