import { LIFTED } from './Functions';
import { tokenize, type Token } from './Tokenizer';

/**
 * A formula from before dynamic arrays, made to say what it meant.
 *
 * In an `.xlsx` written before Excel 365, `=B1:B9*2` in C5 means `B5*2`:
 * a range where one value is wanted gives the one in the formula's own
 * row, which Excel calls implicit intersection. This sheet spills that
 * formula instead, as Excel 365 does — so a formula read from such a
 * file is given the `@` that Excel 365 itself shows when it opens one,
 * wherever the difference would show.
 *
 * Where it is not needed, it is not added:
 *
 * - a range that is a whole argument to a function is a range to that
 *   function, which takes its one value itself if one is all it wants —
 *   `SUM(A1:A9)`, `ABS(A1:A9)`. `IF` is the exception: its test is
 *   taken cell by cell over an array, so a range there is marked;
 * - everything inside `SUMPRODUCT`, which has always worked on whole
 *   arrays, in every version;
 * - a single cell, which is one value already.
 *
 * Done on the text, for the reason renaming a sheet is: printing a
 * tree back would lose the spaces and brackets somebody typed.
 */
export function withIntersections(input: string, names: ReadonlySet<string>): string {
  if (!input.startsWith('=')) {
    return input;
  }
  const body = input.slice(1);
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return input;
  }
  const inserts: number[] = [];
  /**
   * What each open bracket belongs to: a call's name, `@` for `@(…)`,
   * or null for grouping.
   */
  const stack: (string | null)[] = [];
  // Inside `@(…)` the whole is one value already, and inside
  // `SUMPRODUCT` everything is an array on purpose.
  const insideArrays = (): boolean => stack.some(name => name === 'SUMPRODUCT' || name === '@');

  const mark = (first: number, last: number): void => {
    if (insideArrays()) {
      return;
    }
    const before = tokens[first - 1];
    const after = tokens[last + 1];
    const call = stack[stack.length - 1];
    const whole =
      call !== undefined &&
      call !== null &&
      call !== 'IF' &&
      // A function of one value runs once per cell over a range now, so
      // one written when it took one value from it has to say so.
      !LIFTED.has(call) &&
      (before?.kind === 'open' || before?.kind === 'comma') &&
      (after?.kind === 'comma' || after?.kind === 'close');
    if (!whole && !(before?.kind === 'operator' && (before as { value: string }).value === '@')) {
      inserts.push(tokens[first].start);
    }
  };

  for (let at = 0; at < tokens.length; at++) {
    const token = tokens[at];
    if (token.kind === 'open') {
      const previous = tokens[at - 1];
      stack.push(
        previous?.kind === 'word'
          ? previous.value.toUpperCase()
          : previous?.kind === 'operator' && previous.value === '@'
            ? '@'
            : null
      );
      continue;
    }
    if (token.kind === 'close') {
      stack.pop();
      continue;
    }
    // A reference starts at a sheet qualifier when there is one.
    const first = at;
    let word = at;
    if (token.kind === 'sheet') {
      if (tokens[at + 1]?.kind !== 'word') {
        continue;
      }
      word = at + 1;
    } else if (token.kind !== 'word') {
      continue;
    }
    const value = (tokens[word] as { value: string }).value;
    const next = tokens[word + 1];
    if (next?.kind === 'open') {
      const name = value.toUpperCase();
      if (name === 'OFFSET' || name === 'INDIRECT') {
        // A computed reference is a range; it runs to its own bracket.
        let depth = 0;
        let close = word + 1;
        for (; close < tokens.length; close++) {
          if (tokens[close].kind === 'open') {
            depth++;
          } else if (tokens[close].kind === 'close' && --depth === 0) {
            break;
          }
        }
        mark(first, close);
      }
      at = word;
      continue;
    }
    if (next?.kind === 'colon' && tokens[word + 2]?.kind === 'word') {
      mark(first, word + 2);
      at = word + 2;
      continue;
    }
    if (names.has(value.toUpperCase())) {
      mark(first, word);
    }
    at = word;
  }

  if (inserts.length === 0) {
    return input;
  }
  let text = body;
  for (const offset of inserts.sort((a, b) => b - a)) {
    text = `${text.slice(0, offset)}@${text.slice(offset)}`;
  }
  return `=${text}`;
}
