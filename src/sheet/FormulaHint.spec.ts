import { describe, expect, it } from 'vitest';

import { acceptCompletion, hintFor, markedArgument } from './FormulaHint';
import { signatureOf } from './Signatures';

/** Splits `=SUM(|` into the text and the caret offset. */
function at(withCaret: string): [string, number] {
  const caret = withCaret.indexOf('|');
  return [withCaret.slice(0, caret) + withCaret.slice(caret + 1), caret];
}

const hint = (withCaret: string) => hintFor(...at(withCaret));

describe('offering a list of function names', () => {
  it('offers them while a bare word is being typed', () => {
    const found = hint('=SU|');
    expect(found).toMatchObject({ kind: 'completions', prefix: 'SU' });
    expect(found?.kind === 'completions' ? found.names : []).toContain('SUM');
  });

  it('says where the word is, so accepting one can replace it', () => {
    expect(hint('=1+SU|')).toMatchObject({ span: { start: 3, end: 5 } });
  });

  /** Somebody editing the middle of a word is not asking for a list. */
  it('offers nothing from inside a word', () => {
    expect(hint('=S|UM')).toBeNull();
  });

  it('offers nothing once the bracket is typed', () => {
    expect(hint('=SUM|(')).toBeNull();
  });

  /**
   * A word that is an address is an address. Without this, typing `B`
   * on the way to `B2` offers every function beginning with B.
   */
  it('offers nothing for a cell reference', () => {
    expect(hint('=A1|')).toBeNull();
  });

  it('offers nothing for a word it does not know', () => {
    expect(hint('=ZZZ|')).toBeNull();
  });

  it('offers nothing when the text is not a formula', () => {
    expect(hint('SU|')).toBeNull();
  });
});

describe('showing a signature', () => {
  it('shows it once the arguments are being filled in', () => {
    expect(hint('=ROUND(|')).toMatchObject({ kind: 'signature', name: 'ROUND', argument: 0 });
  });

  it('follows the caret from argument to argument', () => {
    expect(hint('=ROUND(A1, |')).toMatchObject({ argument: 1 });
    expect(hint('=IF(A1, 2, |')).toMatchObject({ argument: 2 });
  });

  it('shows the innermost call when they are nested', () => {
    expect(hint('=SUM(ROUND(A1, |')).toMatchObject({ name: 'ROUND', argument: 1 });
  });

  it('goes back to the outer one past the inner bracket', () => {
    expect(hint('=SUM(ROUND(A1, 2), |')).toMatchObject({ name: 'SUM', argument: 1 });
  });

  /** A word being typed is the more specific state, so it wins. */
  it('gives way to a list of names inside a call', () => {
    expect(hint('=SUM(RO|')).toMatchObject({ kind: 'completions', prefix: 'RO' });
  });

  it('shows nothing outside any call', () => {
    expect(hint('=A1+|')).toBeNull();
  });

  it('shows nothing for a name the sheet does not know', () => {
    expect(hint('=NOSUCH(|')).toBeNull();
  });
});

describe('accepting a completion', () => {
  it('writes the name and opens the bracket', () => {
    expect(acceptCompletion('=SU', { start: 1, end: 3 }, 'SUM')).toEqual({ text: '=SUM(', caret: 5 });
  });

  /** The span comes from the hint, so the two agree about the word. */
  it('replaces exactly what the hint pointed at', () => {
    const [text, caret] = at('=1+SU|');
    const found = hintFor(text, caret);
    if (found?.kind !== 'completions') {
      throw new Error('expected a list of names');
    }
    expect(acceptCompletion(text, found.span, 'SUMIF')).toEqual({ text: '=1+SUMIF(', caret: 9 });
  });

  it('leaves what was after the word alone', () => {
    expect(acceptCompletion('=SU+1', { start: 1, end: 3 }, 'SUM')).toEqual({ text: '=SUM(+1', caret: 5 });
  });
});

describe('which argument a signature marks', () => {
  it('marks the one the caret is in', () => {
    expect(markedArgument(signatureOf('ROUND')!, 1)).toBe(1);
  });

  /** `SUM(number, …)`: the fourth argument is still a number. */
  it('keeps marking the last one when it repeats', () => {
    expect(markedArgument(signatureOf('SUM')!, 7)).toBe(1);
  });

  it('marks nothing past the end when it does not repeat', () => {
    expect(markedArgument(signatureOf('ROUND')!, 5)).toBe(-1);
  });

  it('marks nothing for a function that takes none', () => {
    expect(markedArgument(signatureOf('TODAY')!, 0)).toBe(-1);
  });
});
