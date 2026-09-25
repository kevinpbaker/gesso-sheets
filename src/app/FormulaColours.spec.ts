import { describe, expect, it } from 'vitest';

import { colouredReferences, formulaSpans } from './FormulaColours';

/**
 * The colours, and the one invariant that keeps the caret honest.
 *
 * `formulaSpans` feeds an `EditableText`, whose runs supply the
 * paragraph's text. If they do not concatenate to exactly the value,
 * the engine throws them away and draws plainly — so the failure is
 * survivable, and it is still a failure. Every case below checks the
 * concatenation as well as whatever else it is about.
 */

const texts = (text: string) => formulaSpans(text)?.map(run => run.text);
const colours = (text: string) => formulaSpans(text)?.map(run => run.color);

/** The invariant, asserted everywhere it could break. */
function tiles(text: string): void {
  const runs = formulaSpans(text);
  if (runs !== undefined) {
    expect(runs.map(run => run.text).join('')).toBe(text);
  }
}

describe('colouring the references in a formula', () => {
  it('splits the text into the references and the rest', () => {
    expect(texts('=A1+B2')).toEqual(['=', 'A1', '+', 'B2']);
    tiles('=A1+B2');
  });

  it('gives each distinct reference its own colour', () => {
    const [, first, , second] = colours('=A1+B2')!;
    expect(first).not.toBe(second);
    expect(first).toMatch(/^#[0-9a-f]{6}$/);
  });

  it('leaves the gaps between them uncoloured', () => {
    expect(colours('=A1+B2')).toEqual([undefined, expect.any(String), undefined, expect.any(String)]);
  });

  /**
   * The dollars change how a reference copies, not what it points at.
   * Colouring them differently would suggest otherwise.
   */
  it('gives the same cells the same colour however they were written', () => {
    const runs = formulaSpans('=A1+$A$1')!;
    const coloured = runs.filter(run => run.color !== undefined);
    expect(coloured.map(run => run.text)).toEqual(['A1', '$A$1']);
    expect(coloured[0].color).toBe(coloured[1].color);
  });

  it('colours a range as one reference', () => {
    expect(texts('=SUM(A1:B9)')).toEqual(['=SUM(', 'A1:B9', ')']);
    tiles('=SUM(A1:B9)');
  });

  it('repeats the palette rather than running out', () => {
    const runs = formulaSpans('=A1+B1+C1+D1+E1+F1')!;
    const used = runs.filter(run => run.color !== undefined).map(run => run.color);
    expect(used).toHaveLength(6);
    // Five hues, so the sixth reference wears the first one again.
    expect(used[5]).toBe(used[0]);
    expect(new Set(used).size).toBe(5);
  });

  it('is nothing at all when there is nothing to colour', () => {
    expect(formulaSpans('=1+2')).toBeUndefined();
    expect(formulaSpans('North')).toBeUndefined();
    expect(formulaSpans('')).toBeUndefined();
  });

  /** The state a formula spends most of its life in while being typed. */
  it('tiles an unfinished formula', () => {
    for (const text of ['=A1+', '=SUM(A1:', '=SUM(A1,B2', '=A1)', '=  A1  +  B2  ']) {
      tiles(text);
    }
    expect(texts('=SUM(A1:')).toEqual(['=SUM(', 'A1', ':']);
  });

  it('tiles a formula that ends on a reference', () => {
    expect(texts('=1+A1')).toEqual(['=1+', 'A1']);
    tiles('=1+A1');
  });

  it('tiles a formula that is only a reference', () => {
    expect(texts('=A1')).toEqual(['=', 'A1']);
    tiles('=A1');
  });
});

describe('the references a formula points at', () => {
  it('reports the cells, for drawing a box round them', () => {
    const [first] = colouredReferences('=SUM(B2:D7)');
    expect(first.range.start).toMatchObject({ row: 1, column: 1 });
    expect(first.range.end).toMatchObject({ row: 6, column: 3 });
  });

  /** The text and the box have to agree, which is the whole point. */
  it('gives the box the colour the text is drawn in', () => {
    const references = colouredReferences('=A1+B2');
    const runs = formulaSpans('=A1+B2')!.filter(run => run.color !== undefined);
    expect(references.map(r => r.color)).toEqual(runs.map(run => run.color));
  });

  it('reports where in the text each one was written', () => {
    expect(colouredReferences('=A1+B2').map(r => r.span)).toEqual([
      { start: 1, end: 3 },
      { start: 4, end: 6 }
    ]);
  });
});

/**
 * The bracket beside the caret and the one that closes it.
 *
 * Washed rather than coloured: the five hues already mean "this is
 * the reference that box belongs to", and a sixth meaning "these two
 * are a pair" would be one too many to read at a glance.
 */
describe('marking a pair of brackets', () => {
  /** The offsets of the washed runs, which says *which* brackets. */
  function markedAt(text: string, caret?: number): number[] {
    const runs = formulaSpans(text, caret) ?? [];
    const found: number[] = [];
    let at = 0;
    for (const run of runs) {
      if (run.backgroundColor !== undefined) {
        found.push(at);
      }
      at += run.text.length;
    }
    return found;
  }

  //  =SUM(ROUND(A1,2))
  //  0    4     10   16
  const NESTED = '=SUM(ROUND(A1,2))';

  it('marks both when the caret is after the closing one', () => {
    expect(markedAt('=SUM(A1)', 8)).toEqual([4, 7]);
  });

  it('marks both when the caret is before the opening one', () => {
    expect(markedAt('=SUM(A1)', 4)).toEqual([4, 7]);
  });

  it('marks the outer pair from outside it', () => {
    expect(markedAt(NESTED, 17)).toEqual([4, 16]);
  });

  it('marks the inner pair from inside it', () => {
    expect(markedAt(NESTED, 16)).toEqual([10, 15]);
  });

  it('marks nothing when the caret is not on a bracket', () => {
    expect(markedAt('=SUM(A1)', 6)).toEqual([]);
  });

  it('marks nothing when the bracket was never closed', () => {
    expect(markedAt('=SUM(A1', 4)).toEqual([]);
  });

  it('marks nothing when no caret was offered', () => {
    expect(markedAt('=SUM(A1)')).toEqual([]);
  });

  it('still tiles the whole string when it marks', () => {
    for (const caret of [4, 10, 15, 16, 17]) {
      const runs = formulaSpans(NESTED, caret);
      expect((runs ?? []).map(run => run.text).join('')).toBe(NESTED);
    }
  });

  it('keeps the references coloured alongside', () => {
    const runs = formulaSpans('=SUM(A1)', 8)!;
    expect(runs.find(run => run.text === 'A1')?.color).toBeDefined();
    expect(runs.map(run => run.text).join('')).toBe('=SUM(A1)');
  });
});
