import { describe, expect, it } from 'vitest';

import { rewriteFormula } from './Rewrite';

describe('moving a formula', () => {
  it('moves a relative reference by the same amount', () => {
    expect(rewriteFormula('=A1', 1, 0)).toBe('=A2');
    expect(rewriteFormula('=A1', 0, 1)).toBe('=B1');
    expect(rewriteFormula('=A1', 2, 3)).toBe('=D3');
  });

  /**
   * The whole reason `parseRef` keeps the `$` flags rather than
   * resolving them. Filling a share column down has to move the row's
   * revenue and leave the total alone.
   */
  it('leaves an absolute half where it was', () => {
    expect(rewriteFormula('=$A$1', 5, 5)).toBe('=$A$1');
    expect(rewriteFormula('=$A1', 5, 5)).toBe('=$A6');
    expect(rewriteFormula('=A$1', 5, 5)).toBe('=F$1');
  });

  it('moves the references inside a call and a range', () => {
    expect(rewriteFormula('=SUM(A1:A5)', 1, 0)).toBe('=SUM(A2:A6)');
    expect(rewriteFormula('=SUM(A$1:A5)', 1, 0)).toBe('=SUM(A$1:A6)');
  });

  it('moves them through operators and leaves the literals', () => {
    expect(rewriteFormula('=B2*C2', 1, 0)).toBe('=(B3*C3)');
    expect(rewriteFormula('=A1+1', 1, 0)).toBe('=(A2+1)');
    expect(rewriteFormula('=IF(A1>0,"yes","no")', 1, 0)).toBe('=IF((A2>0),"yes","no")');
  });

  /** The seed's share column, which is why it was written that way. */
  it('moves the row and keeps the total', () => {
    expect(rewriteFormula('=ROUND(D2/$D$7*100,1)', 1, 0)).toBe('=ROUND(((D3/$D$7)*100),1)');
  });

  it('is unchanged by a move of nothing', () => {
    expect(rewriteFormula('=B2*C2', 0, 0)).toBe('=B2*C2');
  });

  it('leaves a literal alone', () => {
    expect(rewriteFormula('120', 3, 3)).toBe('120');
    expect(rewriteFormula('North', 3, 3)).toBe('North');
  });

  /**
   * A reference pushed off the sheet becomes `#REF!` in the text,
   * which is what a spreadsheet writes: the formula stays readable and
   * says which part of it no longer points anywhere.
   */
  it('writes #REF! for a reference pushed off the sheet', () => {
    expect(rewriteFormula('=A1', -1, 0)).toBe('=#REF!');
    expect(rewriteFormula('=A1+B1', 0, -1)).toBe('=(#REF!+A1)');
    expect(rewriteFormula('=SUM(A1:B1)', -5, 0)).toBe('=SUM(#REF!)');
  });

  it('moves text it cannot parse without mangling it', () => {
    expect(rewriteFormula('=1+', 1, 0)).toBe('=1+');
  });

  it('keeps a quoted string intact', () => {
    expect(rewriteFormula('=A1&" says ""hi"""', 1, 0)).toBe('=(A2&" says ""hi""")');
  });

  /**
   * A fill moves where a reference points, never which sheet it
   * points at. There is no third delta and there is not meant to be
   * one: `Sheet2!A1` dragged down is `Sheet2!A2`, on Sheet 2.
   */
  it('carries the sheet a reference names', () => {
    expect(rewriteFormula('=Sheet2!A1', 1, 0)).toBe('=Sheet2!A2');
    expect(rewriteFormula('=Sheet2!A1:B2', 1, 0)).toBe('=Sheet2!A2:B3');
    expect(rewriteFormula('=Sheet2!$A$1+A1', 1, 0)).toBe('=(Sheet2!$A$1+A2)');
  });

  it('quotes a sheet name that needs it, and only then', () => {
    expect(rewriteFormula("='Q3 Budget'!A1", 1, 0)).toBe("='Q3 Budget'!A2");
    expect(rewriteFormula("='Kevin''s'!A1", 1, 0)).toBe("='Kevin''s'!A2");
  });

  /**
   * `A:A` means the column and has to go on meaning it. Spelled out
   * as `A$1:A$1048576` it is a rectangle somebody drew, which stops
   * covering the sheet the moment a row is added at the bottom —
   * and spelling it out is what a fill used to do to it.
   */
  it('keeps a whole column whole', () => {
    expect(rewriteFormula('=SUM(A:A)', 0, 1)).toBe('=SUM(B:B)');
    expect(rewriteFormula('=SUM(A:C)', 0, 2)).toBe('=SUM(C:E)');
    expect(rewriteFormula('=SUM(Sheet2!A:A)', 0, 1)).toBe('=SUM(Sheet2!B:B)');
  });
});
