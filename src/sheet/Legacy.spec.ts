import { describe, expect, it } from 'vitest';

import { withIntersections } from './Legacy';

/** Formulas from before dynamic arrays, given the `@` Excel 365 shows them with. */
describe('a formula from an older file', () => {
  const names = new Set(['PRICES', 'QTY']);
  const cases: readonly (readonly [string, string])[] = [
    // A range where one value is wanted.
    ['=B1:B9*2', '=@B1:B9*2'],
    ['=A1:A9', '=@A1:A9'],
    ['=-A1:A9', '=-@A1:A9'],
    ['=Sheet2!B1:B9+1', '=@Sheet2!B1:B9+1'],
    ["='Q3 Budget'!B1:B9+1", "=@'Q3 Budget'!B1:B9+1"],
    ['=Prices*Qty', '=@Prices*@Qty'],
    ['=IF(A1:A9>1, "y", "n")', '=IF(@A1:A9>1, "y", "n")'],
    ['=IF(A1:A9, 1, 2)', '=IF(@A1:A9, 1, 2)'],
    ['=OFFSET(A1, 0, 0, 3, 1)', '=@OFFSET(A1, 0, 0, 3, 1)'],
    ['=SUM(A1:A9*2)', '=SUM(@A1:A9*2)'],
    // A function of one value runs across a range now; one written when
    // it took one value from it says so.
    ['=ABS(A1:A9)', '=ABS(@A1:A9)'],
    ['=ROUND(A1:A9, 2)', '=ROUND(@A1:A9, 2)'],
    ['=IFERROR(A1:A9/B1:B9, 0)', '=IFERROR(@A1:A9/@B1:B9, 0)'],
    // A range that is an argument, as a range.
    ['=SUM(A1:A9)', '=SUM(A1:A9)'],
    ['=SUM(A:A)', '=SUM(A:A)'],
    ['=VLOOKUP(A1, D1:F9, 2, FALSE)', '=VLOOKUP(A1, D1:F9, 2, FALSE)'],
    ['=SUM(Prices)', '=SUM(Prices)'],
    ['=SUM(OFFSET(A1, 0, 0, 3, 1))', '=SUM(OFFSET(A1, 0, 0, 3, 1))'],
    // Arrays all the way down, in every version.
    ['=SUMPRODUCT(A1:A9*B1:B9)', '=SUMPRODUCT(A1:A9*B1:B9)'],
    ['=SUMPRODUCT(--(B5:B20))', '=SUMPRODUCT(--(B5:B20))'],
    // Nothing that is one value already.
    ['=A1*2', '=A1*2'],
    ['=SUM(A1, B1)', '=SUM(A1, B1)'],
    ['="A1:A9"&B1', '="A1:A9"&B1'],
    ['=@A1:A9*2', '=@A1:A9*2'],
    ['=@(A1:A9)*2', '=@(A1:A9)*2'],
    ['plain text', 'plain text']
  ];
  for (const [written, meant] of cases) {
    it(`reads ${written} as ${meant}`, () => {
      expect(withIntersections(written, names)).toBe(meant);
    });
  }
});
