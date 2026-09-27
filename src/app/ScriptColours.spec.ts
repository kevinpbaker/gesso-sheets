import { describe, expect, it } from 'vitest';

import { formulaSpans, scriptSpans } from './FormulaColours';

describe('the script editor’s colouring', () => {
  it('colours the references in a script’s strings as a formula’s are, the same range the same colour', () => {
    const text = 'const v = sheet.range("A2:A9").values;\nsheet.range(\'B2\').write(v);\nsheet.range("A2:A9").format({ bold: true });';
    const coloured = (scriptSpans(text) ?? []).filter(run => run.color !== undefined);
    expect(coloured.map(run => run.text)).toEqual(['A2:A9', 'B2', 'A2:A9']);
    expect(coloured[0].color).toBe(coloured[2].color);
    expect(coloured[1].color).not.toBe(coloured[0].color);
    // The same colours the formula bar would use, in the same order.
    const formula = (formulaSpans('=A2:A9+B2') ?? []).filter(run => run.color !== undefined);
    expect(coloured.slice(0, 2).map(run => run.color)).toEqual(formula.map(run => run.color));
    // And the runs spell the text, or the field would draw it unstyled.
    expect((scriptSpans(text) ?? []).map(run => run.text).join('')).toBe(text);
  });

  it('leaves names outside strings alone, and a script with no references costs nothing', () => {
    expect(scriptSpans('const A1 = 2; console.log(A1)')).toBeUndefined();
    expect(scriptSpans('console.log("Total")')).toBeUndefined();
  });
});
