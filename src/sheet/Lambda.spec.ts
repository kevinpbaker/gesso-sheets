import { describe, expect, it } from 'vitest';

import { MAX_DEPTH } from './Evaluator';
import { rewriteFormula } from './Rewrite';
import { shiftFormula } from './Shift';
import type { Sheet } from './Sheet';
import { Workbook } from './Workbook';
import { formatValue } from './Values';

/**
 * Phase 28: `LET`, `LAMBDA`, and names that hold formulas.
 *
 * Through the whole engine, as the conformance table is: parse, graph,
 * recalculate, read. The table asserts what each function answers; this
 * asserts the things a table of one-liners cannot — what a name can
 * see, who depends on whom, and how deep a function may call itself.
 */

function sheet(): Sheet {
  return new Workbook().sheet(0);
}

/** A formula's answer, as the cell would show it. */
function answer(at: Sheet, formula: string): string {
  at.setCell(50, 50, formula);
  at.recalculate();
  return formatValue(at.value(50, 50));
}

describe('LET', () => {
  it('lets each name see the ones before it, and the result see all of them', () => {
    expect(answer(sheet(), '=LET(a, 2, b, a * 10, c, a + b, c * 100)')).toBe('2200');
  });

  it('shadows a workbook name with its own', () => {
    const at = sheet();
    at.setCell(0, 0, '5');
    at.names.defineFormula('Rate', '=0.5');
    at.namesChanged();
    expect(answer(at, '=Rate * 2')).toBe('1');
    expect(answer(at, '=LET(Rate, 3, Rate * 2)')).toBe('6');
  });

  it('keeps a name inside the LET that bound it', () => {
    expect(answer(sheet(), '=LET(x, 1, x) + x')).toBe('#NAME?');
  });

  it('refuses the same name twice, and a name that is a cell', () => {
    expect(answer(sheet(), '=LET(x, 1, x, 2, x)')).toBe('#VALUE!');
    expect(answer(sheet(), '=LET(A1, 1, A1)')).toBe('#VALUE!');
  });

  it('binds a range as a range, so a function over it sees all of it', () => {
    const at = sheet();
    for (let row = 0; row < 4; row++) {
      at.setCell(row, 0, String(row + 1));
    }
    expect(answer(at, '=LET(r, A1:A4, SUM(r) / COUNT(r))')).toBe('2.5');
  });

  it('spills an array it hands back', () => {
    const at = sheet();
    at.setCell(0, 5, '=LET(n, 3, SEQUENCE(n) * 10)');
    at.recalculate();
    expect([at.value(0, 5), at.value(1, 5), at.value(2, 5)]).toEqual([10, 20, 30]);
  });
});

describe('LAMBDA', () => {
  it('is called where it is written', () => {
    const at = sheet();
    at.setCell(0, 0, '7');
    expect(answer(at, '=LAMBDA(x, x * 2)(A1)')).toBe('14');
  });

  it('closes over the LET it was made in', () => {
    expect(answer(sheet(), '=LET(rate, 0.25, tax, LAMBDA(x, x * rate), tax(80) + tax(20))')).toBe('25');
  });

  it('is a function handed to another', () => {
    expect(answer(sheet(), '=LET(twice, LAMBDA(f, x, f(f(x))), inc, LAMBDA(n, n + 1), twice(inc, 5))')).toBe('7');
  });

  it('is #CALC! in a cell, uncalled', () => {
    expect(answer(sheet(), '=LAMBDA(x, x + 1)')).toBe('#CALC!');
  });

  it('takes exactly its parameters', () => {
    expect(answer(sheet(), '=LAMBDA(x, y, x + y)(1)')).toBe('#VALUE!');
    expect(answer(sheet(), '=LAMBDA(x, x, 1)(1, 2)')).toBe('#VALUE!');
  });

  it('can be handed an error, and deal with it', () => {
    expect(answer(sheet(), '=LAMBDA(x, IFERROR(x, "none"))(1/0)')).toBe('none');
  });
});

describe('an optional parameter', () => {
  it('may be left off, and ISOMITTED says so', () => {
    const at = sheet();
    at.names.defineFormula('Taxed', '=LAMBDA(amount, [rate], amount * (1 + IF(ISOMITTED(rate), 0.2, rate)))');
    at.namesChanged();
    expect(answer(at, '=Taxed(100)')).toBe('120');
    expect(answer(at, '=Taxed(100, 0.5)')).toBe('150');
  });

  it('is FALSE where it is read, left off', () => {
    expect(answer(sheet(), '=LAMBDA(a, [b], b)(1)')).toBe('FALSE');
    expect(answer(sheet(), '=LAMBDA(a, [b], a + b)(1)')).toBe('1');
  });

  it('comes after the ones that must be given', () => {
    expect(answer(sheet(), '=LAMBDA([a], b, b)(1)')).toBe('#VALUE!');
    expect(answer(sheet(), '=LAMBDA(a, [b], a)()')).toBe('#VALUE!');
    expect(answer(sheet(), '=LAMBDA(a, [b], a)(1, 2, 3)')).toBe('#VALUE!');
  });

  it('is only a parameter, and prints back in its brackets', () => {
    expect(answer(sheet(), '=[b]+1')).toBe('#VALUE!');
    expect(rewriteFormula('=LAMBDA(a, [b], a+A1)(B1)', 1, 0)).toBe('=LAMBDA(a,[b],a+A2)(B2)');
  });
});

describe('a name that holds a formula', () => {
  it('is called by its name, like any other function', () => {
    const at = sheet();
    at.setCell(0, 0, '21');
    at.names.defineFormula('Double', '=LAMBDA(x, x * 2)');
    at.namesChanged();
    expect(answer(at, '=Double(A1)')).toBe('42');
  });

  it('is a calculation with a name, when it is not a function', () => {
    const at = sheet();
    at.setCell(0, 0, '100');
    at.names.defineFormula('Gross', '=A1 * 1.2');
    at.namesChanged();
    expect(answer(at, '=Gross + 1')).toBe('121');
  });

  it('is #VALUE! called when it is not a function, and #CALC! uncalled when it is', () => {
    const at = sheet();
    at.names.defineFormula('Rate', '=0.2');
    at.names.defineFormula('Double', '=LAMBDA(x, x * 2)');
    at.namesChanged();
    expect(answer(at, '=Rate(1)')).toBe('#VALUE!');
    expect(answer(at, '=Double')).toBe('#CALC!');
  });

  it('refuses a formula that does not parse, and keeps the name it had', () => {
    const at = sheet();
    at.names.defineFormula('Rate', '=0.2');
    expect(at.names.defineFormula('Rate', '=0.2 +')).toBe('formula');
    expect(at.names.defineFormula('A1', '=1')).toBe('reference');
    at.namesChanged();
    expect(answer(at, '=Rate')).toBe('0.2');
  });

  it('reads an unqualified cell on the sheet that calls it', () => {
    const book = new Workbook();
    book.addSheet('Two');
    const one = book.sheet(0);
    const two = book.sheet(1);
    one.setCell(0, 0, '1');
    two.setCell(0, 0, '2');
    one.names.defineFormula('Here', '=A1 * 10');
    one.namesChanged();
    one.setCell(5, 0, '=Here');
    two.setCell(5, 0, '=Here');
    book.recalculate();
    expect([one.value(5, 0), two.value(5, 0)]).toEqual([10, 20]);
  });
});

describe('recursion', () => {
  function withFactorial(): Sheet {
    const at = sheet();
    at.names.defineFormula('Fact', '=LAMBDA(n, IF(n <= 1, 1, n * Fact(n - 1)))');
    at.namesChanged();
    return at;
  }

  it('lets a named LAMBDA call itself', () => {
    expect(answer(withFactorial(), '=Fact(10)')).toBe('3628800');
  });

  it('goes as deep as the limit allows', () => {
    const at = sheet();
    expect(at.names.defineFormula('Depth', '=LAMBDA(n, IF(n = 0, 0, 1 + Depth(n - 1)))')).toBeNull();
    at.namesChanged();
    expect(answer(at, `=Depth(${MAX_DEPTH - 10})`)).toBe(String(MAX_DEPTH - 10));
  });

  it('is #NUM! past the limit, rather than a tab that stops answering', () => {
    const at = sheet();
    at.names.defineFormula('Forever', '=LAMBDA(n, Forever(n + 1))');
    at.namesChanged();
    expect(answer(at, '=Forever(1)')).toBe('#NUM!');
  });

  it('is #NUM! for a name defined as itself', () => {
    const at = sheet();
    at.names.defineFormula('Loop', '=Loop + 1');
    at.namesChanged();
    expect(answer(at, '=Loop')).toBe('#NUM!');
  });
});

describe('the helpers, each with a LAMBDA', () => {
  function numbers(): Sheet {
    const at = sheet();
    // A1:B3 is 1 2 / 3 4 / 5 6.
    for (let row = 0; row < 3; row++) {
      at.setCell(row, 0, String(row * 2 + 1));
      at.setCell(row, 1, String(row * 2 + 2));
    }
    return at;
  }

  /** Where a formula at F1 spilled, as rows of values. */
  function spilled(at: Sheet, formula: string, rows: number, columns: number): unknown[][] {
    at.setCell(0, 5, formula);
    at.recalculate();
    return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => at.value(row, 5 + column)));
  }

  it('MAP calls it on every value, pairing two arrays', () => {
    expect(spilled(numbers(), '=MAP(A1:A3, B1:B3, LAMBDA(a, b, a * b))', 3, 1)).toEqual([[2], [12], [30]]);
  });

  it('REDUCE folds, and SCAN keeps the steps', () => {
    expect(answer(numbers(), '=REDUCE(0, A1:B3, LAMBDA(total, v, total + v))')).toBe('21');
    expect(spilled(numbers(), '=SCAN(0, A1:A3, LAMBDA(total, v, total + v))', 3, 1)).toEqual([[1], [4], [9]]);
  });

  it('REDUCE can build an array as it goes', () => {
    expect(spilled(numbers(), '=REDUCE(0, A1:A3, LAMBDA(total, v, total + SEQUENCE(1, 2)))', 1, 2)).toEqual([[3, 6]]);
  });

  it('BYROW and BYCOL hand it a row or a column at a time', () => {
    expect(spilled(numbers(), '=BYROW(A1:B3, LAMBDA(row, SUM(row)))', 3, 1)).toEqual([[3], [7], [11]]);
    expect(spilled(numbers(), '=BYCOL(A1:B3, LAMBDA(column, MAX(column)))', 1, 2)).toEqual([[5, 6]]);
  });

  it('MAKEARRAY builds from positions, counting from one', () => {
    expect(spilled(sheet(), '=MAKEARRAY(2, 3, LAMBDA(r, c, r & "," & c))', 2, 3)).toEqual([
      ['1,1', '1,2', '1,3'],
      ['2,1', '2,2', '2,3']
    ]);
  });

  it('takes a named LAMBDA as well as a written one', () => {
    const at = numbers();
    at.names.defineFormula('Square', '=LAMBDA(x, x ^ 2)');
    at.namesChanged();
    expect(spilled(at, '=MAP(A1:A3, Square)', 3, 1)).toEqual([[1], [9], [25]]);
  });

  it('refuses a function with the wrong number of parameters', () => {
    expect(answer(numbers(), '=MAP(A1:A3, LAMBDA(a, b, a))')).toBe('#VALUE!');
    expect(answer(numbers(), '=REDUCE(0, A1:A3, LAMBDA(v, v))')).toBe('#VALUE!');
    expect(answer(numbers(), '=MAP(A1:A3, 5)')).toBe('#VALUE!');
  });
});

/**
 * The graph: a caller depends on what the function reads, not only on
 * what it was handed. Without these edges every one of these answers is
 * right once and stale after.
 */
describe('who depends on a LAMBDA', () => {
  function taxed(): Sheet {
    const at = sheet();
    at.setCell(0, 0, '0.2'); // A1: the rate
    at.setCell(1, 0, '100'); // A2: a price
    at.names.defineFormula('TaxRate', '=Sheet1!$A$1');
    at.names.defineFormula('WithTax', '=LAMBDA(x, x * (1 + TaxRate))');
    at.namesChanged();
    at.setCell(1, 1, '=WithTax(A2)');
    at.recalculate();
    return at;
  }

  it('wakes a caller when a cell its function reads changes', () => {
    const at = taxed();
    expect(at.value(1, 1)).toBe(120);
    at.setCell(0, 0, '0.5');
    at.recalculate();
    expect(at.value(1, 1)).toBe(150);
  });

  it('wakes a caller when the argument changes', () => {
    const at = taxed();
    at.setCell(1, 0, '200');
    at.recalculate();
    expect(at.value(1, 1)).toBeCloseTo(240);
  });

  it('wakes every caller when the name is defined again', () => {
    const at = taxed();
    at.names.defineFormula('WithTax', '=LAMBDA(x, x + 1)');
    at.namesChanged();
    at.recalculate();
    expect(at.value(1, 1)).toBe(101);
  });

  it('makes a caller of a function that reads the clock recalculate on every edit', () => {
    const book = new Workbook();
    let tick = 0;
    book.clock = () => ++tick;
    const at = book.sheet(0);
    at.names.defineFormula('Stamp', '=LAMBDA(x, NOW() + x)');
    at.namesChanged();
    at.setCell(0, 0, '=Stamp(0)');
    at.recalculate();
    const first = at.value(0, 0);
    at.setCell(5, 5, 'anything');
    at.recalculate();
    expect(at.value(0, 0)).not.toBe(first);
  });
});

/** What moves a formula moves the names a LET bound, and keeps them as typed. */
describe('a formula with names in it, moved', () => {
  it('keeps a named range a name when it is filled', () => {
    expect(rewriteFormula('=SUM(Sales)+A1', 1, 0)).toBe('=SUM(Sales)+A2');
  });

  it('keeps LET and LAMBDA names as they were typed', () => {
    expect(rewriteFormula('=LET(price, A1, price*2)', 1, 0)).toBe('=LET(price,A2,price*2)');
    expect(rewriteFormula('=LAMBDA(x, x+A1)(B1)', 0, 1)).toBe('=LAMBDA(x,x+B1)(C1)');
  });

  it('moves the references inside a name that holds a formula', () => {
    expect(shiftFormula('=LAMBDA(x, x*Sheet1!$A$5)', { axis: 'row', at: 2, by: 1, sheet: 'Sheet1' })).toBe(
      '=LAMBDA(x,x*Sheet1!$A$6)'
    );
  });
});
