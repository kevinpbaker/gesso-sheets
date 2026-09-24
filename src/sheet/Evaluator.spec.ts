import { describe, expect, it } from 'vitest';

import { cellKey } from './A1';
import { evaluate } from './Evaluator';
import { parseFormula } from './Parser';
import { CIRC, DIV0, formatValue, NAME, REF, VALUE, type CellValue } from './Values';

/** A sheet as a literal, so a spec reads as the grid it describes. */
function context(grid: Record<string, CellValue>) {
  const cells = new Map<number, CellValue>();
  for (const [address, value] of Object.entries(grid)) {
    const column = address.charCodeAt(0) - 65;
    const row = Number(address.slice(1)) - 1;
    cells.set(cellKey(row, column), value);
  }
  return { valueAt: (key: number) => cells.get(key) ?? null };
}

function run(source: string, grid: Record<string, CellValue> = {}): CellValue {
  return evaluate(parseFormula(source), context(grid));
}

describe('arithmetic', () => {
  it('does the obvious things', () => {
    expect(run('1+2')).toBe(3);
    expect(run('7-2')).toBe(5);
    expect(run('3*4')).toBe(12);
    expect(run('9/2')).toBe(4.5);
    expect(run('2^10')).toBe(1024);
  });

  it('reports a division by zero rather than returning Infinity', () => {
    // Infinity in a cell is a value that keeps arithmetic going and
    // says nothing about where it came from.
    expect(run('1/0')).toBe(DIV0);
    expect(run('0/0')).toBe(DIV0);
  });

  it('reports a power with no real answer rather than returning NaN', () => {
    // A NaN compares false with itself and poisons everything
    // downstream silently.
    expect(run('(0-8)^0.5')).toBe(VALUE);
  });

  it('treats an empty cell as zero', () => {
    expect(run('A1+1')).toBe(1);
  });

  it('reads text that is a number as a number', () => {
    expect(run('A1+1', { A1: '5' })).toBe(6);
  });

  it('refuses text that is not a number', () => {
    expect(run('A1+1', { A1: 'five' })).toBe(VALUE);
  });
});

describe('text', () => {
  it('joins with &', () => {
    expect(run('"a"&"b"')).toBe('ab');
    expect(run('A1&B1', { A1: 'Hello, ', B1: 'world' })).toBe('Hello, world');
  });

  it('writes a number the way a person reads it', () => {
    expect(run('"n="&A1', { A1: 3 })).toBe('n=3');
    // Not 0.30000000000000004, which is what the arithmetic actually
    // produced and not what anybody typed.
    expect(run('""&(0.1+0.2)')).toBe('0.3');
  });

  it('writes a boolean in capitals, as a sheet does', () => {
    expect(run('""&TRUE')).toBe('TRUE');
  });
});

describe('comparison', () => {
  it('compares numbers', () => {
    expect(run('1<2')).toBe(true);
    expect(run('2<=2')).toBe(true);
    expect(run('3<>3')).toBe(false);
    expect(run('A1=1', { A1: 1 })).toBe(true);
  });

  it('compares text without regard to case, as a sheet does', () => {
    expect(run('"abc"="ABC"')).toBe(true);
    expect(run('"a"<"b"')).toBe(true);
  });

  it('orders numbers before text and text before booleans', () => {
    expect(run('A1<B1', { A1: 999, B1: 'a' })).toBe(true);
    expect(run('A1<B1', { A1: 'z', B1: true })).toBe(true);
  });

  it('treats an empty cell as zero, not as empty text', () => {
    expect(run('A1=0')).toBe(true);
  });
});

describe('errors', () => {
  it('travel, so the cell that broke is the one named', () => {
    expect(run('A1+1', { A1: DIV0 })).toBe(DIV0);
    expect(run('SUM(A1:B1)', { A1: 1, B1: REF })).toBe(REF);
    expect(run('"x"&A1', { A1: CIRC })).toBe(CIRC);
  });

  it('are values a formula can name', () => {
    expect(run('#REF!')).toEqual(REF);
  });

  it('report an unknown name', () => {
    expect(run('NOSUCHTHING()')).toBe(NAME);
    expect(run('NOSUCHNAME')).toBe(NAME);
  });

  it('report a reference off the sheet', () => {
    expect(run('XFE1')).toBe(REF);
    expect(run('A1048577')).toBe(REF);
  });
});

describe('the function library', () => {
  const column = { A1: 1, A2: 2, A3: 3, A4: 4 };

  it('sums a range, a list, or both', () => {
    expect(run('SUM(A1:A4)', column)).toBe(10);
    expect(run('SUM(1,2,3)')).toBe(6);
    expect(run('SUM(A1:A2,10)', column)).toBe(13);
  });

  /**
   * Text and blanks in a range are skipped rather than being an error,
   * which is what makes `SUM(A1:A100)` usable on a column with a
   * heading in it. Passing text *directly* is a different claim and is
   * an error.
   */
  it('skips text inside a range and refuses it as an argument', () => {
    expect(run('SUM(A1:A4)', { ...column, A3: 'total' })).toBe(7);
    expect(run('SUM("x")')).toBe(VALUE);
  });

  it('averages, and has no answer for nothing', () => {
    expect(run('AVERAGE(A1:A4)', column)).toBe(2.5);
    expect(run('AVERAGE(B1:B4)')).toBe(DIV0);
  });

  it('takes a minimum and a maximum', () => {
    expect(run('MIN(A1:A4)', column)).toBe(1);
    expect(run('MAX(A1:A4)', column)).toBe(4);
  });

  it('counts numbers and nothing else', () => {
    expect(run('COUNT(A1:A4)', { A1: 1, A2: 'two', A3: null, A4: true })).toBe(1);
  });

  /**
   * COUNT does not propagate errors. Counting is the one thing still
   * possible over a range with a broken cell in it, and a COUNT that
   * reported `#DIV/0!` would fail exactly when the count is what you
   * need to find the problem.
   */
  it('counts past an error', () => {
    expect(run('COUNT(A1:A4)', { A1: 1, A2: DIV0, A3: 3, A4: 4 })).toBe(3);
  });

  it('rounds away from zero on a tie, as a sheet does', () => {
    expect(run('ROUND(2.5)')).toBe(3);
    // Math.round(-2.5) is -2, which is not what a spreadsheet answers.
    expect(run('ROUND(0-2.5)')).toBe(-3);
    expect(run('ROUND(3.14159,2)')).toBe(3.14);
    expect(run('ROUND(1234,0-2)')).toBe(1200);
  });

  it('takes an absolute value', () => {
    expect(run('ABS(0-3)')).toBe(3);
    expect(run('ABS(3)')).toBe(3);
  });

  it('concatenates values and ranges', () => {
    expect(run('CONCAT("a","b")')).toBe('ab');
    expect(run('CONCAT(A1:A4)', column)).toBe('1234');
  });

  it('reports the wrong number of arguments', () => {
    expect(run('ABS()')).toBe(VALUE);
    expect(run('ABS(1,2)')).toBe(VALUE);
  });
});

describe('IF', () => {
  it('chooses a branch', () => {
    expect(run('IF(TRUE,"y","n")')).toBe('y');
    expect(run('IF(FALSE,"y","n")')).toBe('n');
    expect(run('IF(A1>2,"big","small")', { A1: 5 })).toBe('big');
  });

  /**
   * The reason IF is not in the function table. `=IF(B1=0,"n/a",A1/B1)`
   * is how everyone writes a guarded division, and a sheet that
   * evaluated both branches would answer `#DIV/0!` to the formula
   * written specifically to avoid it.
   */
  it('does not evaluate the branch it did not take', () => {
    expect(run('IF(B1=0,"n/a",A1/B1)', { A1: 1, B1: 0 })).toBe('n/a');
  });

  it('is FALSE when the branch is missing', () => {
    expect(run('IF(FALSE,"y")')).toBe(false);
  });

  it('propagates an error in the condition', () => {
    expect(run('IF(A1,"y","n")', { A1: REF })).toBe(REF);
  });
});

describe('what the screen shows', () => {
  it('writes an error as its code', () => {
    expect(formatValue(REF)).toBe('#REF!');
  });

  it('writes an empty cell as nothing', () => {
    expect(formatValue(null)).toBe('');
  });

  it('does not show the representation error of a binary float', () => {
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
    expect(formatValue(1 / 3)).toBe('0.333333333333333');
  });
});
