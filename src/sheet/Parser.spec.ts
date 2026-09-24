import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import type { Ast } from './Ast';
import { referencesOf } from './Ast';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { tokenize } from './Tokenizer';

/** The tree as a compact string, so precedence is readable in a spec. */
function show(node: Ast): string {
  switch (node.kind) {
    case 'number':
      return String(node.value);
    case 'text':
      return JSON.stringify(node.value);
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE';
    case 'error':
      return node.code;
    case 'ref':
      return `${node.ref.columnAbsolute ? '$' : ''}${node.ref.column}:${node.ref.rowAbsolute ? '$' : ''}${node.ref.row}`;
    case 'range':
      return `range(${show({ kind: 'ref', ref: node.range.start })},${show({ kind: 'ref', ref: node.range.end })})`;
    case 'call':
      return `${node.name}(${node.args.map(show).join(',')})`;
    case 'unary':
      return `${node.op}${show(node.operand)}`;
    case 'binary':
      return `(${show(node.left)} ${node.op} ${show(node.right)})`;
  }
}

describe('the tokenizer', () => {
  it('reads numbers, including exponents', () => {
    expect(tokenize('1 2.5 .5 1e3 1E-2').filter(t => t.kind === 'number')).toEqual([
      { kind: 'number', value: 1 },
      { kind: 'number', value: 2.5 },
      { kind: 'number', value: 0.5 },
      { kind: 'number', value: 1000 },
      { kind: 'number', value: 0.01 }
    ]);
  });

  it('reads a quoted string, in which a doubled quote is one quote', () => {
    expect(tokenize('"he said ""no"""')[0]).toEqual({ kind: 'text', value: 'he said "no"' });
  });

  it('reads the five error values as values', () => {
    for (const code of ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#CIRC!'] as const) {
      expect(tokenize(code)[0]).toEqual({ kind: 'error', code });
    }
  });

  /**
   * `LOG10` is a good function name and a good cell address, and only
   * what follows says which. The lexer refuses to guess, which is why
   * both arrive as `word`.
   */
  it('does not try to tell a reference from a function name', () => {
    expect(tokenize('SUM')[0]).toEqual({ kind: 'word', value: 'SUM' });
    expect(tokenize('A1')[0]).toEqual({ kind: 'word', value: 'A1' });
    expect(tokenize('LOG10')[0]).toEqual({ kind: 'word', value: 'LOG10' });
  });

  it('refuses a string that is never closed', () => {
    expect(() => tokenize('"open')).toThrow(FormulaSyntaxError);
  });
});

describe('the parser', () => {
  it('gives multiplication tighter binding than addition', () => {
    expect(show(parseFormula('1+2*3'))).toBe('(1 + (2 * 3))');
    expect(show(parseFormula('(1+2)*3'))).toBe('((1 + 2) * 3)');
  });

  it('leaves addition and subtraction left-associative', () => {
    expect(show(parseFormula('1-2-3'))).toBe('((1 - 2) - 3)');
  });

  it('makes the power right-associative', () => {
    expect(show(parseFormula('2^3^2'))).toBe('(2 ^ (3 ^ 2))');
  });

  /**
   * The one place this follows the spreadsheets rather than the
   * mathematics. `=-2^2` is 4 in Excel and in Sheets because unary
   * minus binds tighter than the power. It is a wart; copying it is
   * the point, because a sheet that disagreed with every other sheet
   * about a formula this short would be wrong in the way that matters.
   */
  it('binds unary minus tighter than the power, as a spreadsheet does', () => {
    expect(show(parseFormula('-2^2'))).toBe('(-2 ^ 2)');
  });

  it('puts concatenation below arithmetic and above comparison', () => {
    expect(show(parseFormula('1+2&"x"'))).toBe('((1 + 2) & "x")');
    expect(show(parseFormula('"a"&"b"="ab"'))).toBe('(("a" & "b") = "ab")');
  });

  it('reads the comparison operators, including the two-character ones', () => {
    expect(show(parseFormula('1<>2'))).toBe('(1 <> 2)');
    expect(show(parseFormula('1<=2'))).toBe('(1 <= 2)');
    expect(show(parseFormula('1>=2'))).toBe('(1 >= 2)');
  });

  it('reads a call, with no arguments and with several', () => {
    expect(show(parseFormula('SUM()'))).toBe('SUM()');
    expect(show(parseFormula('sum(1,2,3)'))).toBe('SUM(1,2,3)');
    expect(show(parseFormula('IF(A1>0,"y","n")'))).toBe('IF((0:0 > 0),"y","n")');
  });

  it('reads a reference and a range', () => {
    expect(parseFormula('A1')).toEqual({ kind: 'ref', ref: relativeRef(0, 0) });
    expect(show(parseFormula('A1:B2'))).toBe('range(0:0,1:1)');
    expect(show(parseFormula('$A$1:$B$2'))).toBe('range($0:$0,$1:$1)');
  });

  it('decides between a call and a reference on the bracket', () => {
    // Column LOG is 8508 in bijective base 26, row 10 is index 9.
    expect(show(parseFormula('LOG10'))).toBe('8508:9');
    expect(show(parseFormula('LOG10(2)'))).toBe('LOG10(2)');
  });

  it('reads TRUE and FALSE as booleans, not as cells', () => {
    expect(parseFormula('TRUE')).toEqual({ kind: 'boolean', value: true });
    expect(parseFormula('false')).toEqual({ kind: 'boolean', value: false });
  });

  /**
   * A name with no meaning still parses. `=NOSUCHNAME` is a `#NAME?`
   * in the cell, not a refusal to accept what was typed — which is
   * where Excel puts it too.
   */
  it('parses an unknown name into a call that will fail at evaluation', () => {
    expect(show(parseFormula('NOSUCHNAME'))).toBe('NOSUCHNAME()');
  });

  it('ignores whitespace', () => {
    expect(show(parseFormula('  1  +  2  '))).toBe('(1 + 2)');
  });

  it('refuses what is not a formula', () => {
    for (const source of ['1+', '(1', '1)', 'SUM(1,', '*', 'A1:', 'A1:SUM']) {
      expect(() => parseFormula(source), source).toThrow(FormulaSyntaxError);
    }
  });
});

describe('the references a formula reads', () => {
  function references(source: string): string[] {
    const found: string[] = [];
    referencesOf(parseFormula(source), {
      ref: ref => found.push(`${ref.row},${ref.column}`),
      range: range => found.push(`${range.start.row},${range.start.column}:${range.end.row},${range.end.column}`)
    });
    return found;
  }

  it('finds them through operators and calls', () => {
    expect(references('A1+B2*SUM(C3:D4)')).toEqual(['0,0', '1,1', '2,2:3,3']);
  });

  /**
   * Both branches of an IF, and this is not an oversight. A cell that
   * only depended on the branch currently taken would hold a stale
   * value the moment the condition flipped, because nothing would have
   * marked it dirty when the untaken branch changed.
   */
  it('finds them in branches that are not taken', () => {
    expect(references('IF(A1,B1,C1)')).toEqual(['0,0', '0,1', '0,2']);
  });

  it('finds none in a formula that reads no cells', () => {
    expect(references('1+2*3')).toEqual([]);
  });
});
