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
    expect(tokenize('1 2.5 .5 1e3 1E-2').filter(t => t.kind === 'number').map(t => t.value)).toEqual([
      1, 2.5, 0.5, 1000, 0.01
    ]);
  });

  /**
   * Positions are part of the contract, not an implementation detail.
   *
   * Phase 12 is built almost entirely out of the question "what is
   * under the caret" — which reference to colour, which argument of
   * which call the caret sits in, which bracket matches the one beside
   * it — and none of that can be asked of a token list that has
   * forgotten where the tokens were. Re-deriving the positions by
   * scanning the text again would be a second scanner to disagree
   * with this one.
   */
  it('says where each token was', () => {
    expect(tokenize('1 + B2').map(t => [t.kind, t.start, t.end])).toEqual([
      ['number', 0, 1],
      ['operator', 2, 3],
      ['word', 4, 6],
      ['end', 6, 6]
    ]);
  });

  it('spans a quoted string from quote to quote', () => {
    const [text] = tokenize('"ab" + 1');
    expect([text.start, text.end]).toEqual([0, 4]);
  });

  it('puts the end token at the end, with no width', () => {
    const tokens = tokenize('1+2');
    const last = tokens[tokens.length - 1];
    expect([last.kind, last.start, last.end]).toEqual(['end', 3, 3]);
  });

  it('reads a quoted string, in which a doubled quote is one quote', () => {
    expect(tokenize('"he said ""no"""')[0]).toMatchObject({ kind: 'text', value: 'he said "no"' });
  });

  it('reads the six error values as values', () => {
    for (const code of ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#CIRC!', '#N/A'] as const) {
      expect(tokenize(code)[0]).toMatchObject({ kind: 'error', code });
    }
  });

  /**
   * `LOG10` is a good function name and a good cell address, and only
   * what follows says which. The lexer refuses to guess, which is why
   * both arrive as `word`.
   */
  it('does not try to tell a reference from a function name', () => {
    expect(tokenize('SUM')[0]).toMatchObject({ kind: 'word', value: 'SUM' });
    expect(tokenize('A1')[0]).toMatchObject({ kind: 'word', value: 'A1' });
    expect(tokenize('LOG10')[0]).toMatchObject({ kind: 'word', value: 'LOG10' });
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

/**
 * `Sheet2!A1` — the third axis, in the text.
 *
 * The sheet travels as the **name** somebody typed rather than as an
 * id, so a formula prints back as what was written and a rename has
 * something to rewrite. What it never becomes is a range through the
 * workbook: `Sheet1!A1:Sheet2!B9` is refused rather than quietly
 * resolved to one sheet or the other.
 */
describe('a reference that names its sheet', () => {
  const parse = (text: string) => parseFormula(text);
  const sheetOf = (text: string): string | undefined => {
    const node = parse(text);
    if (node.kind === 'ref') {
      return node.ref.sheet;
    }
    if (node.kind === 'range') {
      return node.range.start.sheet;
    }
    throw new Error(`${text} is not a reference`);
  };

  it('reads a bare sheet name', () => {
    expect(sheetOf('Sheet2!A1')).toBe('Sheet2');
    expect(show(parse('Sheet2!A1'))).toBe('0:0');
  });

  it('reads a quoted one, spaces and all', () => {
    expect(sheetOf("'Q3 Budget'!A1")).toBe('Q3 Budget');
  });

  it('reads a quote inside a quoted one', () => {
    expect(sheetOf("'Kevin''s'!A1")).toBe("Kevin's");
  });

  it('qualifies both ends of a range from one name', () => {
    const node = parse('Sheet2!A1:B9');
    if (node.kind !== 'range') {
      throw new Error('not a range');
    }
    expect(node.range.start.sheet).toBe('Sheet2');
    expect(node.range.end.sheet).toBe('Sheet2');
  });

  it('takes the name written on both ends', () => {
    expect(sheetOf('Sheet2!A1:Sheet2!B9')).toBe('Sheet2');
  });

  it('takes a whole column on another sheet', () => {
    const node = parse('Sheet2!C:D');
    if (node.kind !== 'range') {
      throw new Error('not a range');
    }
    expect(node.range.wholeColumn).toBe(true);
    expect(node.range.start.sheet).toBe('Sheet2');
  });

  it('leaves a plain reference with no sheet at all', () => {
    expect(sheetOf('A1')).toBeUndefined();
    expect('sheet' in (parse('A1') as { ref: object }).ref).toBe(false);
  });

  /** The qualifier belongs to the reference, not to the formula. */
  it('works through calls and operators, and only where it was written', () => {
    const node = parse('SUM(Sheet2!A1:A9)+B1');
    expect(show(node)).toBe('(SUM(range(0:0,0:8)) + 1:0)');

    const sheets: (string | undefined)[] = [];
    referencesOf(node, {
      ref: ref => sheets.push(ref.sheet),
      range: range => sheets.push(range.start.sheet)
    });
    expect(sheets).toEqual(['Sheet2', undefined]);
  });

  it('refuses a range that runs from one sheet to another', () => {
    expect(() => parse('Sheet1!A1:Sheet2!B9')).toThrow(FormulaSyntaxError);
  });

  it('refuses a sheet in front of a function', () => {
    expect(() => parse('Sheet2!SUM(A1:A9)')).toThrow(FormulaSyntaxError);
  });

  it('refuses a sheet in front of a name', () => {
    expect(() => parse('Sheet2!Sales')).toThrow(FormulaSyntaxError);
  });

  it('refuses a quoted name with no reference after it', () => {
    expect(() => parse("'Q3 Budget'")).toThrow(FormulaSyntaxError);
    expect(() => parse("'Q3 Budget'+1")).toThrow(FormulaSyntaxError);
  });

  it('refuses an empty name', () => {
    expect(() => parse("''!A1")).toThrow(FormulaSyntaxError);
  });

  it('refuses a quoted name that is never closed', () => {
    expect(() => parse("'Q3!A1")).toThrow(FormulaSyntaxError);
  });
});
