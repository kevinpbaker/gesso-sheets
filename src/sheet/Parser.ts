import { columnIndex, parseRef, wholeColumnRange, type CellRef } from './A1';
import type { Ast, BinaryOperator } from './Ast';
import { FormulaSyntaxError, tokenize, type Token } from './Tokenizer';

/**
 * A formula's text as a tree.
 *
 * Precedence, loosest first: comparison, `&`, `+ -`, `* /`, `^`,
 * unary `-`. That last pair is the one place this follows the
 * spreadsheets rather than the mathematics: `=-2^2` is **4** here, as
 * it is in Excel and Sheets, because unary minus binds tighter than
 * the power. It is a wart, and copying it is the point — a sheet that
 * disagreed with every other sheet about a formula this short would be
 * wrong in the only way that matters, whatever the textbooks say.
 *
 * `^` is right-associative, so `2^3^2` is 512.
 */
export function parseFormula(source: string): Ast {
  const parser = new Parser(tokenize(source));
  const expression = parser.expression();
  parser.expect('end');
  return expression;
}

export { FormulaSyntaxError };

const COMPARISONS = new Set<BinaryOperator>(['=', '<>', '<', '<=', '>', '>=']);

class Parser {
  private at = 0;

  constructor(private readonly tokens: readonly Token[]) {}

  expression(): Ast {
    return this.comparison();
  }

  private comparison(): Ast {
    let left = this.concatenation();
    for (;;) {
      const token = this.peek();
      if (token.kind !== 'operator' || !COMPARISONS.has(token.value as BinaryOperator)) {
        return left;
      }
      this.at++;
      left = { kind: 'binary', op: token.value as BinaryOperator, left, right: this.concatenation() };
    }
  }

  private concatenation(): Ast {
    let left = this.additive();
    while (this.isOperator('&')) {
      this.at++;
      left = { kind: 'binary', op: '&', left, right: this.additive() };
    }
    return left;
  }

  private additive(): Ast {
    let left = this.multiplicative();
    for (;;) {
      const token = this.peek();
      if (token.kind !== 'operator' || (token.value !== '+' && token.value !== '-')) {
        return left;
      }
      this.at++;
      left = { kind: 'binary', op: token.value, left, right: this.multiplicative() };
    }
  }

  private multiplicative(): Ast {
    let left = this.power();
    for (;;) {
      const token = this.peek();
      if (token.kind !== 'operator' || (token.value !== '*' && token.value !== '/')) {
        return left;
      }
      this.at++;
      left = { kind: 'binary', op: token.value, left, right: this.power() };
    }
  }

  /** Right-associative, and its operands are already sign-bearing. */
  private power(): Ast {
    const left = this.unary();
    if (!this.isOperator('^')) {
      return left;
    }
    this.at++;
    return { kind: 'binary', op: '^', left, right: this.power() };
  }

  private unary(): Ast {
    const token = this.peek();
    if (token.kind === 'operator' && (token.value === '-' || token.value === '+')) {
      this.at++;
      return { kind: 'unary', op: token.value, operand: this.unary() };
    }
    return this.primary();
  }

  private primary(): Ast {
    const token = this.peek();
    switch (token.kind) {
      case 'number':
        this.at++;
        return { kind: 'number', value: token.value };
      case 'text':
        this.at++;
        return { kind: 'text', value: token.value };
      case 'error':
        this.at++;
        return { kind: 'error', code: token.code };
      case 'open': {
        this.at++;
        const inner = this.expression();
        this.expect('close');
        return inner;
      }
      case 'word':
        this.at++;
        return this.word(token.value);
      default:
        throw new FormulaSyntaxError(`Expected a value, found ${describe(token)}.`);
    }
  }

  /**
   * A bare word, which is a call, a reference, a range, or TRUE/FALSE.
   *
   * The `(` decides between the first two, which is why the lexer does
   * not: `LOG10(2)` is a call and `LOG10` on its own is the cell in
   * column LOG, row 10, and both are things somebody writes.
   */
  private word(text: string): Ast {
    if (this.peek().kind === 'open') {
      this.at++;
      return { kind: 'call', name: text.toUpperCase(), args: this.arguments() };
    }
    const upper = text.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { kind: 'boolean', value: upper === 'TRUE' };
    }
    const ref = parseRef(text);
    if (ref === null) {
      const whole = this.maybeWholeColumns(text);
      if (whole !== null) {
        return whole;
      }
      // A name the sheet has no meaning for. It parses — the formula is
      // well formed — and fails at evaluation, which is where Excel
      // puts it too: `=NOSUCHNAME` is a `#NAME?` in the cell, not a
      // refusal to accept what was typed.
      return { kind: 'call', name: upper, args: [] };
    }
    return this.maybeRange(ref);
  }

  /**
   * `A:A`, or `B:D` — a column name, a colon, and another.
   *
   * Reached only after `parseRef` has said the word is not a cell,
   * which is what distinguishes `A:A` from `A1:A9`. Nothing is
   * consumed unless the whole shape is there, so a bare `A` still
   * falls through to being an unknown name.
   *
   * Whole *rows* — `1:3` — are deliberately not here. They would need
   * the parser to read a range out of two number tokens, which is a
   * second shape for a case nobody in this application has asked for;
   * `A1:Z1` says the same thing and says which columns it means.
   */
  private maybeWholeColumns(text: string): Ast | null {
    if (this.peek().kind !== 'colon') {
      return null;
    }
    const after = this.tokens[this.at + 1] ?? { kind: 'end' as const, start: 0, end: 0 };
    if (after.kind !== 'word') {
      return null;
    }
    const first = columnIndex(text.replace(/\$/g, ''));
    const last = columnIndex(after.value.replace(/\$/g, ''));
    if (first === null || last === null) {
      return null;
    }
    this.at += 2;
    return {
      kind: 'range',
      range: wholeColumnRange(Math.min(first, last), Math.max(first, last))
    };
  }

  private maybeRange(start: CellRef): Ast {
    if (this.peek().kind !== 'colon') {
      return { kind: 'ref', ref: start };
    }
    this.at++;
    const token = this.peek();
    if (token.kind !== 'word') {
      throw new FormulaSyntaxError(`A range needs a cell after the colon, found ${describe(token)}.`);
    }
    const end = parseRef(token.value);
    if (end === null) {
      throw new FormulaSyntaxError(`${token.value} is not a cell reference.`);
    }
    this.at++;
    return { kind: 'range', range: { start, end } };
  }

  private arguments(): Ast[] {
    const args: Ast[] = [];
    if (this.peek().kind === 'close') {
      this.at++;
      return args;
    }
    for (;;) {
      args.push(this.expression());
      const token = this.peek();
      if (token.kind === 'comma') {
        this.at++;
        continue;
      }
      this.expect('close');
      return args;
    }
  }

  private peek(): Token {
    return this.tokens[this.at] ?? { kind: 'end', start: 0, end: 0 };
  }

  private isOperator(value: string): boolean {
    const token = this.peek();
    return token.kind === 'operator' && token.value === value;
  }

  expect(kind: Token['kind']): void {
    const token = this.peek();
    if (token.kind !== kind) {
      throw new FormulaSyntaxError(`Expected ${kind}, found ${describe(token)}.`);
    }
    this.at++;
  }
}

function describe(token: Token): string {
  switch (token.kind) {
    case 'end':
      return 'the end of the formula';
    case 'operator':
      return `'${token.value}'`;
    case 'word':
      return `'${token.value}'`;
    case 'number':
      return String(token.value);
    case 'text':
      return JSON.stringify(token.value);
    case 'error':
      return token.code;
    default:
      return token.kind;
  }
}
