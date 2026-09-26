import { describe, expect, it } from 'vitest';

import { relativeRef } from './A1';
import type { Ast, BinaryOperator } from './Ast';
import { parseFormula } from './Parser';
import { printFormula } from './Print';

/**
 * The printer puts back only the brackets the tree needs, and this is
 * what keeps it honest: whatever it prints, the parser has to read back
 * as the tree it was given. The hand-picked cases are the corners of
 * the precedence table; the random ones are everything else.
 */

const print = (source: string): string => printFormula(parseFormula(source));

describe('printing a formula', () => {
  it.each([
    ['B2*C2', 'B2*C2'],
    ['(B2*C2)', 'B2*C2'],
    ['(A1+B1)*2', '(A1+B1)*2'],
    ['A1+B1*2', 'A1+B1*2'],
    ['A1-(B1-C1)', 'A1-(B1-C1)'],
    ['(A1-B1)-C1', 'A1-B1-C1'],
    ['A1/(B1*C1)', 'A1/(B1*C1)'],
    // ^ groups to the right, so it is the left side that needs them.
    ['2^3^2', '2^3^2'],
    ['(2^3)^2', '(2^3)^2'],
    // Unary minus binds tighter than ^, as in every spreadsheet.
    ['-2^2', '-2^2'],
    ['-(2^2)', '-(2^2)'],
    ['-(A1+B1)', '-(A1+B1)'],
    ['2^-1', '2^-1'],
    ['A1&B1&"!"', 'A1&B1&"!"'],
    ['(A1&B1)=C1', 'A1&B1=C1'],
    ['A1=(B1=C1)', 'A1=(B1=C1)'],
    ['SUM(A1:A5)*(1+B1)', 'SUM(A1:A5)*(1+B1)'],
    ['IF(A1>0,A1*2,-A1)', 'IF(A1>0,A1*2,-A1)']
  ])('%s prints as %s', (source, printed) => {
    expect(print(source)).toBe(printed);
  });
});

/** A tree, of the depth asked for, from a seeded generator. */
function randomTree(next: () => number, depth: number): Ast {
  const pick = <T>(items: readonly T[]): T => items[Math.floor(next() * items.length)];
  if (depth === 0 || next() < 0.25) {
    return pick<() => Ast>([
      () => ({ kind: 'number', value: Math.floor(next() * 100) }),
      () => ({ kind: 'ref', ref: relativeRef(Math.floor(next() * 20), Math.floor(next() * 5)) }),
      () => ({ kind: 'text', value: pick(['a', 'b c', '"q"']) })
    ])();
  }
  const roll = next();
  if (roll < 0.15) {
    return { kind: 'unary', op: pick(['-', '+'] as const), operand: randomTree(next, depth - 1) };
  }
  if (roll < 0.25) {
    return { kind: 'call', name: 'SUM', args: [randomTree(next, depth - 1), randomTree(next, depth - 1)] };
  }
  const ops: readonly BinaryOperator[] = ['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>=', '<=', '>'];
  return { kind: 'binary', op: pick(ops), left: randomTree(next, depth - 1), right: randomTree(next, depth - 1) };
}

/** A small deterministic generator, so a failure is the same failure next time. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state * 1664525 + 1013904223) >>> 0;
    return state / 0x100000000;
  };
}

describe('what the printer prints, the parser reads back', () => {
  it('for four thousand random trees', () => {
    const next = seeded(16);
    for (let at = 0; at < 4_000; at++) {
      const tree = randomTree(next, 5);
      const printed = printFormula(tree);
      expect(parseFormula(printed), printed).toEqual(tree);
    }
  });
});
