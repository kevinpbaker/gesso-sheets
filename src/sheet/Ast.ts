import type { CellRef, RangeRef } from './A1';
import type { ErrorCode } from './Values';

/**
 * A parsed formula.
 *
 * Deliberately a plain tagged union with no methods: the evaluator
 * walks it, the dependency scan walks it, and Phase 5's fill handle
 * will rewrite it. Three readers with different jobs, none of which
 * belongs on the node.
 */
export type Ast =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly code: ErrorCode }
  | { readonly kind: 'ref'; readonly ref: CellRef }
  | { readonly kind: 'range'; readonly range: RangeRef }
  | { readonly kind: 'call'; readonly name: string; readonly args: readonly Ast[] }
  | { readonly kind: 'unary'; readonly op: UnaryOperator; readonly operand: Ast }
  | { readonly kind: 'binary'; readonly op: BinaryOperator; readonly left: Ast; readonly right: Ast };

export type UnaryOperator = '-' | '+';

export type BinaryOperator = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '<=' | '>' | '>=';

/**
 * Every cell an expression reads, as references and ranges.
 *
 * The dependency graph is built from this rather than from evaluation,
 * so a cell's precedents are known whether or not the branch that
 * mentions them is taken: `=IF(A1, B1, C1)` depends on all three, and
 * must, or editing C1 while the condition is true would leave a stale
 * value waiting to appear the moment it flips.
 */
export function referencesOf(node: Ast, into: { ref(ref: CellRef): void; range(range: RangeRef): void }): void {
  switch (node.kind) {
    case 'ref':
      into.ref(node.ref);
      return;
    case 'range':
      into.range(node.range);
      return;
    case 'call':
      for (const arg of node.args) {
        referencesOf(arg, into);
      }
      return;
    case 'unary':
      referencesOf(node.operand, into);
      return;
    case 'binary':
      referencesOf(node.left, into);
      referencesOf(node.right, into);
      return;
    default:
      return;
  }
}

/**
 * Every function a formula calls, by name.
 *
 * The dependency graph is built from references, and two things it
 * cannot see are written as calls: a formula that is *volatile* reads
 * nothing yet must still be recalculated, and one that calls
 * `INDIRECT` or `OFFSET` reads cells that its own text never names.
 * Both are found here, once, when the formula is parsed — walking the
 * tree again on every recalculation would be the same answer at a
 * cost per evaluation rather than per edit.
 */
export function callNamesOf(node: Ast, into: Set<string>): void {
  switch (node.kind) {
    case 'call':
      into.add(node.name);
      for (const arg of node.args) {
        callNamesOf(arg, into);
      }
      return;
    case 'unary':
      callNamesOf(node.operand, into);
      return;
    case 'binary':
      callNamesOf(node.left, into);
      callNamesOf(node.right, into);
      return;
    default:
      return;
  }
}

/**
 * The bare words a formula uses that are not calls: candidate names.
 *
 * A named range parses as a call with no arguments, because the
 * parser cannot know a name from a misspelt function and deliberately
 * does not try. Which of these is a name is the *sheet's* question,
 * answered against its own table — this only says which words were
 * there.
 */
export function bareWordsOf(node: Ast, into: Set<string>): void {
  switch (node.kind) {
    case 'call':
      if (node.args.length === 0) {
        into.add(node.name);
      }
      for (const arg of node.args) {
        bareWordsOf(arg, into);
      }
      return;
    case 'unary':
      bareWordsOf(node.operand, into);
      return;
    case 'binary':
      bareWordsOf(node.left, into);
      bareWordsOf(node.right, into);
      return;
    default:
      return;
  }
}
