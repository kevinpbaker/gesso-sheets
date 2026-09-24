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
