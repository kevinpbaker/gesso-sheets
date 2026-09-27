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
  | CallNode
  | { readonly kind: 'unary'; readonly op: UnaryOperator; readonly operand: Ast }
  | { readonly kind: 'binary'; readonly op: BinaryOperator; readonly left: Ast; readonly right: Ast }
  /**
   * A value called: `LAMBDA(x, x*2)(A1)`, or a `LAMBDA` a name or a
   * `LET` handed back, called where it came out.
   */
  | { readonly kind: 'invoke'; readonly callee: Ast; readonly args: readonly Ast[] };

/**
 * A function called by name, or a bare word.
 *
 * A bare word — `Sales`, `TaxRate`, the `x` of `LAMBDA(x, x*2)` —
 * is a call with no arguments that carries `word`, the spelling it was
 * typed with. It is what tells `Sales` from `Sales()`, which are two
 * different things to say: the first is a name and the second is a
 * call to a function that does not exist. It is also what prints back,
 * so a name filled down a column keeps the case it was typed in.
 */
export interface CallNode {
  readonly kind: 'call';
  readonly name: string;
  readonly args: readonly Ast[];
  readonly word?: string;
  /** A LAMBDA's optional parameter, written `[b]`. Only ever on a bare word. */
  readonly optional?: true;
}

/** A bare word, by the name it is looked up under (upper case), or null. */
export function wordOf(node: Ast): string | null {
  return node.kind === 'call' && node.word !== undefined ? node.name : null;
}

/** The functions whose first arguments are names they bind, not values. */
export const BINDERS: ReadonlySet<string> = new Set(['LET', 'LAMBDA']);

/**
 * Every name a `LET` or a `LAMBDA` in a formula binds, upper-cased.
 *
 * Not scoped: a name bound anywhere in the formula is in the set. It is
 * for the two questions about a whole formula that need it — whether a
 * file's formula mentions only things this sheet knows, and which words
 * the file format marks as parameters — and neither is about where.
 */
export function bindingNamesOf(node: Ast, into = new Set<string>()): Set<string> {
  switch (node.kind) {
    case 'call': {
      if (node.name === 'LAMBDA') {
        node.args.slice(0, -1).forEach(arg => {
          const word = wordOf(arg);
          if (word !== null) {
            into.add(word);
          }
        });
      } else if (node.name === 'LET') {
        node.args.forEach((arg, at) => {
          const word = at % 2 === 0 && at < node.args.length - 1 ? wordOf(arg) : null;
          if (word !== null) {
            into.add(word);
          }
        });
      }
      node.args.forEach(arg => bindingNamesOf(arg, into));
      return into;
    }
    case 'invoke':
      bindingNamesOf(node.callee, into);
      node.args.forEach(arg => bindingNamesOf(arg, into));
      return into;
    case 'unary':
      return bindingNamesOf(node.operand, into);
    case 'binary':
      bindingNamesOf(node.left, into);
      return bindingNamesOf(node.right, into);
    default:
      return into;
  }
}

/** `@` is implicit intersection: the one value of a range, where Excel 365 writes it. */
export type UnaryOperator = '-' | '+' | '@';

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
    case 'invoke':
      referencesOf(node.callee, into);
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
    case 'invoke':
      callNamesOf(node.callee, into);
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
    case 'invoke':
      bareWordsOf(node.callee, into);
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
