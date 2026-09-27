import { formatRange, formatRef, inBounds, type RangeRef } from './A1';
import type { Ast, BinaryOperator } from './Ast';
import { formatNumber } from './Values';

/**
 * A parsed formula back as text.
 *
 * Shared by the two things that move a formula — `Rewrite` for a fill
 * and `Shift` for an insert — because two printers are two chances to
 * disagree about what a tree means, and the place they would disagree
 * is the place a formula silently changes.
 *
 * **With the brackets the tree needs and no others.** This used to
 * bracket every operation — `=B2*C2` filled down became `=(B3*C3)` —
 * on the argument that a printer dropping brackets must know the
 * precedence table as exactly as the parser does, and the one place the
 * two could differ is the one that matters. The argument was right and
 * so was its answer, which this file named: a printer that knows the
 * table, held to it by a spec that parses everything it prints back
 * and asks for the same tree — `Print.spec.ts`, over hand-picked cases
 * and a few thousand random trees. The table is `PRECEDENCE`, copied
 * from `Parser.ts` with the one asymmetry that matters: `^` groups to
 * the right, everything else to the left.
 */
export function printFormula(node: Ast): string {
  switch (node.kind) {
    case 'number':
      return formatNumber(node.value);
    case 'text':
      return `"${node.value.replace(/"/g, '""')}"`;
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE';
    case 'error':
      return node.code;
    case 'ref':
      return inBounds(node.ref.row, node.ref.column) ? formatRef(node.ref) : '#REF!';
    case 'range':
      return printRange(node.range);
    case 'call':
      if (node.word !== undefined) {
        // A name, or a name a LET or LAMBDA bound, as it was typed:
        // `Sales`, not `SALES()`, which would be a call to nothing.
        return node.word;
      }
      if (node.name === 'ANCHORARRAY' && node.args.length === 1 && node.args[0].kind === 'ref') {
        return `${printFormula(node.args[0])}#`;
      }
      return `${node.name}(${node.args.map(printFormula).join(',')})`;
    case 'invoke': {
      // A call or another invocation can be called as it stands; anything
      // else is bracketed first, as it had to be to be called at all.
      const callee = node.callee.kind === 'invoke' || (node.callee.kind === 'call' && node.callee.word === undefined)
        ? printFormula(node.callee)
        : `(${printFormula(node.callee)})`;
      return `${callee}(${node.args.map(printFormula).join(',')})`;
    }
    case 'unary':
      // Unary minus binds tighter than every binary operator, so any
      // operation under it needs its brackets back: -(A1+B1).
      return `${node.op}${node.operand.kind === 'binary' ? `(${printFormula(node.operand)})` : printFormula(node.operand)}`;
    case 'binary': {
      const own = PRECEDENCE[node.op];
      const right = node.op === '^';
      return `${operand(node.left, own, right)}${node.op}${operand(node.right, own, !right)}`;
    }
  }
}

/** How tightly each operator binds, loosest first; see `Parser.ts`. */
const PRECEDENCE: Readonly<Record<BinaryOperator, number>> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '<=': 1,
  '>': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5
};

/**
 * A child of a binary operation, bracketed when the parser would
 * otherwise group it differently.
 *
 * A looser child always needs them: `(A1+B1)*2`. An equally tight one
 * needs them on the side its associativity does not reach: `A1-(B1-C1)`
 * on the right of a left-grouping operator, `(2^3)^2` on the left of
 * `^`, which groups the other way.
 */
function operand(child: Ast, parent: number, againstAssociativity: boolean): string {
  const text = printFormula(child);
  if (child.kind !== 'binary') {
    return text;
  }
  const own = PRECEDENCE[child.op];
  return own < parent || (own === parent && againstAssociativity) ? `(${text})` : text;
}

/**
 * A range, or `#REF!` when either corner has left the sheet.
 *
 * The whole range and not the corner: half a range is not a range,
 * and `A1:#REF!` is not something anybody can read or the parser can
 * take back.
 */
function printRange(range: RangeRef): string {
  const off = !inBounds(range.start.row, range.start.column) || !inBounds(range.end.row, range.end.column);
  return off ? '#REF!' : formatRange(range);
}
