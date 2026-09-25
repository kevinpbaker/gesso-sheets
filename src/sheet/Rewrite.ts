import type { CellRef } from './A1';
import type { Ast } from './Ast';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { printFormula } from './Print';

/**
 * Moving a formula, with its relative references adjusted and its
 * absolute ones left alone.
 *
 * This is what the `$` flags have been carried for since Phase 1: a
 * fill of `=B2*C2` down a column has to become `=B3*C3`, while
 * `=ROUND(D2/$D$7*100,1)` becomes `=ROUND(D3/$D$7*100,1)` — the total
 * stays put because somebody wrote a dollar sign to say so. Without
 * the flags there is no way to tell the two apart after parsing, which
 * is why `parseRef` keeps them rather than resolving them.
 *
 * A reference pushed off the sheet becomes `#REF!` in the text, which
 * is what a spreadsheet writes there: the formula stays readable and
 * says which part of it no longer points anywhere.
 */
export function rewriteFormula(input: string, rowDelta: number, columnDelta: number): string {
  if (!input.startsWith('=')) {
    return input;
  }
  if (rowDelta === 0 && columnDelta === 0) {
    return input;
  }
  let formula: Ast;
  try {
    formula = parseFormula(input.slice(1));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    // Unparseable text is moved unchanged. It is already showing an
    // error; mangling it further would lose what somebody typed.
    return input;
  }
  return `=${printFormula(shift(formula, rowDelta, columnDelta))}`;
}

function shift(node: Ast, rowDelta: number, columnDelta: number): Ast {
  switch (node.kind) {
    case 'ref':
      return { kind: 'ref', ref: shiftRef(node.ref, rowDelta, columnDelta) };
    case 'range':
      return {
        kind: 'range',
        range: {
          start: shiftRef(node.range.start, rowDelta, columnDelta),
          end: shiftRef(node.range.end, rowDelta, columnDelta)
        }
      };
    case 'call':
      return { kind: 'call', name: node.name, args: node.args.map(arg => shift(arg, rowDelta, columnDelta)) };
    case 'unary':
      return { kind: 'unary', op: node.op, operand: shift(node.operand, rowDelta, columnDelta) };
    case 'binary':
      return {
        kind: 'binary',
        op: node.op,
        left: shift(node.left, rowDelta, columnDelta),
        right: shift(node.right, rowDelta, columnDelta)
      };
    default:
      return node;
  }
}

/** Only the halves that were written without a `$` move. */
function shiftRef(ref: CellRef, rowDelta: number, columnDelta: number): CellRef {
  return {
    row: ref.rowAbsolute ? ref.row : ref.row + rowDelta,
    column: ref.columnAbsolute ? ref.column : ref.column + columnDelta,
    rowAbsolute: ref.rowAbsolute,
    columnAbsolute: ref.columnAbsolute
  };
}
