import type { CellRef, RangeRef } from './A1';
import type { Ast } from './Ast';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { printFormula } from './Print';

/**
 * A block of cells moved somewhere else: a cut, pasted.
 *
 * Excel's rule, which is not a fill's: a formula that pointed *into*
 * the block points at the same cells in their new place, wherever that
 * formula is — and a moved formula's other references do not move at
 * all, because the cells they name did not. So the rewrite is by what
 * a reference names, not by where the formula sits. A range moves only
 * when all of it was inside the block; one that was partly inside
 * keeps its corners, as it would in Excel.
 */
export interface Move {
  /** The sheet the block was on. */
  readonly sheet: string;
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
  readonly rowDelta: number;
  readonly columnDelta: number;
  /** Where it went, when that is another sheet; absent for the same one. */
  readonly toSheet?: string;
}

/**
 * A formula with every reference into the moved block pointing at its
 * new place.
 *
 * `carried` is for a formula that is itself in the block, going to
 * another sheet: its references to anything outside the block named
 * their sheet by leaving it out, and have to say it now, or they would
 * start reading the sheet they arrived on.
 */
export function moveFormula(input: string, move: Move, onSheet: string, carried = false): string {
  if (!input.startsWith('=')) {
    return input;
  }
  let formula: Ast;
  try {
    formula = parseFormula(input.slice(1));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    return input;
  }
  const moved = moveAst(formula, move, onSheet, carried && move.toSheet !== undefined);
  return moved === formula ? input : `=${printFormula(moved)}`;
}

function on(ref: CellRef, move: Move, onSheet: string): boolean {
  return (ref.sheet ?? onSheet).toUpperCase() === move.sheet.toUpperCase();
}

function inside(ref: CellRef, move: Move): boolean {
  return ref.row >= move.firstRow && ref.row <= move.lastRow && ref.column >= move.firstColumn && ref.column <= move.lastColumn;
}

function moved(ref: CellRef, move: Move, onSheet: string): CellRef {
  const next: CellRef = { ...ref, row: ref.row + move.rowDelta, column: ref.column + move.columnDelta };
  if (move.toSheet === undefined || move.toSheet.toUpperCase() === (ref.sheet ?? onSheet).toUpperCase()) {
    return next;
  }
  // To another sheet: a reference that named its own sheet by leaving
  // the sheet out has to name the new one now.
  return { ...next, sheet: move.toSheet };
}

function moveAst(node: Ast, move: Move, onSheet: string, qualify: boolean): Ast {
  const pinned = (ref: CellRef): CellRef => (ref.sheet === undefined ? { ...ref, sheet: onSheet } : ref);
  switch (node.kind) {
    case 'ref':
      if (on(node.ref, move, onSheet) && inside(node.ref, move)) {
        return { kind: 'ref', ref: moved(node.ref, move, onSheet) };
      }
      return qualify && node.ref.sheet === undefined ? { kind: 'ref', ref: pinned(node.ref) } : node;
    case 'range': {
      const range = node.range;
      if (range.wholeColumn === true || !on(range.start, move, onSheet) || !inside(range.start, move) || !inside(range.end, move)) {
        return qualify && range.start.sheet === undefined
          ? { kind: 'range', range: { ...range, start: pinned(range.start), end: pinned(range.end) } }
          : node;
      }
      const next: RangeRef = { ...range, start: moved(range.start, move, onSheet), end: moved(range.end, move, onSheet) };
      return { kind: 'range', range: next };
    }
    case 'call': {
      let changed = false;
      const args = node.args.map(arg => {
        const next = moveAst(arg, move, onSheet, qualify);
        changed ||= next !== arg;
        return next;
      });
      return changed ? { ...node, args } : node;
    }
    case 'invoke': {
      const callee = moveAst(node.callee, move, onSheet, qualify);
      let changed = callee !== node.callee;
      const args = node.args.map(arg => {
        const next = moveAst(arg, move, onSheet, qualify);
        changed ||= next !== arg;
        return next;
      });
      return changed ? { kind: 'invoke', callee, args } : node;
    }
    case 'unary': {
      const operand = moveAst(node.operand, move, onSheet, qualify);
      return operand === node.operand ? node : { kind: 'unary', op: node.op, operand };
    }
    case 'binary': {
      const left = moveAst(node.left, move, onSheet, qualify);
      const right = moveAst(node.right, move, onSheet, qualify);
      return left === node.left && right === node.right ? node : { kind: 'binary', op: node.op, left, right };
    }
    default:
      return node;
  }
}
