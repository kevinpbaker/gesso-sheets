import type { CellRef } from './A1';
import type { Ast } from './Ast';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { printFormula } from './Print';

/**
 * Moving a formula because the rows underneath it were reordered.
 *
 * This is not `rewriteFormula`, and the difference is the whole point.
 * A fill *makes a new formula* one row down, so its relative references
 * should travel the same distance: `=B2*C2` filled into row 3 means
 * `=B3*C3`. A sort *permutes rows that already exist*. Nothing new is
 * written, and a reference only has to move if the cell it names is one
 * of the cells the sort picked up and put somewhere else.
 *
 * Sorting with fill semantics is how a total row gets destroyed. Give
 * `=SUM(B2:B4)` to a sort that moves its row three rows up and the fill
 * rule says `=SUM(B-1:B1)`, which is `#REF!`; move it one row up and the
 * rule says `=SUM(B1:B3)`, which is not an error and is adding up the
 * heading and two of the three rows. The second is the worse one. Both
 * came out of the demo sheet by accident, which is the only reason
 * either was found.
 *
 * So: a reference to a row the sort moved follows that row to where it
 * went, and everything else is left exactly as it was typed.
 *
 * Two consequences worth stating plainly.
 *
 * `$` does not protect a reference here. A `$` says "do not drift when
 * I am copied", and this is not a copy — if `$D$7` names a row that the
 * sort carried to row 3, then `$D$3` is where that number now is, and
 * leaving it at 7 would point it at whatever row took its place. This
 * matches `Shift`, where an insert moves `$` references too.
 *
 * **Ranges are never rewritten.** A range names a region of the sheet,
 * and a sort permutes rows *within* a region: every row that was inside
 * `B2:B4` is still inside it afterwards, so the set of cells it covers
 * holds the same values and the range still means what it meant. This
 * is exactly why the total row survives. The limit is a range covering
 * only part of what was sorted — `=SUM(B2:B3)` out of four sorted rows
 * still reads rows 2 and 3, which now hold different data. There is no
 * rewrite that saves that one: an arbitrary permutation does not map a
 * partial range onto any range at all. Leaving it alone is the answer
 * that keeps the sheet honest — the range says where it reads, and it
 * still reads there — rather than inventing a region nobody wrote.
 */
export function permuteFormula(
  input: string,
  moved: ReadonlyMap<number, number>,
  columns: ColumnSpan
): string {
  if (!input.startsWith('=') || moved.size === 0) {
    return input;
  }
  let formula: Ast;
  try {
    formula = parseFormula(input.slice(1));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    // Unparseable text moves unchanged, as it does in a fill: it is
    // already showing an error and mangling it further loses what
    // somebody typed.
    return input;
  }
  const permuted = permute(formula, moved, columns);
  // Reprinting costs the formula its original spacing and adds the
  // printer's parentheses, so a formula that names nothing the sort
  // touched is handed back untouched rather than laundered through the
  // printer for no reason.
  return permuted === null ? input : `=${printFormula(permuted)}`;
}

/** The rewritten tree, or null when no reference in it moved. */
function permute(node: Ast, moved: ReadonlyMap<number, number>, columns: ColumnSpan): Ast | null {
  switch (node.kind) {
    case 'ref': {
      const ref = permuteRef(node.ref, moved, columns);
      return ref === null ? null : { kind: 'ref', ref };
    }
    case 'range':
      // Left alone, always. See the note above.
      return null;
    case 'call': {
      const args = node.args.map(arg => permute(arg, moved, columns));
      return args.every(arg => arg === null)
        ? null
        : { kind: 'call', name: node.name, args: node.args.map((arg, at) => args[at] ?? arg) };
    }
    case 'unary': {
      const operand = permute(node.operand, moved, columns);
      return operand === null ? null : { kind: 'unary', op: node.op, operand };
    }
    case 'binary': {
      const left = permute(node.left, moved, columns);
      const right = permute(node.right, moved, columns);
      return left === null && right === null
        ? null
        : { kind: 'binary', op: node.op, left: left ?? node.left, right: right ?? node.right };
    }
    default:
      return null;
  }
}

/**
 * The reference's new home, or null when the sort did not move that cell.
 *
 * Both halves have to be checked, and forgetting the column is a bug a
 * spec caught: a sort of columns A and B does not touch column E, so
 * `$E$1` names a cell that is exactly where it was, however the rows
 * beside it were shuffled.
 */
function permuteRef(ref: CellRef, moved: ReadonlyMap<number, number>, columns: ColumnSpan): CellRef | null {
  if (ref.column < columns.first || ref.column > columns.last) {
    return null;
  }
  const to = moved.get(ref.row);
  if (to === undefined || to === ref.row) {
    return null;
  }
  return { ...ref, row: to };
}

/** The columns a sort picked up; cells outside them did not move. */
export interface ColumnSpan {
  readonly first: number;
  readonly last: number;
}

/** A run of rows a formula reads, inclusive. */
export interface RowSpan {
  readonly first: number;
  readonly last: number;
}

/**
 * Which rows a formula reads, within a span of columns.
 *
 * Used to find the summary rows of a block — the total at the bottom
 * that reads the rows above it — so a sort can leave them where they
 * are instead of shuffling a total into the middle of its own data.
 * Unparseable text reads nothing; it is already an error.
 */
export function rowsRead(input: string, columns: ColumnSpan): readonly RowSpan[] {
  if (!input.startsWith('=')) {
    return [];
  }
  let formula: Ast;
  try {
    formula = parseFormula(input.slice(1));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    return [];
  }
  const spans: RowSpan[] = [];
  collect(formula, columns, spans);
  return spans;
}

function collect(node: Ast, columns: ColumnSpan, into: RowSpan[]): void {
  switch (node.kind) {
    case 'ref':
      if (node.ref.column >= columns.first && node.ref.column <= columns.last) {
        into.push({ first: node.ref.row, last: node.ref.row });
      }
      return;
    case 'range': {
      const { start, end } = node.range;
      const firstColumn = Math.min(start.column, end.column);
      const lastColumn = Math.max(start.column, end.column);
      if (lastColumn >= columns.first && firstColumn <= columns.last) {
        into.push({ first: Math.min(start.row, end.row), last: Math.max(start.row, end.row) });
      }
      return;
    }
    case 'call':
      node.args.forEach(arg => collect(arg, columns, into));
      return;
    case 'unary':
      collect(node.operand, columns, into);
      return;
    case 'binary':
      collect(node.left, columns, into);
      collect(node.right, columns, into);
      return;
    default:
  }
}
