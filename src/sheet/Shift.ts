import { MAX_COLUMNS, MAX_ROWS, type CellRef, type RangeRef } from './A1';
import type { Ast } from './Ast';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { printFormula } from './Print';

/**
 * Moving references because the sheet changed shape.
 *
 * **This is not `Rewrite`, and the difference is the thing every
 * implementation gets wrong exactly once.** A fill moves the relative
 * references and pins the absolute ones — that is what `$` is for.
 * An insert moves *both*: `$A$1` means "the cell at A1", and when a
 * column appears to its left that cell is at B1, so the reference has
 * to say B1 or it now points at somebody else's data. The dollar sign
 * says "do not move when I am copied", not "do not notice the sheet".
 *
 * The other half is that shifting is **positional and not relative**.
 * A formula in row 1 reading `=A900` has to be rewritten when a row is
 * inserted at row 500, even though the formula itself did not move.
 * So an insert rewrites every formula in the sheet that references at
 * or beyond the line — which is what `Sheet.insertRows` walks, and
 * what the budget spec counts.
 */

export type Axis = 'row' | 'column';

export interface Shift {
  readonly axis: Axis;
  /** The first index inserted, or the first deleted. */
  readonly at: number;
  /** How many were inserted (positive) or deleted (negative). */
  readonly by: number;
}

/**
 * A formula with its references moved, or left alone.
 *
 * Returns the input unchanged when nothing in it moved, which is what
 * lets the caller count how many formulas an insert actually touched
 * rather than how many it looked at.
 */
export function shiftFormula(input: string, shift: Shift): string {
  if (!input.startsWith('=') || shift.by === 0) {
    return input;
  }
  let formula: Ast;
  try {
    formula = parseFormula(input.slice(1));
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
    // Unparseable text is left as typed. It is already showing an
    // error, and mangling it further would lose what somebody wrote.
    return input;
  }
  const moved = shiftAst(formula, shift);
  if (moved === formula) {
    return input;
  }
  return `=${printFormula(moved)}`;
}

/**
 * The tree with its references moved.
 *
 * Returns the *same node* when nothing under it changed, so that
 * `shiftFormula` can tell "this formula did not reference the line"
 * from "this formula referenced it and came back the same". A deep
 * equality check would answer the second question and not the first.
 */
function shiftAst(node: Ast, shift: Shift): Ast {
  switch (node.kind) {
    case 'ref': {
      const moved = shiftRef(node.ref, shift);
      return moved === node.ref ? node : { kind: 'ref', ref: moved };
    }
    case 'range': {
      const moved = shiftRange(node.range, shift);
      return moved === node.range ? node : { kind: 'range', range: moved };
    }
    case 'call': {
      let changed = false;
      const args = node.args.map(arg => {
        const moved = shiftAst(arg, shift);
        changed = changed || moved !== arg;
        return moved;
      });
      return changed ? { kind: 'call', name: node.name, args } : node;
    }
    case 'unary': {
      const operand = shiftAst(node.operand, shift);
      return operand === node.operand ? node : { kind: 'unary', op: node.op, operand };
    }
    case 'binary': {
      const left = shiftAst(node.left, shift);
      const right = shiftAst(node.right, shift);
      return left === node.left && right === node.right ? node : { kind: 'binary', op: node.op, left, right };
    }
    default:
      return node;
  }
}

/** Where an index on the shifted axis ends up, or -1 when it is gone. */
export function shiftIndex(index: number, shift: Shift): number {
  if (shift.by > 0) {
    return index >= shift.at ? index + shift.by : index;
  }
  const removed = -shift.by;
  if (index < shift.at) {
    return index;
  }
  // The cell itself was deleted. There is no index to point at, and
  // saying so is the only honest answer: `#REF!`.
  return index < shift.at + removed ? -1 : index - removed;
}

function indexOf(ref: CellRef, axis: Axis): number {
  return axis === 'row' ? ref.row : ref.column;
}

function withIndex(ref: CellRef, axis: Axis, index: number): CellRef {
  return axis === 'row' ? { ...ref, row: index } : { ...ref, column: index };
}

/**
 * A single reference.
 *
 * `$` is not consulted, and that is the whole point of this file.
 */
function shiftRef(ref: CellRef, shift: Shift): CellRef {
  const index = indexOf(ref, shift.axis);
  const moved = shiftIndex(index, shift);
  if (moved === index) {
    return ref;
  }
  // Off the sheet is out of bounds, which `printFormula` writes as
  // `#REF!` — the same answer a deleted cell gets, by the same route.
  return withIndex(ref, shift.axis, moved === -1 ? offSheet(shift.axis) : moved);
}

/**
 * A range, whose two corners do not move together.
 *
 * Three behaviours in one function, and they are what makes a range a
 * range rather than two references:
 *
 *   - **Inserting inside it grows it.** `SUM(A1:A10)` with a row
 *     inserted at row 5 becomes `SUM(A1:A11)`, because the rows it
 *     was adding are all still in it and there is now one more.
 *   - **Deleting inside it shrinks it.** The same sum over a deleted
 *     row is `SUM(A1:A9)` and not `#REF!`: the range still names a
 *     run of cells, and there are fewer of them.
 *   - **Deleting a corner clamps rather than breaks.** A range whose
 *     first row went away starts at the first row that survived. Only
 *     when *every* row it covered is gone does it become `#REF!`,
 *     which is the one case where there is nothing left to name.
 */
export function shiftRange(range: RangeRef, shift: Shift): RangeRef {
  const first = Math.min(indexOf(range.start, shift.axis), indexOf(range.end, shift.axis));
  const last = Math.max(indexOf(range.start, shift.axis), indexOf(range.end, shift.axis));

  let movedFirst: number;
  let movedLast: number;
  if (shift.by > 0) {
    // A row inserted *at* the end of a range extends it; one inserted
    // past the end does not. `>` rather than `>=` on the last index
    // is that distinction, and it is why the two corners are not the
    // same arithmetic.
    movedFirst = first >= shift.at ? first + shift.by : first;
    movedLast = last >= shift.at ? last + shift.by : last;
  } else {
    const removed = -shift.by;
    const lastRemoved = shift.at + removed - 1;
    if (first >= shift.at && last <= lastRemoved) {
      // Every cell it named is gone.
      return { start: offSheetRef(range.start, shift.axis), end: offSheetRef(range.end, shift.axis) };
    }
    movedFirst = first < shift.at ? first : Math.max(shift.at, first - removed);
    movedLast = last <= lastRemoved ? shift.at - 1 : last - removed;
  }
  if (movedFirst === first && movedLast === last) {
    return range;
  }
  const startIsFirst = indexOf(range.start, shift.axis) <= indexOf(range.end, shift.axis);
  return {
    start: withIndex(range.start, shift.axis, startIsFirst ? movedFirst : movedLast),
    end: withIndex(range.end, shift.axis, startIsFirst ? movedLast : movedFirst)
  };
}

/** An index past the end of the sheet, which prints as `#REF!`. */
function offSheet(axis: Axis): number {
  return axis === 'row' ? MAX_ROWS : MAX_COLUMNS;
}

function offSheetRef(ref: CellRef, axis: Axis): CellRef {
  return withIndex(ref, axis, offSheet(axis));
}
