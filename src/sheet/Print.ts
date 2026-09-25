import { formatRange, formatRef, inBounds, type RangeRef } from './A1';
import type { Ast } from './Ast';
import { formatNumber } from './Values';

/**
 * A parsed formula back as text.
 *
 * Shared by the two things that move a formula — `Rewrite` for a fill
 * and `Shift` for an insert — because two printers are two chances to
 * disagree about what a tree means, and the place they would disagree
 * is the place a formula silently changes.
 *
 * **Fully parenthesised rather than minimally.** A writer that dropped
 * brackets would have to know the precedence table as exactly as the
 * parser does, and the one place the two could differ is the one that
 * matters. Extra brackets are ugly and cannot be wrong. (A
 * precedence-aware printer with a parse-back round-trip spec is the
 * way to lose them, and it belongs with the formula editor in Phase
 * 12 rather than here.)
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
      return `${node.name}(${node.args.map(printFormula).join(',')})`;
    case 'unary':
      return `${node.op}${printFormula(node.operand)}`;
    case 'binary':
      return `(${printFormula(node.left)}${node.op}${printFormula(node.right)})`;
  }
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
