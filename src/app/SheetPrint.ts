import type { CellEdge, CellPaint } from '../sheet/Format';
import { CELL_FONT_SIZE, ROW_HEIGHT } from './dimensions';
import type { SheetDocument } from './SheetDocument';

/**
 * The sheet in view as a page a browser can print — Phase 37.
 *
 * A canvas prints as a picture of the screen, cut at the window's edge,
 * which is not what anybody means by printing a spreadsheet. What they
 * mean is the table: every row and column with something in it, as it
 * is shown, on as many pages as it takes. That is an HTML table, and
 * the browser already knows how to break one across pages, repeat
 * nothing it should not and save it as a PDF.
 *
 * Built on the application worker, which is where the values, the
 * formats and the merges are, and handed to `public/print.html` over a
 * `BroadcastChannel` (see `AppWorker.ts`). Everything that came from a
 * cell is escaped; a colour that is not a plain colour is dropped.
 */

export interface PrintJob {
  /** Bumped per job, so the page draws each once. */
  readonly id: number;
  /** The page's title, which is also the PDF's suggested name. */
  readonly title: string;
  /** The table, as markup for the page's body. */
  readonly html: string;
  /** Whether it was asked for as a PDF, which the page says how to make. */
  readonly pdf: boolean;
}

/** The last row and column worth printing: anything typed, spilled into, or merged over. */
export function printedExtent(document: SheetDocument): { rows: number; columns: number } {
  let lastRow = -1;
  let lastColumn = -1;
  const sheet = document.shown;
  for (const { row, column } of document.sheet.entries()) {
    const spill = sheet.spillOf(row, column);
    lastRow = Math.max(lastRow, row + (spill === null ? 0 : spill.rows - 1));
    lastColumn = Math.max(lastColumn, column + (spill === null ? 0 : spill.columns - 1));
  }
  for (const rect of document.merges.all) {
    lastRow = Math.max(lastRow, rect.lastRow);
    lastColumn = Math.max(lastColumn, rect.lastColumn);
  }
  return { rows: lastRow + 1, columns: lastColumn + 1 };
}

/**
 * The table for the sheet in view: hidden and filtered rows left out,
 * hidden columns left out, column widths kept, merges spanned.
 * `paintOf` is the paint as drawn, a rule's over the cell's own; the
 * cell's own format is used without it.
 */
export function printTable(document: SheetDocument, paintOf?: (row: number, column: number) => CellPaint): string {
  const { rows, columns } = printedExtent(document);
  if (rows === 0) {
    return '<p class="empty">This sheet is empty.</p>';
  }
  const widths = document.columnWidths;
  const shownColumns = Array.from({ length: columns }, (_, column) => column).filter(column => (widths[column] ?? 1) > 0);
  const hidden = (row: number) => document.hiddenRows.has(row) || document.filteredRows.has(row);
  const parts: string[] = ['<table class="sheet"><colgroup>'];
  for (const column of shownColumns) {
    parts.push(`<col style="width:${Math.round(widths[column] ?? 104)}px">`);
  }
  parts.push('</colgroup><tbody>');
  for (let row = 0; row < rows; row++) {
    if (hidden(row)) {
      continue;
    }
    const height = document.rowHeights.get(row) ?? document.fittedRows.get(row) ?? ROW_HEIGHT;
    parts.push(height === ROW_HEIGHT ? '<tr>' : `<tr style="height:${Math.round(height)}px">`);
    for (const column of shownColumns) {
      if (document.merges.isHidden(row, column)) {
        continue;
      }
      const merge = document.merges.at(row, column);
      let span = '';
      if (merge !== null && merge.firstRow === row && merge.firstColumn === column) {
        let rowSpan = 0;
        for (let each = merge.firstRow; each <= merge.lastRow; each++) {
          rowSpan += hidden(each) ? 0 : 1;
        }
        const columnSpan = shownColumns.filter(each => each >= merge.firstColumn && each <= merge.lastColumn).length;
        span = `${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ''}${columnSpan > 1 ? ` colspan="${columnSpan}"` : ''}`;
      }
      const value = document.shown.value(row, column);
      const text = document.display(row, column);
      const paint = paintOf?.(row, column) ?? document.formats.formatAt(row, column).paint;
      const style = cellStyle(paint, typeof value === 'number' ? 'right' : typeof value === 'boolean' ? 'center' : 'left');
      parts.push(`<td${span}${style === '' ? '' : ` style="${style}"`}>${escape(text)}</td>`);
    }
    parts.push('</tr>');
  }
  parts.push('</tbody></table>');
  return parts.join('');
}

/** The inline style for one cell's paint; empty for a plain cell. */
export function cellStyle(paint: CellPaint, natural: 'left' | 'right' | 'center'): string {
  const rules: string[] = [];
  if (paint.bold) {
    rules.push('font-weight:bold');
  }
  if (paint.italic) {
    rules.push('font-style:italic');
  }
  if (paint.underline) {
    rules.push('text-decoration:underline');
  }
  if (paint.fontSize !== 0 && paint.fontSize !== CELL_FONT_SIZE) {
    rules.push(`font-size:${paint.fontSize}pt`);
  }
  const family = (paint as CellPaint & { fontFamily?: string }).fontFamily;
  if (family !== undefined && family !== '' && /^[\w\s,'-]+$/.test(family)) {
    rules.push(`font-family:${family}`);
  }
  if (isColour(paint.color)) {
    rules.push(`color:${paint.color}`);
  }
  if (isColour(paint.fill)) {
    rules.push(`background:${paint.fill}`);
  }
  const align = paint.align === 'auto' ? natural : paint.align === 'start' ? 'left' : paint.align === 'end' ? 'right' : 'center';
  if (align !== 'left') {
    rules.push(`text-align:${align}`);
  }
  if (paint.wrap) {
    rules.push('white-space:normal');
  }
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const edge: CellEdge = paint.borders[side];
    if (edge.width > 0) {
      rules.push(`border-${side}:${edge.width}px solid ${isColour(edge.color) ? edge.color : '#000'}`);
    }
  }
  return rules.join(';');
}

/** A colour as this application writes them: a hex, or a plain CSS colour name. */
function isColour(colour: string): boolean {
  return /^#[0-9a-f]{3,8}$/i.test(colour) || /^[a-z]{3,20}$/i.test(colour);
}

export function escape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
