import { columnIndex } from './A1';
import type { Ast } from './Ast';
import { DEFAULT_FORMAT, GENERAL, NO_BORDERS, PLAIN, type CellEdge, type CellFormat, type CellPaint, type DatePattern, type NumberFormat } from './Format';
import { isSheetFunction } from './Functions';
import type { MergeRect } from './Merges';
import { parseFormula } from './Parser';
import { rewriteFormula } from './Rewrite';
import { literalOf } from './Workbook';
import { child, children, parseXml, type XmlElement } from './Xml';
import { zipEntries, zipRead, type Inflate } from './Zip';

/**
 * An Excel workbook, read.
 *
 * **Import only**, and that is the decision Phase 16 made: a reader is
 * allowed to drop what it does not understand, and a writer is not.
 * So this keeps what this application can hold — cells, formulas,
 * number formats, bold and colour and fills and borders, alignment and
 * wrap, column widths, hidden rows, frozen panes, merges and defined
 * names — and leaves the rest: charts, pivot tables, conditional
 * formats, validations, comments, images, themes. Each of those is a
 * thing the file has and this sheet would show wrongly, which is worse
 * than not showing it.
 *
 * Formulas come across as formulas, so a workbook keeps working after
 * it is opened. Where one cannot — a function this sheet does not have,
 * a reference into another workbook, a syntax the parser does not know
 * — the value Excel last calculated is kept instead, and the count of
 * those is reported, because a formula that silently became a number
 * is a sheet that stops updating without saying so.
 *
 * Nothing here touches a platform: the zip's bytes and an inflate
 * function come in, plain records go out, and the application worker
 * turns them into a document.
 */

export interface XlsxCell {
  readonly row: number;
  readonly column: number;
  /** What would be typed: `=B4*C4`, `42`, `TRUE`, or the text itself. */
  readonly input: string;
  /** Index into the sheet's `formats`. */
  readonly style: number;
  /**
   * Text that would not read back as text if it were typed — `007`,
   * `TRUE`, `=1+2` — and has to be formatted Text to stay what Excel
   * said it was.
   */
  readonly asText: boolean;
}

export interface XlsxSheet {
  readonly name: string;
  readonly cells: readonly XlsxCell[];
  /**
   * Cells with a style and nothing in them: the fill across a header
   * row's blank cells, a border drawn round an empty box. Kept apart
   * from `cells` because they have no input, and dropped with it they
   * would take the look of the sheet with them.
   */
  readonly styled: readonly { readonly row: number; readonly column: number; readonly style: number }[];
  /** A format per cell style in the file, shared by the workbook's sheets. */
  readonly formats: readonly CellFormat[];
  /** Pixel widths of the columns the file gives a width, by index. */
  readonly columnWidths: ReadonlyMap<number, number>;
  readonly hiddenRows: readonly number[];
  readonly merges: readonly MergeRect[];
  readonly frozenRows: number;
  readonly frozenColumns: number;
}

export interface XlsxName {
  readonly name: string;
  /** The sheet its range is on, or null when the reference names none. */
  readonly sheet: string | null;
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

export interface XlsxBook {
  readonly sheets: readonly XlsxSheet[];
  readonly names: readonly XlsxName[];
  /** Formulas whose last value was kept, because they could not be. */
  readonly valuesKept: number;
  /** Rows and columns past the sheet's edge, left out. */
  readonly cut: { readonly rows: number; readonly columns: number };
  /** Defined names that are not a plain range, and so were not kept. */
  readonly namesSkipped: number;
}

export class XlsxError extends Error {}

/** How much of a file fits: this application's sheet size. */
export interface XlsxLimits {
  readonly rows: number;
  readonly columns: number;
}

/** Reads a workbook from the bytes of an `.xlsx`. */
export async function openXlsx(bytes: Uint8Array, inflate: Inflate, limits: XlsxLimits): Promise<XlsxBook> {
  let entries;
  try {
    entries = zipEntries(bytes);
  } catch {
    throw new XlsxError('It is not an Excel workbook; an .xlsx is a zip file, and this is not one.');
  }
  const decoder = new TextDecoder();
  const byName = new Map(entries.map(entry => [entry.name.replace(/^\//, ''), entry]));
  const parts = new Map<string, string>();
  for (const [name, entry] of byName) {
    // Only the parts a reader of cells needs: the workbook, its
    // relationships, the strings, the styles and the worksheets. A
    // file's images and pivot caches can be most of its bytes, and
    // inflating them to throw them away would be most of the time.
    if (/^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|styles\.xml|worksheets\/[^/]+\.xml)$/i.test(name)) {
      parts.set(name.toLowerCase(), decoder.decode(await zipRead(bytes, entry, inflate)));
    }
  }
  return readXlsx(path => parts.get(path.toLowerCase()) ?? null, limits);
}

/**
 * Reads a workbook from its parts, by path within the zip.
 *
 * Separate from `openXlsx` so a spec can hand it the XML a producer
 * would write — a shared formula, a frozen pane — without building a
 * zip around each one.
 */
export function readXlsx(read: (path: string) => string | null, limits: XlsxLimits): XlsxBook {
  const workbookText = read('xl/workbook.xml');
  if (workbookText === null) {
    throw new XlsxError('It is not an Excel workbook; there is no xl/workbook.xml in it.');
  }
  const workbook = parseXml(workbookText);
  const targets = relationships(read('xl/_rels/workbook.xml.rels'));
  const strings = sharedStrings(read('xl/sharedStrings.xml'));
  const formats = styles(read('xl/styles.xml'));

  const names: XlsxName[] = [];
  let namesSkipped = 0;
  for (const defined of children(child(workbook, 'definedNames'), 'definedName')) {
    const name = defined.attributes.name ?? '';
    // Excel's own names — the print area, a filter's range — start
    // `_xlnm.` and are the application's bookkeeping, not somebody's.
    if (name === '' || name.startsWith('_xlnm.') || defined.attributes.hidden === '1' || defined.attributes.hidden === 'true') {
      continue;
    }
    const range = rangeOf(defined.text.trim());
    if (range === null) {
      namesSkipped++;
      continue;
    }
    names.push({ name, ...range });
  }
  // A bare name parses as a call with no arguments, so a formula that
  // reads one is runnable when the name is one of these.
  const known = new Set(names.map(each => each.name.toUpperCase()));

  const cut = { rows: 0, columns: 0 };
  let valuesKept = 0;
  const sheets: XlsxSheet[] = [];
  for (const entry of children(child(workbook, 'sheets'), 'sheet')) {
    const target = targets.get(entry.attributes.id ?? '');
    const text = target === undefined ? null : read(target);
    if (text === null) {
      continue;
    }
    const sheet = worksheet(entry.attributes.name ?? `Sheet${sheets.length + 1}`, parseXml(text), strings, formats, limits, known);
    valuesKept += sheet.valuesKept;
    cut.rows = Math.max(cut.rows, sheet.cutRows);
    cut.columns = Math.max(cut.columns, sheet.cutColumns);
    sheets.push(sheet.sheet);
  }
  if (sheets.length === 0) {
    throw new XlsxError('It has no worksheets this can read.');
  }
  return { sheets, names, valuesKept, cut, namesSkipped };
}

// ---------------------------------------------------------------------------
// The workbook's parts
// ---------------------------------------------------------------------------

/** Relationship ids to part paths, resolved against `xl/`. */
function relationships(text: string | null): Map<string, string> {
  const map = new Map<string, string>();
  if (text === null) {
    return map;
  }
  for (const relationship of children(parseXml(text), 'Relationship')) {
    const target = relationship.attributes.Target ?? '';
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    map.set(relationship.attributes.Id ?? '', normalise(path));
  }
  return map;
}

function normalise(path: string): string {
  const out: string[] = [];
  for (const part of path.split('/')) {
    if (part === '..') {
      out.pop();
    } else if (part !== '.' && part !== '') {
      out.push(part);
    }
  }
  return out.join('/');
}

/**
 * The shared strings: the text of each `<si>`, runs joined.
 *
 * A rich string is runs of `<r><t>`; its formatting is dropped, since
 * a cell here is one format. A phonetic guide (`<rPh>`) is *not* part
 * of the text — it is how Japanese spreadsheets say how a name is
 * read — and joining it in would put the reading after every name.
 */
function sharedStrings(text: string | null): string[] {
  if (text === null) {
    return [];
  }
  return children(parseXml(text), 'si').map(stringOf);
}

function stringOf(item: XmlElement): string {
  const plain = child(item, 't');
  if (plain !== null) {
    return plain.text;
  }
  return children(item, 'r')
    .map(run => child(run, 't')?.text ?? '')
    .join('');
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

/**
 * A format per cell style (`cellXfs`), in file order, so a cell's `s`
 * indexes straight into it.
 *
 * The ids are followed and the `apply*` flags are not. The flags say
 * whether a style *overrides* its parent's, and Excel itself draws from
 * the ids either way — LibreOffice writes `applyFont="false"` on a
 * bold style, and the bold is real.
 */
function styles(text: string | null): CellFormat[] {
  if (text === null) {
    return [DEFAULT_FORMAT];
  }
  const sheet = parseXml(text);
  const codes = new Map<number, string>();
  for (const format of children(child(sheet, 'numFmts'), 'numFmt')) {
    codes.set(Number(format.attributes.numFmtId), format.attributes.formatCode ?? 'General');
  }
  const fonts = children(child(sheet, 'fonts'), 'font');
  const baseSize = Number(child(fonts[0], 'sz')?.attributes.val ?? 11);
  const fills = children(child(sheet, 'fills'), 'fill');
  const borders = children(child(sheet, 'borders'), 'border');
  const xfs = children(child(sheet, 'cellXfs'), 'xf');
  if (xfs.length === 0) {
    return [DEFAULT_FORMAT];
  }
  return xfs.map(xf => {
    const id = Number(xf.attributes.numFmtId ?? 0);
    const number = numberFormatOf(codes.get(id) ?? BUILT_IN[id] ?? 'General');
    const font = fonts[Number(xf.attributes.fontId ?? 0)];
    const fill = fills[Number(xf.attributes.fillId ?? 0)];
    const border = borders[Number(xf.attributes.borderId ?? 0)];
    const alignment = child(xf, 'alignment');
    const size = Number(child(font, 'sz')?.attributes.val ?? baseSize);
    const paint: CellPaint = {
      ...PLAIN,
      bold: flag(child(font, 'b')),
      italic: flag(child(font, 'i')),
      underline: child(font, 'u') !== null && child(font, 'u')!.attributes.val !== 'none',
      fontSize: size === baseSize || !Number.isFinite(size) ? 0 : size,
      color: colourOf(child(font, 'color')),
      fill: fillOf(fill),
      align: alignOf(alignment?.attributes.horizontal),
      wrap: alignment?.attributes.wrapText === '1' || alignment?.attributes.wrapText === 'true',
      borders:
        border === undefined
          ? NO_BORDERS
          : {
              top: edgeOf(child(border, 'top')),
              right: edgeOf(child(border, 'right')),
              bottom: edgeOf(child(border, 'bottom')),
              left: edgeOf(child(border, 'left'))
            }
    };
    return { number, paint };
  });
}

/** `<b/>`, `<b val="1"/>` and `<b val="true"/>` are bold; `<b val="0"/>` is not. */
function flag(element: XmlElement | null): boolean {
  if (element === null) {
    return false;
  }
  const value = element.attributes.val;
  return value === undefined || value === '1' || value === 'true';
}

/**
 * An `ARGB` colour as `#rrggbb`, or '' for one given by theme or index.
 *
 * Theme colours need the theme part and a tint calculation, and an
 * indexed palette is a table from 1995; both are dropped rather than
 * guessed, which draws the text in the default colour — the right
 * failure for a colour, because it is always legible.
 */
function colourOf(element: XmlElement | null): string {
  const rgb = element?.attributes.rgb;
  return rgb !== undefined && /^[0-9a-f]{8}$/i.test(rgb) ? `#${rgb.slice(2).toLowerCase()}` : '';
}

function fillOf(fill: XmlElement | undefined): string {
  const pattern = child(fill, 'patternFill');
  return pattern?.attributes.patternType === 'solid' ? colourOf(child(pattern, 'fgColor')) : '';
}

function alignOf(horizontal: string | undefined): CellPaint['align'] {
  switch (horizontal) {
    case 'left':
      return 'start';
    case 'center':
    case 'centerContinuous':
      return 'center';
    case 'right':
      return 'end';
    default:
      return 'auto';
  }
}

const EDGE_WIDTHS: Readonly<Record<string, number>> = {
  hair: 1,
  thin: 1,
  dotted: 1,
  dashed: 1,
  medium: 2,
  mediumDashed: 2,
  thick: 3,
  double: 3
};

function edgeOf(edge: XmlElement | null): CellEdge {
  const width = EDGE_WIDTHS[edge?.attributes.style ?? ''] ?? 0;
  return width === 0 ? { width: 0, color: '' } : { width, color: colourOf(child(edge, 'color')) };
}

/**
 * The formats Excel numbers without writing out, by id.
 *
 * 14 to 22 are *locale* formats — Excel shows id 14 as the reader's own
 * short date — and are written here as the US ones, which is what a
 * file saved in the US shows and the least surprising guess anywhere.
 */
const BUILT_IN: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0',
  6: '$#,##0',
  7: '$#,##0.00',
  8: '$#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0',
  38: '#,##0',
  39: '#,##0.00',
  40: '#,##0.00',
  44: '$#,##0.00',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '0.0E+0',
  49: '@'
};

/**
 * What a number format code comes to, in the formats this sheet has.
 *
 * A code is a small language — sections, colours, conditions, literal
 * text, padding — and this sheet has a handful of named formats, so
 * this reads the code for the one it is closest to: the first section,
 * with its quoted text, escapes and bracketed parts set aside, and
 * then dates, times, percentages, scientific, currency and plain
 * numbers, in that order. What it cannot place is General, which
 * shows the value as it is.
 */
export function numberFormatOf(code: string): NumberFormat {
  if (code.trim() === '@') {
    return { kind: 'text' };
  }
  const first = code.split(';')[0];
  // A currency symbol is written bracketed, `[$€-407]`, or quoted,
  // `"$"#,##0.00` — the second is how Excel writes its own currency
  // formats, so stripping quoted text before looking would lose it.
  const currency = /\[\$([^\-\]]*)[^\]]*\]/.exec(first)?.[1] ?? /"([$€£¥])"/.exec(first)?.[1];
  const bare = first
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_.|\*./g, '');
  if (/general/i.test(bare) || bare.trim() === '') {
    return currency === undefined ? GENERAL : { kind: 'currency', places: 2, symbol: currency };
  }
  const hasDate = /[yd]/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare));
  const hasTime = /[hs]/i.test(bare) || /am\/pm/i.test(bare);
  if (hasDate || hasTime) {
    const time = /s/i.test(bare.replace(/am\/pm/i, '')) ? ('hms' as const) : ('hm' as const);
    if (!hasDate) {
      return { kind: 'time', pattern: time };
    }
    const date = dateOrder(bare);
    return hasTime ? { kind: 'datetime', date, time } : { kind: 'date', pattern: date };
  }
  const places = /\.([0#?]+)/.exec(bare)?.[1].length ?? 0;
  if (bare.includes('%')) {
    return { kind: 'percent', places };
  }
  if (/E[+-]/i.test(bare)) {
    return { kind: 'scientific', places };
  }
  const symbol = currency ?? /[$€£¥]/.exec(bare)?.[0];
  if (symbol !== undefined && symbol !== '') {
    return { kind: 'currency', places, symbol };
  }
  if (/[0#]/.test(bare)) {
    return { kind: 'number', places, thousands: bare.includes(',') };
  }
  return GENERAL;
}

/** Which of the three orders a date code writes its day, month and year in. */
function dateOrder(code: string): DatePattern {
  const lower = code.toLowerCase();
  const y = lower.indexOf('y');
  const d = lower.indexOf('d');
  const m = lower.indexOf('m');
  if (y !== -1 && (d === -1 || y < d) && (m === -1 || y < m)) {
    return 'ymd';
  }
  if (d !== -1 && (m === -1 || d < m)) {
    return 'dmy';
  }
  return 'mdy';
}

// ---------------------------------------------------------------------------
// A worksheet
// ---------------------------------------------------------------------------

interface ReadSheet {
  readonly sheet: XlsxSheet;
  readonly valuesKept: number;
  readonly cutRows: number;
  readonly cutColumns: number;
}

/** A cell address, `C12`, as zero-based row and column. */
function addressOf(ref: string): { row: number; column: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(ref);
  if (match === null) {
    return null;
  }
  const column = columnIndex(match[1].toUpperCase());
  return column === null ? null : { row: Number(match[2]) - 1, column };
}

/**
 * A defined name's range, `Rates!$B$1` or `'Q3 plan'!$A$1:$C$9`, or
 * null for anything that is not one plain range.
 */
function rangeOf(text: string): Omit<XlsxName, 'name'> | null {
  const match = /^(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?([$A-Z0-9]+)(?::([$A-Z0-9]+))?$/i.exec(text);
  if (match === null) {
    return null;
  }
  const sheet = match[1]?.replace(/''/g, "'") ?? match[2] ?? null;
  const start = addressOf(match[3]);
  const end = match[4] === undefined ? start : addressOf(match[4]);
  if (start === null || end === null) {
    return null;
  }
  return {
    sheet,
    firstRow: Math.min(start.row, end.row),
    firstColumn: Math.min(start.column, end.column),
    lastRow: Math.max(start.row, end.row),
    lastColumn: Math.max(start.column, end.column)
  };
}

/**
 * Excel's column width, in characters of the default font's digits,
 * as pixels — the arithmetic in the format's own documentation,
 * rounded, for the seven-pixel digit of Calibri 11 that the width is
 * measured against.
 */
function pixelsOf(width: number): number {
  return Math.round(Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7));
}

function worksheet(
  name: string,
  root: XmlElement,
  strings: readonly string[],
  formats: readonly CellFormat[],
  limits: XlsxLimits,
  known: ReadonlySet<string>
): ReadSheet {
  const cells: XlsxCell[] = [];
  const styled: { row: number; column: number; style: number }[] = [];
  const hiddenRows: number[] = [];
  const columnWidths = new Map<number, number>();
  /** A shared formula's first cell, by its `si`: the formula and where it stood. */
  const shared = new Map<string, { input: string; row: number; column: number }>();
  let valuesKept = 0;
  let cutRows = 0;
  let cutColumns = 0;

  for (const column of children(child(root, 'cols'), 'col')) {
    const min = Number(column.attributes.min) - 1;
    const max = Math.min(Number(column.attributes.max) - 1, limits.columns - 1);
    const hidden = column.attributes.hidden === '1' || column.attributes.hidden === 'true';
    const width = Number(column.attributes.width);
    if (!hidden && !Number.isFinite(width)) {
      continue;
    }
    for (let at = min; at <= max; at++) {
      // A hidden column is one of width zero, which is how this sheet
      // hides one; see `hideColumns`.
      columnWidths.set(at, hidden ? 0 : pixelsOf(width));
    }
  }

  const sheetData = child(root, 'sheetData');
  let nextRow = 0;
  for (const rowElement of children(sheetData, 'row')) {
    const row = rowElement.attributes.r === undefined ? nextRow : Number(rowElement.attributes.r) - 1;
    nextRow = row + 1;
    if (row >= limits.rows) {
      cutRows = Math.max(cutRows, row + 1 - limits.rows);
      continue;
    }
    if (rowElement.attributes.hidden === '1' || rowElement.attributes.hidden === 'true') {
      hiddenRows.push(row);
    }
    let nextColumn = 0;
    for (const cell of children(rowElement, 'c')) {
      const at = cell.attributes.r === undefined ? { row, column: nextColumn } : addressOf(cell.attributes.r);
      if (at === null) {
        continue;
      }
      nextColumn = at.column + 1;
      if (at.column >= limits.columns) {
        cutColumns = Math.max(cutColumns, at.column + 1 - limits.columns);
        continue;
      }
      const style = Math.min(Number(cell.attributes.s ?? 0), formats.length - 1);
      const read = cellInput(cell, at.row, at.column, strings, shared, known);
      if (read === null) {
        if (style > 0) {
          styled.push({ row: at.row, column: at.column, style });
        }
        continue;
      }
      if (read.kept) {
        valuesKept++;
      }
      cells.push({ row: at.row, column: at.column, input: read.input, style: Math.max(0, style), asText: read.asText });
    }
  }

  const merges: MergeRect[] = [];
  for (const merge of children(child(root, 'mergeCells'), 'mergeCell')) {
    const range = rangeOf(merge.attributes.ref ?? '');
    if (range !== null && range.lastRow < limits.rows && range.lastColumn < limits.columns) {
      merges.push({
        firstRow: range.firstRow,
        lastRow: range.lastRow,
        firstColumn: range.firstColumn,
        lastColumn: range.lastColumn
      });
    }
  }

  // A frozen pane is a `pane` with a frozen state; a split one that is
  // not frozen is a view preference this sheet does not have.
  let frozenRows = 0;
  let frozenColumns = 0;
  const pane = child(child(child(root, 'sheetViews'), 'sheetView'), 'pane');
  if (pane !== null && (pane.attributes.state === 'frozen' || pane.attributes.state === 'frozenSplit')) {
    frozenRows = Math.min(Math.max(0, Math.round(Number(pane.attributes.ySplit ?? 0))), limits.rows);
    frozenColumns = Math.min(Math.max(0, Math.round(Number(pane.attributes.xSplit ?? 0))), limits.columns);
  }

  return {
    sheet: { name, cells, styled, formats, columnWidths, hiddenRows, merges, frozenRows, frozenColumns },
    valuesKept,
    cutRows,
    cutColumns
  };
}

/**
 * What one `<c>` becomes: its input, whether it has to be kept as
 * text, and whether a formula had to be traded for its value.
 */
function cellInput(
  cell: XmlElement,
  row: number,
  column: number,
  strings: readonly string[],
  shared: Map<string, { input: string; row: number; column: number }>,
  known: ReadonlySet<string>
): { input: string; asText: boolean; kept: boolean } | null {
  const type = cell.attributes.t ?? 'n';
  const raw = child(cell, 'v')?.text ?? null;
  const value = valueOf(type, raw, cell, strings);

  const formula = child(cell, 'f');
  if (formula !== null) {
    const written = formulaOf(formula, row, column, shared);
    if (written !== null && canRun(written, known)) {
      return { input: written, asText: false, kept: false };
    }
    if (value === null) {
      return null;
    }
    return { ...literal(value), kept: true };
  }
  return value === null ? null : { ...literal(value), kept: false };
}

type Value = { kind: 'text'; text: string } | { kind: 'other'; text: string };

function valueOf(type: string, raw: string | null, cell: XmlElement, strings: readonly string[]): Value | null {
  switch (type) {
    case 's': {
      const text = raw === null ? undefined : strings[Number(raw)];
      return text === undefined ? null : { kind: 'text', text };
    }
    case 'inlineStr': {
      const item = child(cell, 'is');
      return item === null ? null : { kind: 'text', text: stringOf(item) };
    }
    case 'str':
      return raw === null ? null : { kind: 'text', text: raw };
    case 'b':
      return raw === null ? null : { kind: 'other', text: raw === '1' ? 'TRUE' : 'FALSE' };
    case 'e':
      // An error is kept as its code, as text: this sheet has no way to
      // type an error value, and `#N/A` in a cell reads the same.
      return raw === null ? null : { kind: 'text', text: raw };
    default:
      return raw === null || raw === '' ? null : { kind: 'other', text: raw };
  }
}

/** A value as an input, marked text when typing it would not give text back. */
function literal(value: Value): { input: string; asText: boolean } {
  if (value.kind === 'other') {
    return { input: value.text, asText: false };
  }
  const text = value.text;
  return { input: text, asText: text.startsWith('=') || typeof literalOf(text) !== 'string' };
}

/**
 * A formula as this sheet would write it, or null.
 *
 * Excel writes a formula once for a block of cells that share it —
 * `<f t="shared" ref="D4:D7" si="0">B4*C4</f>` on the first, and just
 * `<f t="shared" si="0"/>` on the rest — so the rest are the first one
 * moved, exactly as a fill would move it. The prefixes Excel puts on
 * functions newer than 2007 (`_xlfn.IFS`) are dropped: they mark a
 * function an old Excel would not know, and this sheet either knows it
 * by its name or does not know it at all.
 */
function formulaOf(
  formula: XmlElement,
  row: number,
  column: number,
  shared: Map<string, { input: string; row: number; column: number }>
): string | null {
  const text = formula.text.trim();
  const index = formula.attributes.si;
  if (formula.attributes.t === 'shared' && index !== undefined) {
    if (text !== '') {
      const input = `=${clean(text)}`;
      shared.set(index, { input, row, column });
      return input;
    }
    const first = shared.get(index);
    return first === undefined ? null : rewriteFormula(first.input, row - first.row, column - first.column);
  }
  return text === '' ? null : `=${clean(text)}`;
}

function clean(text: string): string {
  return text.replace(/_xl(?:fn|ws|pm)\./g, '');
}

/** Whether this sheet can parse a formula and knows every function and name it calls. */
function canRun(input: string, known: ReadonlySet<string>): boolean {
  let tree: Ast;
  try {
    tree = parseFormula(input.slice(1));
  } catch {
    return false;
  }
  const knowsAll = (node: Ast): boolean => {
    switch (node.kind) {
      case 'call': {
        const name = node.name.toUpperCase();
        const found = isSheetFunction(name) || (node.args.length === 0 && known.has(name));
        return found && node.args.every(knowsAll);
      }
      case 'unary':
        return knowsAll(node.operand);
      case 'binary':
        return knowsAll(node.left) && knowsAll(node.right);
      default:
        return true;
    }
  };
  return knowsAll(tree);
}
