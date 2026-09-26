import { columnName, quoteSheetName } from './A1';
import type { Validation } from './Validation';
import { dxfOf, writeConditionals, writeValidations } from './XlsxRules';
import type { ConditionalPaint, ConditionalRule } from './Conditional';
import type { CellEdge, CellFormat, NumberFormat } from './Format';
import { withIntersections } from './Legacy';
import type { MergeRect } from './Merges';
import { tokenize, type Token } from './Tokenizer';
import { isError, type CellValue } from './Values';
import { zipWrite, type Deflate } from './Zip';

/**
 * A workbook written as an `.xlsx`: the reader in `Xlsx.ts`, run the
 * other way.
 *
 * What goes in is what the reader brings out — values and formulas,
 * number formats, fonts, fills, borders, alignment and wrap, column
 * widths, row heights, hidden rows and columns, merges, frozen panes,
 * defined names and iterative calculation — so a workbook that came in
 * from Excel goes back out with what it came in with. Conditional
 * formats, validations and charts are left out, as the reader leaves
 * them out, and the caller says so.
 *
 * Plain records in, bytes out, and nothing about the document: the
 * application worker builds the records, as it builds a snapshot.
 *
 * **Formulas are written as Excel reads them.** A function newer than
 * the file format carries Excel's `_xlfn.` prefix. A formula that
 * spills is an array formula over the cells it fills; so is one that
 * does arithmetic over a range without spilling (`=SUM(A1:A3*B1:B3)`),
 * because written plainly an Excel from before dynamic arrays would
 * take one value from each range. `@` is Excel's default in a plain
 * formula and is left out there; in an array formula it is written as
 * `_xlfn.SINGLE`, which is how Excel 365 writes it.
 */

export interface XlsxOutCell {
  readonly row: number;
  readonly column: number;
  /** What was typed; empty for a cell that is only formatted, or only spilled into. */
  readonly input: string;
  readonly value: CellValue;
  /** Index into `XlsxOut.formats`. */
  readonly style: number;
  /** How far the formula's array spills, for a formula that does. */
  readonly spill?: { readonly rows: number; readonly columns: number };
}

export interface XlsxOutRow {
  /** In points; absent for the default height. */
  readonly height?: number;
  /** A height somebody set, as against one fitted to the contents. */
  readonly custom: boolean;
  readonly hidden: boolean;
}

export interface XlsxOutSheet {
  readonly name: string;
  readonly cells: readonly XlsxOutCell[];
  /** Excel's width in characters, by column; 0 is hidden, and absent the default. */
  readonly columnWidths: ReadonlyMap<number, number>;
  readonly rows: ReadonlyMap<number, XlsxOutRow>;
  readonly merges: readonly MergeRect[];
  readonly frozenRows: number;
  readonly frozenColumns: number;
  /** The notes on the sheet's cells, written as Excel's comments. */
  readonly notes?: readonly { readonly row: number; readonly column: number; readonly text: string }[];
  /** What the cells may hold; see `XlsxRules.writeValidations`. */
  readonly validations?: readonly Validation[];
  /** The formats that think; see `XlsxRules.writeConditionals`. */
  readonly conditional?: readonly ConditionalRule[];
}

export interface XlsxOutName {
  readonly name: string;
  readonly sheet: string;
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

export interface XlsxOut {
  readonly sheets: readonly XlsxOutSheet[];
  /** Every format a cell uses, 0 the default. */
  readonly formats: readonly CellFormat[];
  readonly names: readonly XlsxOutName[];
  readonly active: number;
  readonly iteration: { readonly count: number; readonly delta: number } | null;
}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** The workbook as the bytes of an `.xlsx`. */
export async function writeXlsx(book: XlsxOut, deflate?: Deflate): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  return zipWrite(
    xlsxParts(book).map(part => ({ name: part.name, bytes: encoder.encode(part.text) })),
    deflate
  );
}

/** The parts of the package, as text: separate from the zip so a spec can read them. */
export function xlsxParts(book: XlsxOut): { name: string; text: string }[] {
  const sheets = book.sheets.length === 0 ? [emptySheet()] : book.sheets;
  // Every conditional format's paint, once, in the order first met:
  // the workbook's `dxfs`, which each rule points into by index.
  const dxfs: string[] = [];
  const dxfIdOf = (paint: ConditionalPaint): number => {
    const written = dxfOf(paint);
    const at = dxfs.indexOf(written);
    if (at !== -1) {
      return at;
    }
    dxfs.push(written);
    return dxfs.length - 1;
  };
  const priority = { next: 1 };
  const ruled = (book.sheets.length === 0 ? [] : book.sheets).map(sheet =>
    writeConditionals(sheet.conditional ?? [], dxfIdOf, input => excelFormula(input, false), priority)
  );
  const ranged = new Set(
    book.names
      .filter(name => name.firstRow !== name.lastRow || name.firstColumn !== name.lastColumn)
      .map(name => name.name.toUpperCase())
  );
  const parts: { name: string; text: string }[] = [
    { name: '[Content_Types].xml', text: contentTypes(sheets.length, sheets.map(sheet => (sheet.notes?.length ?? 0) > 0)) },
    {
      name: '_rels/.rels',
      text: xml(
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
      )
    },
    { name: 'xl/workbook.xml', text: workbookOf(book, sheets) },
    {
      name: 'xl/_rels/workbook.xml.rels',
      text: xml(
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">${sheets
          .map(
            (_, at) =>
              `<Relationship Id="rId${at + 1}" Type="${RELATIONSHIPS}/worksheet" Target="worksheets/sheet${at + 1}.xml"/>`
          )
          .join('')}<Relationship Id="rId${sheets.length + 1}" Type="${RELATIONSHIPS}/styles" Target="styles.xml"/></Relationships>`
      )
    },
    { name: 'xl/styles.xml', text: stylesOf(book.formats, dxfs) }
  ];
  sheets.forEach((sheet, at) => {
    const noted = (sheet.notes?.length ?? 0) > 0;
    parts.push({ name: `xl/worksheets/sheet${at + 1}.xml`, text: worksheetOf(sheet, ranged, noted, ruled[at] ?? '') });
    if (noted) {
      parts.push(...notesOf(sheet.notes ?? [], at + 1));
    }
  });
  return parts;
}

function emptySheet(): XlsxOutSheet {
  return { name: 'Sheet1', cells: [], columnWidths: new Map(), rows: new Map(), merges: [], frozenRows: 0, frozenColumns: 0 };
}

function xml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

/** Text as XML may hold it: escaped, and without the control characters it may not. */
function escape(text: string): string {
  return text
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function contentTypes(sheets: number, noted: readonly boolean[] = []): string {
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`
  ];
  for (let at = 1; at <= sheets; at++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${at}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
    if (noted[at - 1] === true) {
      overrides.push(
        `<Override PartName="/xl/comments${at}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`
      );
    }
  }
  const vml = noted.some(Boolean)
    ? `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`
    : '';
  return xml(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${vml}${overrides.join('')}</Types>`
  );
}

function workbookOf(book: XlsxOut, sheets: readonly XlsxOutSheet[]): string {
  const names = book.names
    .map(name => {
      const start = `$${columnName(name.firstColumn)}$${name.firstRow + 1}`;
      const end = `$${columnName(name.lastColumn)}$${name.lastRow + 1}`;
      const ref = `${quoteSheetName(name.sheet)}!${start}${start === end ? '' : `:${end}`}`;
      return `<definedName name="${escape(name.name)}">${escape(ref)}</definedName>`;
    })
    .join('');
  const iteration =
    book.iteration === null
      ? ''
      : ` iterate="1" iterateCount="${book.iteration.count}" iterateDelta="${book.iteration.delta}"`;
  return xml(
    `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}"><workbookPr/><bookViews><workbookView activeTab="${Math.min(
      Math.max(0, book.active),
      sheets.length - 1
    )}"/></bookViews><sheets>${sheets
      .map((sheet, at) => `<sheet name="${escape(sheet.name)}" sheetId="${at + 1}" r:id="rId${at + 1}"/>`)
      .join('')}</sheets>${names === '' ? '' : `<definedNames>${names}</definedNames>`}` +
      // Recalculated on open, because the values written are this
      // sheet's and Excel's arithmetic should have the last word.
      `<calcPr calcId="191029" fullCalcOnLoad="1"${iteration}/></workbook>`
  );
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

/** A number format as the code Excel writes, which `numberFormatOf` reads back as the same format. */
export function codeOf(format: NumberFormat): string {
  const decimals = (places: number) => (places > 0 ? `.${'0'.repeat(places)}` : '');
  switch (format.kind) {
    case 'general':
      return 'General';
    case 'number':
      return `${format.thousands ? '#,##0' : '0'}${decimals(format.places)}`;
    case 'currency': {
      const symbol = /^[$€£¥]$/.test(format.symbol) ? `"${format.symbol}"` : `[$${format.symbol}]`;
      return `${symbol}#,##0${decimals(format.places)}`;
    }
    case 'percent':
      return `0${decimals(format.places)}%`;
    case 'scientific':
      return `0${decimals(format.places)}E+00`;
    case 'date':
      return DATE_CODES[format.pattern];
    case 'time':
      return TIME_CODES[format.pattern];
    case 'datetime':
      return `${DATE_CODES[format.date]} ${TIME_CODES[format.time]}`;
    case 'text':
      return '@';
  }
}

const DATE_CODES = { ymd: 'yyyy-mm-dd', dmy: 'd mmm yyyy', mdy: 'mmm d, yyyy' } as const;
const TIME_CODES = { hm: 'hh:mm', hms: 'hh:mm:ss' } as const;

const EDGE_STYLES = ['', 'thin', 'medium', 'thick'];

/** `#rrggbb` as Excel's ARGB, or null for anything that is not a colour written out. */
function argb(colour: string): string | null {
  return /^#[0-9a-f]{6}$/i.test(colour) ? `FF${colour.slice(1).toUpperCase()}` : null;
}

function stylesOf(formats: readonly CellFormat[], dxfs: readonly string[] = []): string {
  const numFmts: string[] = [];
  const numIds = new Map<string, number>([
    ['General', 0],
    ['@', 49]
  ]);
  const fonts: string[] = [];
  const fontIds = new Map<string, number>();
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const fillIds = new Map<string, number>([['', 0]]);
  const borders: string[] = [];
  const borderIds = new Map<string, number>();
  const intern = (list: string[], ids: Map<string, number>, key: string, make: () => string): number => {
    let id = ids.get(key);
    if (id === undefined) {
      id = list.length;
      list.push(make());
      ids.set(key, id);
    }
    return id;
  };
  // The defaults first, so every index 0 is the plain one Excel expects.
  intern(fonts, fontIds, fontKey(formats[0]), () => fontOf(formats[0]));
  intern(borders, borderIds, borderKey(formats[0]), () => borderOf(formats[0]));

  const xfs = formats.map(format => {
    const code = codeOf(format.number);
    let numId = numIds.get(code);
    if (numId === undefined) {
      numId = 164 + numFmts.length;
      numIds.set(code, numId);
      numFmts.push(`<numFmt numFmtId="${numId}" formatCode="${escape(code)}"/>`);
    }
    const fontId = intern(fonts, fontIds, fontKey(format), () => fontOf(format));
    const fill = argb(format.paint.fill);
    const fillId =
      fill === null
        ? 0
        : intern(fills, fillIds, fill, () => `<fill><patternFill patternType="solid"><fgColor rgb="${fill}"/><bgColor indexed="64"/></patternFill></fill>`);
    const borderId = intern(borders, borderIds, borderKey(format), () => borderOf(format));
    const horizontal = { auto: '', start: 'left', center: 'center', end: 'right' }[format.paint.align];
    const alignment =
      horizontal === '' && !format.paint.wrap
        ? ''
        : `<alignment${horizontal === '' ? '' : ` horizontal="${horizontal}"`}${format.paint.wrap ? ' wrapText="1"' : ''}/>`;
    return (
      `<xf numFmtId="${numId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
      `${numId === 0 ? '' : ' applyNumberFormat="1"'}${fontId === 0 ? '' : ' applyFont="1"'}` +
      `${fillId === 0 ? '' : ' applyFill="1"'}${borderId === 0 ? '' : ' applyBorder="1"'}` +
      `${alignment === '' ? '/>' : ` applyAlignment="1">${alignment}</xf>`}`
    );
  });

  return xml(
    `<styleSheet xmlns="${MAIN}">` +
      (numFmts.length === 0 ? '' : `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>`) +
      `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
      `<fills count="${fills.length}">${fills.join('')}</fills>` +
      `<borders count="${borders.length}">${borders.join('')}</borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      (dxfs.length === 0 ? '' : `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`) +
      `</styleSheet>`
  );
}

function fontKey(format: CellFormat | undefined): string {
  const paint = format?.paint;
  return JSON.stringify([paint?.bold, paint?.italic, paint?.underline, paint?.fontSize, paint?.color]);
}

/** A font, children in the order the schema insists on. */
function fontOf(format: CellFormat | undefined): string {
  const paint = format?.paint;
  const colour = argb(paint?.color ?? '');
  return (
    '<font>' +
    (paint?.bold === true ? '<b/>' : '') +
    (paint?.italic === true ? '<i/>' : '') +
    (paint?.underline === true ? '<u/>' : '') +
    // The reader keeps a size the font has as it is; the default is 11.
    `<sz val="${paint !== undefined && paint.fontSize > 0 ? paint.fontSize : 11}"/>` +
    (colour === null ? '' : `<color rgb="${colour}"/>`) +
    '<name val="Calibri"/><family val="2"/></font>'
  );
}

function borderKey(format: CellFormat | undefined): string {
  return JSON.stringify(format?.paint.borders ?? null);
}

/** The four edges, in the schema's order: left, right, top, bottom, diagonal. */
function borderOf(format: CellFormat | undefined): string {
  const edges = format?.paint.borders;
  const edge = (name: string, value: CellEdge | undefined): string => {
    const style = EDGE_STYLES[Math.min(value?.width ?? 0, 3)] ?? 'thick';
    if (value === undefined || style === '') {
      return `<${name}/>`;
    }
    const colour = argb(value.color);
    return `<${name} style="${style}">${colour === null ? '<color auto="1"/>' : `<color rgb="${colour}"/>`}</${name}>`;
  };
  return `<border>${edge('left', edges?.left)}${edge('right', edges?.right)}${edge('top', edges?.top)}${edge('bottom', edges?.bottom)}<diagonal/></border>`;
}

// ---------------------------------------------------------------------------
// Worksheets
// ---------------------------------------------------------------------------

const address = (row: number, column: number): string => `${columnName(column)}${row + 1}`;

function worksheetOf(sheet: XlsxOutSheet, ranged: ReadonlySet<string>, noted = false, conditional = ''): string {
  const byRow = new Map<number, XlsxOutCell[]>();
  let lastRow = 0;
  let lastColumn = 0;
  for (const cell of sheet.cells) {
    const list = byRow.get(cell.row) ?? [];
    list.push(cell);
    byRow.set(cell.row, list);
    lastRow = Math.max(lastRow, cell.row);
    lastColumn = Math.max(lastColumn, cell.column);
  }
  const rows = [...new Set([...byRow.keys(), ...sheet.rows.keys()])].sort((a, b) => a - b);

  const data = rows
    .map(row => {
      const shape = sheet.rows.get(row);
      const attributes =
        shape === undefined
          ? ''
          : `${shape.height === undefined ? '' : ` ht="${round(shape.height, 2)}"${shape.custom ? ' customHeight="1"' : ''}`}${
              shape.hidden ? ' hidden="1"' : ''
            }`;
      const cells = (byRow.get(row) ?? [])
        .sort((a, b) => a.column - b.column)
        .map(cell => cellOf(cell, ranged))
        .join('');
      return `<row r="${row + 1}"${attributes}>${cells}</row>`;
    })
    .join('');

  const columns = [...sheet.columnWidths]
    .sort((a, b) => a[0] - b[0])
    .map(
      ([column, width]) =>
        `<col min="${column + 1}" max="${column + 1}" width="${width === 0 ? 0 : round(width, 4)}" customWidth="1"${
          width === 0 ? ' hidden="1"' : ''
        }/>`
    )
    .join('');

  const pane = paneOf(sheet.frozenRows, sheet.frozenColumns);
  const merges = sheet.merges
    .map(rect => `<mergeCell ref="${address(rect.firstRow, rect.firstColumn)}:${address(rect.lastRow, rect.lastColumn)}"/>`)
    .join('');

  return xml(
    `<worksheet xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}">` +
      `<dimension ref="A1${lastRow + lastColumn === 0 ? '' : `:${address(lastRow, lastColumn)}`}"/>` +
      `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>` +
      (columns === '' ? '' : `<cols>${columns}</cols>`) +
      `<sheetData>${data}</sheetData>` +
      (merges === '' ? '' : `<mergeCells count="${sheet.merges.length}">${merges}</mergeCells>`) +
      // After the merges and before the drawing, which is the order the
      // schema has them in and the order Excel refuses a file without.
      conditional +
      writeValidations(sheet.validations ?? []).xml +
      // The comments' drawing: Excel shows a comment only through the
      // shape this names; see `notesOf`.
      (noted ? `<legacyDrawing r:id="rId2"/>` : '') +
      `</worksheet>`
  );
}

function round(value: number, places: number): string {
  return String(Number(value.toFixed(places)));
}

function paneOf(rows: number, columns: number): string {
  if (rows === 0 && columns === 0) {
    return '';
  }
  const active = rows > 0 && columns > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
  return (
    `<pane${columns > 0 ? ` xSplit="${columns}"` : ''}${rows > 0 ? ` ySplit="${rows}"` : ''}` +
    ` topLeftCell="${address(rows, columns)}" activePane="${active}" state="frozen"/>`
  );
}

function cellOf(cell: XlsxOutCell, ranged: ReadonlySet<string>): string {
  const at = address(cell.row, cell.column);
  const style = cell.style === 0 ? '' : ` s="${cell.style}"`;
  const { value } = cell;
  if (cell.input.startsWith('=')) {
    const array = cell.spill !== undefined || withIntersections(cell.input, ranged) !== cell.input;
    const formula = excelFormula(cell.input, array);
    const ref =
      cell.spill === undefined
        ? at
        : `${at}:${address(cell.row + cell.spill.rows - 1, cell.column + cell.spill.columns - 1)}`;
    const f = array ? `<f t="array" ref="${ref}">${escape(formula)}</f>` : `<f>${escape(formula)}</f>`;
    return `<c r="${at}"${style}${typeAttribute(value, true)}>${f}${valueElement(value)}</c>`;
  }
  if (value === null) {
    return `<c r="${at}"${style}/>`;
  }
  if (typeof value === 'string') {
    return `<c r="${at}"${style} t="inlineStr"><is><t xml:space="preserve">${escape(value)}</t></is></c>`;
  }
  return `<c r="${at}"${style}${typeAttribute(value, false)}>${valueElement(value)}</c>`;
}

function typeAttribute(value: CellValue, formula: boolean): string {
  if (typeof value === 'string') {
    return formula ? ' t="str"' : ' t="inlineStr"';
  }
  if (typeof value === 'boolean') {
    return ' t="b"';
  }
  if (isError(value)) {
    return ' t="e"';
  }
  return '';
}

function valueElement(value: CellValue): string {
  if (value === null) {
    return '';
  }
  if (typeof value === 'boolean') {
    return `<v>${value ? 1 : 0}</v>`;
  }
  if (isError(value)) {
    return `<v>${escape(value.code)}</v>`;
  }
  return `<v>${escape(String(value))}</v>`;
}

// ---------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------

/**
 * Functions Excel added after the file format, which a file names with
 * a prefix so an older Excel can say it does not have them rather than
 * mistake them for names.
 */
const XLFN = new Set([
  'IFNA',
  'IFS',
  'SWITCH',
  'XLOOKUP',
  'XMATCH',
  'CONCAT',
  'TEXTJOIN',
  'MAXIFS',
  'MINIFS',
  'DAYS',
  'ISOWEEKNUM',
  'SEQUENCE',
  'UNIQUE',
  'STDEV.S',
  'STDEV.P',
  'VAR.S',
  'VAR.P',
  'MODE.SNGL',
  'PERCENTILE.INC',
  'PERCENTILE.EXC',
  'QUARTILE.INC',
  'QUARTILE.EXC',
  'RANK.EQ',
  'RANK.AVG',
  'CEILING.MATH',
  'FLOOR.MATH',
  'NORM.DIST',
  'NORM.INV',
  'NORM.S.DIST',
  'NORM.S.INV'
]);
/** And two more that carry the worksheet prefix as well. */
const XLWS = new Set(['FILTER', 'SORT']);

/** A formula as Excel's file format spells it; see the file's header. */
export function excelFormula(input: string, array: boolean): string {
  const body = input.startsWith('=') ? input.slice(1) : input;
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return body;
  }
  const edits: { at: number; to: number; text: string }[] = [];
  for (let at = 0; at < tokens.length; at++) {
    const token = tokens[at];
    if (token.kind === 'word' && tokens[at + 1]?.kind === 'open') {
      const name = token.value.toUpperCase();
      const prefix = XLWS.has(name) ? '_xlfn._xlws.' : XLFN.has(name) ? '_xlfn.' : '';
      if (prefix !== '') {
        edits.push({ at: token.start, to: token.start, text: prefix });
      }
      continue;
    }
    // `A1#` is written as the call Excel's files use for it.
    if (token.kind === 'operator' && token.value === '#') {
      let first = at - 1;
      if (tokens[first - 1]?.kind === 'sheet') {
        first--;
      }
      edits.push({ at: tokens[first].start, to: tokens[first].start, text: '_xlfn.ANCHORARRAY(' });
      edits.push({ at: token.start, to: token.end, text: ')' });
      continue;
    }
    if (token.kind === 'operator' && token.value === '@') {
      if (!array) {
        edits.push({ at: token.start, to: token.end, text: '' });
        continue;
      }
      const last = operandEnd(tokens, at + 1);
      edits.push({ at: token.start, to: token.end, text: '_xlfn.SINGLE(' });
      edits.push({ at: tokens[last].end, to: tokens[last].end, text: ')' });
    }
  }
  let text = body;
  for (const edit of edits.sort((a, b) => b.at - a.at || b.to - a.to)) {
    text = text.slice(0, edit.at) + edit.text + text.slice(edit.to);
  }
  return text;
}

/** The last token of the operand that starts at `first`: a reference, a range, a call or a bracket. */
function operandEnd(tokens: readonly Token[], first: number): number {
  let at = first;
  if (tokens[at]?.kind === 'sheet') {
    at++;
  }
  const token = tokens[at];
  if (token === undefined) {
    return first;
  }
  const closing = (open: number): number => {
    let depth = 0;
    for (let next = open; next < tokens.length; next++) {
      if (tokens[next].kind === 'open') {
        depth++;
      } else if (tokens[next].kind === 'close' && --depth === 0) {
        return next;
      }
    }
    return tokens.length - 1;
  };
  if (token.kind === 'open') {
    return closing(at);
  }
  if (token.kind === 'word') {
    if (tokens[at + 1]?.kind === 'open') {
      return closing(at + 1);
    }
    if (tokens[at + 1]?.kind === 'colon' && tokens[at + 2]?.kind === 'word') {
      return at + 2;
    }
  }
  return at;
}

/**
 * A sheet's notes as Excel's comments: the comments part, the drawing
 * Excel will not show a comment without, and the relationships from
 * the worksheet to both.
 *
 * The drawing is VML, Office's pre-2007 vector format, and it is not
 * optional: a comments part on its own is read by Excel and never
 * shown, and some versions say the file needs repairing. One hidden
 * shape per note, anchored beside its cell, is what Excel itself
 * writes for a note nobody has moved. The comments have no author —
 * a note here does not record one — so the author list is one empty
 * name, which Excel shows as no name at all.
 */
function notesOf(notes: readonly { row: number; column: number; text: string }[], sheet: number): { name: string; text: string }[] {
  const comments = notes
    .map(
      note =>
        `<comment ref="${address(note.row, note.column)}" authorId="0"><text><t xml:space="preserve">${escape(note.text)}</t></text></comment>`
    )
    .join('');
  const shapes = notes
    .map((note, at) => {
      const anchor = [note.column + 1, 15, Math.max(0, note.row - 1), 10, note.column + 3, 15, note.row + 3, 4].join(', ');
      return (
        `<v:shape id="_x0000_s${sheet * 1024 + at + 1}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:${at + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
        `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>` +
        `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
        `<x:Row>${note.row}</x:Row><x:Column>${note.column}</x:Column></x:ClientData></v:shape>`
      );
    })
    .join('');
  return [
    {
      name: `xl/comments${sheet}.xml`,
      text: xml(`<comments xmlns="${MAIN}"><authors><author></author></authors><commentList>${comments}</commentList></comments>`)
    },
    {
      name: `xl/drawings/vmlDrawing${sheet}.vml`,
      text:
        `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
        `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheet}"/></o:shapelayout>` +
        `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
        `${shapes}</xml>`
    },
    {
      name: `xl/worksheets/_rels/sheet${sheet}.xml.rels`,
      text: xml(
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">` +
          `<Relationship Id="rId1" Type="${RELATIONSHIPS}/comments" Target="../comments${sheet}.xml"/>` +
          `<Relationship Id="rId2" Type="${RELATIONSHIPS}/vmlDrawing" Target="../drawings/vmlDrawing${sheet}.vml"/>` +
          `</Relationships>`
      )
    }
  ];
}
