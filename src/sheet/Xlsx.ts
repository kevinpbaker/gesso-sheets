import { columnIndex } from './A1';
import { bindingNamesOf, type Ast } from './Ast';
import { DEFAULT_FORMAT, GENERAL, NO_BORDERS, PLAIN, type CellEdge, type CellFormat, type CellPaint, type DatePattern, type NumberFormat } from './Format';
import { isSheetFunction } from './Functions';
import type { MergeRect } from './Merges';
import { parseFormula } from './Parser';
import { rewriteFormula } from './Rewrite';
import { withIntersections } from './Legacy';
import { literalOf } from './Workbook';
import type { Validation } from './Validation';
import { child, children, parseXml, type XmlElement } from './Xml';
import { readConditionals, readValidations, type PendingList } from './XlsxRules';
import { readCharts, type XlsxChart } from './XlsxCharts';
import type { ConditionalPaint, ConditionalRule } from './Conditional';
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
   * For a formula, the value the file says it last calculated to, as
   * text — which is how a producer's arithmetic can be checked against
   * this sheet's. Undefined for a cell that is not a formula.
   */
  readonly cached?: string;
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
  /**
   * Rows somebody set a height for, by index, as a share of the file's
   * default row height — 2 is a row twice as tall. A share rather than
   * points because a sheet's rows are not measured in points, and a
   * row the default height in Excel is the default height here.
   * Heights Excel fitted to the contents are left out: this sheet fits
   * its own rows.
   */
  readonly rowHeights: ReadonlyMap<number, number>;
  readonly merges: readonly MergeRect[];
  readonly frozenRows: number;
  readonly frozenColumns: number;
  /**
   * The sheet's comments, as notes: where each one is and what it
   * says. Excel's threaded comments also write a plain comment for
   * every thread, which is the part read here — the replies are
   * joined into it by Excel already.
   */
  readonly notes: readonly { readonly row: number; readonly column: number; readonly text: string }[];
  /** What the sheet's cells may hold; see `XlsxRules.readValidations`. */
  readonly validations: readonly Validation[];
  /** The formats that think; see `XlsxRules.readConditionals`. */
  readonly conditional: readonly ConditionalRule[];
  /** The charts over the sheet; see `XlsxCharts.readCharts`. */
  readonly charts: readonly XlsxChart[];
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
  /**
   * Names that hold a formula rather than a range — a named `LAMBDA`,
   * which Excel writes `_xlfn.LAMBDA(_xlpm.x, …)`, or a named
   * calculation — as this sheet spells them, `=` and all.
   */
  readonly formulaNames: readonly { readonly name: string; readonly formula: string }[];
  /** Formulas whose last value was kept, because they could not be. */
  readonly valuesKept: number;
  /** Rows and columns past the sheet's edge, left out. */
  readonly cut: { readonly rows: number; readonly columns: number };
  /** Defined names that are neither a plain range nor a formula this sheet can read, and so were not kept. */
  readonly namesSkipped: number;
  /** Excel's iterative calculation, when the workbook turns it on; see `Workbook.iteration`. */
  readonly iteration: { readonly count: number; readonly delta: number } | null;
  /**
   * Rules the file has that this sheet cannot keep, by what they are:
   * `validations`, and in later phases the rest. For the import's
   * sentence, which says how many were left out rather than nothing.
   */
  readonly leftOut: Readonly<Record<string, number>>;
}

export class XlsxError extends Error {}

/** How much of a file fits: this application's sheet size. */
export interface XlsxLimits {
  readonly rows: number;
  readonly columns: number;
}

/** Reads a workbook from the bytes of an `.xlsx`. */
export async function openXlsx(bytes: Uint8Array, inflate: Inflate, limits: XlsxLimits): Promise<XlsxBook> {
  // An OLE compound file: what Excel writes for a workbook with a
  // password to open it, and for the old binary .xls.
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    throw new XlsxError('It is password-protected, or an old .xls; save it from Excel as an .xlsx without a password.');
  }
  let entries;
  try {
    entries = zipEntries(bytes);
  } catch {
    throw new XlsxError('It is not an Excel workbook; an .xlsx is a zip file, and this is not one.');
  }
  const decoder = new TextDecoder();
  // Some producers write Windows paths into the zip: `xl\\workbook.xml`.
  const byName = new Map(entries.map(entry => [entry.name.replace(/\\/g, '/').replace(/^\//, ''), entry]));
  const parts = new Map<string, string>();
  for (const [name, entry] of byName) {
    // Only the parts a reader of cells needs: the workbook, its
    // relationships, the strings, the styles and the worksheets. A
    // file's images and pivot caches can be most of its bytes, and
    // inflating them to throw them away would be most of the time.
    if (
      /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|styles\.xml|theme\/theme1\.xml|worksheets\/[^/]+\.xml|worksheets\/_rels\/[^/]+\.xml\.rels|comments[^/]*\.xml|drawings\/[^/]+\.xml|drawings\/_rels\/[^/]+\.xml\.rels|charts\/chart[^/]*\.xml)$/i.test(
        name
      )
    ) {
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
  const theme = themeColours(read('xl/theme/theme1.xml'));
  const formats = styles(read('xl/styles.xml'), theme);
  const dxfs = differentialFormats(read('xl/styles.xml'), theme);

  const names: XlsxName[] = [];
  const formulaNames: { name: string; formula: string }[] = [];
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
      const formula = clean(defined.text.trim().replace(/^=/, ''));
      if (formula !== '' && parses(formula)) {
        formulaNames.push({ name, formula: `=${formula}` });
      } else {
        namesSkipped++;
      }
      continue;
    }
    names.push({ name, ...range });
  }
  // A bare name parses as a call with no arguments, so a formula that
  // reads one is runnable when the name is one of these — and one that
  // calls a named LAMBDA, when it is one of the formulas.
  const callable = new Set(formulaNames.map(each => each.name.toUpperCase()));
  const known = new Set([...names.map(each => each.name.toUpperCase()), ...callable]);
  // A name for one cell is one value already, and needs no `@`.
  const ranged = new Set(
    names
      .filter(each => each.firstRow !== each.lastRow || each.firstColumn !== each.lastColumn)
      .map(each => each.name.toUpperCase())
  );
  const pr = child(workbook, 'workbookPr')?.attributes.date1904;
  const date1904 = pr === '1' || pr === 'true';

  const cut = { rows: 0, columns: 0 };
  let valuesKept = 0;
  const sheets: XlsxSheet[] = [];
  const lists: { sheet: number; pending: readonly PendingList[] }[] = [];
  let validationsLeftOut = 0;
  let conditionalLeftOut = 0;
  let chartsLeftOut = 0;
  for (const entry of children(child(workbook, 'sheets'), 'sheet')) {
    const target = targets.get(entry.attributes.id ?? '');
    const text = target === undefined ? null : read(target);
    if (text === null) {
      continue;
    }
    const sheetName = entry.attributes.name ?? `Sheet${sheets.length + 1}`;
    const root = parseXml(text);
    const sheet = worksheet(sheetName, root, strings, formats, limits, { known, ranged, callable }, date1904, dxfs);
    const charted = target === undefined ? { charts: [], skipped: 0 } : readCharts(root, target, sheetName, read);
    chartsLeftOut += charted.skipped;
    valuesKept += sheet.valuesKept;
    cut.rows = Math.max(cut.rows, sheet.cutRows);
    cut.columns = Math.max(cut.columns, sheet.cutColumns);
    const commentsAt = target === undefined ? null : commentsTarget(read(relsPathOf(target)), target);
    const notes = comments(commentsAt === null ? null : read(commentsAt)).filter(
      note => note.row < limits.rows && note.column < limits.columns
    );
    lists.push({ sheet: sheets.length, pending: sheet.pendingLists });
    validationsLeftOut += sheet.validationsSkipped;
    conditionalLeftOut += sheet.conditionalSkipped;
    sheets.push({ ...sheet.sheet, notes, charts: charted.charts });
  }
  // A list whose values are a range of cells, filled from those cells
  // now every sheet is read — a list on one sheet usually names a
  // column on another. Its values are the cells as they are now: the
  // rule keeps the list and not the reference, so a value added to the
  // range later is not in it. What is in the cells is what was typed,
  // for a literal, or the file's own answer, for a formula.
  for (const { sheet: at, pending } of lists) {
    const sheet = sheets[at];
    const validations = [...sheet.validations];
    const dropped = new Set<number>();
    for (const pendingList of pending) {
      // A defined name is its range, read off the workbook's names.
      const defined =
        pendingList.name === undefined ? null : names.find(each => each.name.toUpperCase() === pendingList.name?.toUpperCase());
      const list = defined == null ? pendingList : { ...pendingList, ...defined };
      const from = list.sheet === null ? sheet : sheets.find(each => each.name.toUpperCase() === list.sheet?.toUpperCase());
      const values: string[] = [];
      for (const cell of from?.cells ?? []) {
        if (cell.row >= list.firstRow && cell.row <= list.lastRow && cell.column >= list.firstColumn && cell.column <= list.lastColumn) {
          const shown = cell.input.startsWith('=') ? (cell.cached ?? '') : cell.input;
          if (shown !== '' && !values.includes(shown)) {
            values.push(shown);
          }
        }
      }
      const held = validations[list.at];
      if (values.length === 0 || held === undefined) {
        dropped.add(list.at);
        continue;
      }
      validations[list.at] = { ...held, rule: { kind: 'list', values } };
    }
    validationsLeftOut += dropped.size;
    sheets[at] = { ...sheet, validations: validations.filter((_, index) => !dropped.has(index)) };
  }
  if (sheets.length === 0) {
    throw new XlsxError('It has no worksheets this can read.');
  }
  const calc = child(workbook, 'calcPr')?.attributes;
  const iterate = calc?.iterate === '1' || calc?.iterate === 'true';
  const count = Number(calc?.iterateCount ?? 100);
  const delta = Number(calc?.iterateDelta ?? 0.001);
  const iteration = iterate
    ? { count: Number.isInteger(count) && count > 0 ? count : 100, delta: delta > 0 ? delta : 0.001 }
    : null;
  return { sheets, names, formulaNames, valuesKept, cut, namesSkipped, iteration, leftOut: { validations: validationsLeftOut, 'conditional formats': conditionalLeftOut, charts: chartsLeftOut } };
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

/** Where a part's own relationships are: `xl/worksheets/_rels/sheet1.xml.rels`. */
function relsPathOf(part: string): string {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

/**
 * The comments part a worksheet points at, or null.
 *
 * Resolved against the worksheet's own folder, which is what a
 * relationship's target is relative to — `../comments1.xml` from
 * `xl/worksheets/` is `xl/comments1.xml`.
 */
function commentsTarget(text: string | null, part: string): string | null {
  if (text === null) {
    return null;
  }
  const folder = part.slice(0, part.lastIndexOf('/') + 1);
  for (const relationship of children(parseXml(text), 'Relationship')) {
    if (!(relationship.attributes.Type ?? '').endsWith('/comments')) {
      continue;
    }
    const target = relationship.attributes.Target ?? '';
    return normalise(target.startsWith('/') ? target.slice(1) : `${folder}${target}`);
  }
  return null;
}

/**
 * A comment's text as a note says it.
 *
 * Two things come off. The name Excel writes in front, in bold, as
 * `Balaji:` and a line break — the one who wrote it, which a note here
 * does not record; recognised by being a bold first run that ends in a
 * colon, or the comment's own author, because the bold name is the
 * writer's and is not always the name in the file's list of authors.
 * And the wrapping Excel puts round a threaded comment for readers that
 * do not have threads — `[Threaded comment]`, a paragraph about
 * versions of Excel, `Comment:` — leaving the comment and its replies.
 */
function noteText(body: XmlElement, author: string): string {
  const runs = children(body, 'r');
  let said = stringOf(body).replace(/\r\n?/g, '\n');
  const first = runs[0];
  const firstText = child(first, 't')?.text ?? '';
  const bold = child(child(first, 'rPr'), 'b') !== null;
  if (runs.length > 1 && bold && firstText.trimEnd().endsWith(':')) {
    said = said.slice(firstText.length);
  } else if (author !== '' && said.startsWith(`${author}:`)) {
    said = said.slice(author.length + 1);
  }
  if (said.startsWith('[Threaded comment]')) {
    const lines = said.split('\n');
    const from = lines.findIndex(line => line.trim() === 'Comment:');
    if (from !== -1) {
      said = lines
        .slice(from + 1)
        .map(line => line.trim())
        .join('\n');
    }
  }
  return said.trim();
}

/**
 * A comments part as notes.
 *
 * Excel puts the author's name in front of a note's text, in bold, as
 * `Kevin Baker:` and a line break, because a note in Excel shows who
 * wrote it. A note here does not have an author, so the name comes off
 * — and only when it is exactly the comment's own author, so a note
 * that happens to start with a word and a colon keeps it.
 */
function comments(text: string | null): { row: number; column: number; text: string }[] {
  if (text === null) {
    return [];
  }
  const root = parseXml(text);
  const authors = children(child(root, 'authors'), 'author').map(author => author.text);
  const notes: { row: number; column: number; text: string }[] = [];
  for (const comment of children(child(root, 'commentList'), 'comment')) {
    const at = addressOf(comment.attributes.ref ?? '');
    const body = child(comment, 'text');
    if (at === null || body === null) {
      continue;
    }
    const said = noteText(body, authors[Number(comment.attributes.authorId ?? -1)] ?? '');
    if (said !== '') {
      notes.push({ row: at.row, column: at.column, text: said });
    }
  }
  return notes;
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
/**
 * The styles' differential formats, which is what a conditional
 * format paints with: a `dxf` says only what it changes. Its fill is
 * the pattern's *background* colour, where a cell style's is its
 * foreground — the one place the format reverses the two.
 */
function differentialFormats(text: string | null, theme: readonly string[]): ConditionalPaint[] {
  if (text === null) {
    return [];
  }
  return children(child(parseXml(text), 'dxfs'), 'dxf').map(dxf => {
    const font = child(dxf, 'font');
    const pattern = child(child(dxf, 'fill'), 'patternFill');
    const fill = colourIn(child(pattern, 'bgColor') ?? child(pattern, 'fgColor'), theme);
    const color = colourIn(child(font, 'color'), theme);
    const flag = (name: string): boolean => {
      const node = child(font, name);
      return node !== null && node.attributes.val !== '0' && node.attributes.val !== 'false';
    };
    return {
      ...(fill === '' ? {} : { fill }),
      ...(color === '' ? {} : { color }),
      ...(flag('b') ? { bold: true } : {}),
      ...(flag('i') ? { italic: true } : {})
    };
  });
}

function styles(text: string | null, theme: readonly string[]): CellFormat[] {
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
  const colourOf = (element: XmlElement | null): string => colourIn(element, theme);
  const fillOf = (fill: XmlElement | undefined): string => {
    const pattern = child(fill, 'patternFill');
    return pattern?.attributes.patternType === 'solid' ? colourOf(child(pattern, 'fgColor')) : '';
  };
  const edgeOf = (edge: XmlElement | null): CellEdge => {
    const width = EDGE_WIDTHS[edge?.attributes.style ?? ''] ?? 0;
    return width === 0 ? { width: 0, color: '' } : { width, color: colourOf(child(edge, 'color')) };
  };
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
 * A colour as `#rrggbb`, or '' for the default.
 *
 * Three spellings. `rgb` is the colour itself, as ARGB. `theme` is a
 * slot in the workbook's theme, lightened or darkened by `tint` — which
 * is how Excel writes nearly every colour a person picks from its
 * palette, so dropping it lost most of the colour in a real workbook.
 * `indexed` is the 1995 palette of 64, which old files and some
 * exporters still use; 64 and above are "the system's own colour",
 * which is the default here. Anything else is the default too, which
 * is the right failure for a colour: text in it is always legible.
 */
function colourIn(element: XmlElement | null, theme: readonly string[]): string {
  if (element === null) {
    return '';
  }
  const rgb = element.attributes.rgb;
  let hex: string | undefined;
  if (rgb !== undefined && /^[0-9a-f]{8}$/i.test(rgb)) {
    hex = rgb.slice(2);
  } else if (element.attributes.theme !== undefined) {
    hex = theme[Number(element.attributes.theme)];
  } else if (element.attributes.indexed !== undefined) {
    hex = INDEXED[Number(element.attributes.indexed)];
  }
  if (hex === undefined) {
    return '';
  }
  const tint = Number(element.attributes.tint ?? 0);
  return `#${(Number.isFinite(tint) && tint !== 0 ? tinted(hex, tint) : hex).toLowerCase()}`;
}

/**
 * The theme's twelve colours, in the order a `theme` attribute counts
 * them — which is not the order the theme part writes them. The first
 * two pairs are swapped: SpreadsheetML's 0 is the light background and
 * 1 the dark text, where the part lists dark first.
 */
function themeColours(text: string | null): string[] {
  if (text === null) {
    return [];
  }
  const scheme = findElement(parseXml(text), 'clrScheme');
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map(name => {
    const slot = child(scheme, name);
    const colour = slot?.children[0];
    return colour?.attributes.val !== undefined && colour.name === 'srgbClr'
      ? colour.attributes.val
      : (colour?.attributes.lastClr ?? '');
  });
}

function findElement(element: XmlElement, name: string): XmlElement | null {
  if (element.name === name) {
    return element;
  }
  for (const each of element.children) {
    const found = findElement(each, name);
    if (found !== null) {
      return found;
    }
  }
  return null;
}

/**
 * A colour lightened (a tint above zero) or darkened (below), as
 * Excel does it: on the lightness in HSL, moved that fraction of the
 * way to white or to black.
 */
function tinted(hex: string, tint: number): string {
  const [r, g, b] = [0, 2, 4].map(at => parseInt(hex.slice(at, at + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let lightness = (max + min) / 2;
  const delta = max - min;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  let hue = 0;
  if (delta !== 0) {
    hue = max === r ? ((g - b) / delta) % 6 : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  }
  lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = chroma * (1 - Math.abs((((hue % 6) + 6) % 6) % 2 - 1));
  const m = lightness - chroma / 2;
  const sector = Math.floor((((hue % 6) + 6) % 6));
  const [r1, g1, b1] = [
    [chroma, x, 0],
    [x, chroma, 0],
    [0, chroma, x],
    [0, x, chroma],
    [x, 0, chroma],
    [chroma, 0, x]
  ][sector];
  return [r1, g1, b1].map(value => Math.round((value + m) * 255).toString(16).padStart(2, '0')).join('');
}

/** The legacy palette an `indexed` colour counts into; Excel's defaults. */
const INDEXED: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
];

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
  readonly pendingLists: readonly PendingList[];
  readonly validationsSkipped: number;
  readonly conditionalSkipped: number;
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
  { known, ranged, callable }: { known: ReadonlySet<string>; ranged: ReadonlySet<string>; callable: ReadonlySet<string> },
  date1904: boolean,
  dxfs: readonly ConditionalPaint[] = []
): ReadSheet {
  const cells: XlsxCell[] = [];
  const styled: { row: number; column: number; style: number }[] = [];
  const hiddenRows: number[] = [];
  const columnWidths = new Map<number, number>();
  const rowHeights = new Map<number, number>();
  const defaultHeight = Number(child(root, 'sheetFormatPr')?.attributes.defaultRowHeight ?? 15) || 15;
  /** A shared formula's first cell, by its `si`: the formula and where it stood. */
  const shared = new Map<string, { input: string; row: number; column: number }>();
  /** The areas array formulas cover, whose other cells are the formulas' to fill. */
  const arrays: { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number }[] = [];
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
    const custom = rowElement.attributes.customHeight === '1' || rowElement.attributes.customHeight === 'true';
    const height = Number(rowElement.attributes.ht);
    if (custom && Number.isFinite(height) && height > 0) {
      rowHeights.set(row, height / defaultHeight);
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
      const formulaElement = child(cell, 'f');
      if (formulaElement?.attributes.t === 'array') {
        const area = rangeOf(formulaElement.attributes.ref ?? '');
        if (area !== null) {
          arrays.push(area);
        }
      } else if (formulaElement === null && arrays.some(area => inside(area, at.row, at.column))) {
        // A cell an array formula spilled into: Excel keeps its value
        // there, and here the formula fills it again. Kept, it would
        // be in the way and the formula would say #SPILL!.
        if (style > 0) {
          styled.push({ row: at.row, column: at.column, style });
        }
        continue;
      }
      const read = cellInput(cell, at.row, at.column, strings, shared, known, ranged, callable);
      if (read === null) {
        if (style > 0) {
          styled.push({ row: at.row, column: at.column, style });
        }
        continue;
      }
      if (read.kept) {
        valuesKept++;
      }
      const dated = date1904 && isDated(formats[Math.max(0, style)]);
      cells.push({
        row: at.row,
        column: at.column,
        input: dated ? shifted(read.input) : read.input,
        style: Math.max(0, style),
        asText: read.asText,
        ...(read.cached === undefined ? {} : { cached: dated ? shifted(read.cached) : read.cached })
      });
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

  const validated = readValidations(root, date1904);
  const conditioned = readConditionals(root, dxfs, text => `=${clean(text)}`);
  return {
    // The notes are in a part of their own, read beside this one.
    sheet: {
      name,
      cells,
      styled,
      formats,
      columnWidths,
      hiddenRows,
      rowHeights,
      merges,
      frozenRows,
      frozenColumns,
      notes: [],
      charts: [],
      validations: validated.validations,
      conditional: conditioned.rules
    },
    pendingLists: validated.pending,
    validationsSkipped: validated.skipped,
    conditionalSkipped: conditioned.skipped,
    valuesKept,
    cutRows,
    cutColumns
  };
}

/**
 * The 1904 date system, which Excel for the Mac used: serial 0 is
 * 1 January 1904, 1,462 days after this sheet's. A date in such a file
 * is moved to this sheet's count as it comes in, so `YEAR` and `DATE`
 * and every date format agree with it. Only a date is moved — a
 * duration or a time of day is the same in both systems — and only a
 * number: a formula that works out a date from a written-in serial is
 * the one thing that comes out 4 years early, which is rare enough to
 * accept for a sheet with only one date system.
 */
const SHIFT_1904 = 1462;

function isDated(format: CellFormat | undefined): boolean {
  const kind = format?.number.kind;
  return kind === 'date' || kind === 'datetime';
}

function shifted(input: string): string {
  if (input === '' || input.startsWith('=')) {
    return input;
  }
  const number = Number(input);
  return Number.isFinite(number) ? String(number + SHIFT_1904) : input;
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
  known: ReadonlySet<string>,
  ranged: ReadonlySet<string>,
  callable: ReadonlySet<string>
): { input: string; asText: boolean; kept: boolean; cached?: string } | null {
  const type = cell.attributes.t ?? 'n';
  const raw = child(cell, 'v')?.text ?? null;
  const value = valueOf(type, raw, cell, strings);

  const formula = child(cell, 'f');
  if (formula !== null) {
    const written = formulaOf(formula, row, column, shared, ranged);
    if (written !== null && canRun(written, known, callable)) {
      return { input: written, asText: false, kept: false, ...(value === null ? {} : { cached: value.text }) };
    }
    if (value === null) {
      return null;
    }
    return { ...literal(value), kept: true };
  }
  return value === null ? null : { ...literal(value), kept: false };
}

type Value = { kind: 'text'; text: string } | { kind: 'other'; text: string } | { kind: 'error'; text: string };

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
      return raw === null ? null : { kind: 'error', text: raw };
    default:
      return raw === null || raw === '' ? null : { kind: 'other', text: raw };
  }
}

/** A value as an input, marked text when typing it would not give text back. */
function literal(value: Value): { input: string; asText: boolean } {
  if (value.kind === 'other') {
    return { input: value.text, asText: false };
  }
  if (value.kind === 'error') {
    // An error cell is the error, not its name as text: `ISNA` has to
    // see a `#N/A` there, and `LEN` must not count five characters.
    // Written as a formula, since a typed `#N/A` is text here; a code
    // this sheet does not know (`#SPILL!`, `#GETTING_DATA`) stays text.
    return KNOWN_ERRORS.has(value.text) ? { input: `=${value.text}`, asText: false } : { input: value.text, asText: true };
  }
  const text = value.text;
  if (text === '') {
    // An empty string is not a blank — `COUNTA` counts it and `""=A1`
    // is about it — and an empty input here *clears* a cell, so it is
    // written as the formula that gives it.
    return { input: '=""', asText: false };
  }
  return { input: text, asText: text.startsWith('=') || typeof literalOf(text) !== 'string' };
}

const KNOWN_ERRORS: ReadonlySet<string> = new Set(['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A', '#NUM!', '#SPILL!', '#CALC!']);

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
  shared: Map<string, { input: string; row: number; column: number }>,
  names: ReadonlySet<string>
): string | null {
  const text = formula.text.trim();
  const index = formula.attributes.si;
  // An array formula — typed with Ctrl+Shift+Enter, or a dynamic one
  // from Excel 365 — is an array already; anything else was written
  // when a range where one value is wanted meant one value.
  const meant = (input: string): string => (formula.attributes.t === 'array' ? input : withIntersections(input, names));
  if (formula.attributes.t === 'shared' && index !== undefined) {
    if (text !== '') {
      const input = meant(`=${clean(text)}`);
      shared.set(index, { input, row, column });
      return input;
    }
    const first = shared.get(index);
    return first === undefined ? null : rewriteFormula(first.input, row - first.row, column - first.column);
  }
  return text === '' ? null : meant(`=${clean(text)}`);
}

/**
 * Excel's prefixes for functions newer than the file format, taken
 * off; `_xlfn.SINGLE(…)` is how Excel 365 writes `@(…)` into a file.
 */
function clean(text: string): string {
  return text
    .replace(/_xlfn\.SINGLE\(/gi, '@(')
    .replace(/_xlfn\.ANCHORARRAY\(((?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?(\$?[A-Z]+\$?\d+)\)/gi, '$1$2#')
    .replace(/_xl(?:fn|ws|pm)\./g, '');
}

/** Whether a formula's text parses here. */
function parses(formula: string): boolean {
  try {
    parseFormula(formula);
    return true;
  } catch {
    return false;
  }
}

function inside(area: { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number }, row: number, column: number): boolean {
  return row >= area.firstRow && row <= area.lastRow && column >= area.firstColumn && column <= area.lastColumn;
}

/** Whether this sheet can parse a formula and knows every function and name it calls. */
function canRun(input: string, known: ReadonlySet<string>, callable: ReadonlySet<string> = new Set()): boolean {
  let tree: Ast;
  try {
    tree = parseFormula(input.slice(1));
  } catch {
    return false;
  }
  // A name a LET or LAMBDA in the formula binds is known inside it.
  const bound = bindingNamesOf(tree);
  const knowsAll = (node: Ast): boolean => {
    switch (node.kind) {
      case 'call': {
        const name = node.name.toUpperCase();
        const found =
          isSheetFunction(name) || bound.has(name) || callable.has(name) || (node.args.length === 0 && known.has(name));
        return found && node.args.every(knowsAll);
      }
      case 'invoke':
        return knowsAll(node.callee) && node.args.every(knowsAll);
      case 'unary':
        return knowsAll(node.operand);
      case 'binary':
        return knowsAll(node.left) && knowsAll(node.right);
      default:
        return true;
    }
  };
  // A reference into another workbook — `'[Prices.xlsx]Sheet1'!A1` —
  // parses as a sheet this workbook does not have, and would say
  // #REF!. The other file is not here to read, so the value Excel last
  // read from it is the best there is.
  const external = (node: Ast): boolean => {
    switch (node.kind) {
      case 'ref':
        return node.ref.sheet?.startsWith('[') === true;
      case 'range':
        return node.range.start.sheet?.startsWith('[') === true;
      case 'call':
        return node.args.some(external);
      case 'invoke':
        return external(node.callee) || node.args.some(external);
      case 'unary':
        return external(node.operand);
      case 'binary':
        return external(node.left) || external(node.right);
      default:
        return false;
    }
  };
  return knowsAll(tree) && !external(tree);
}
