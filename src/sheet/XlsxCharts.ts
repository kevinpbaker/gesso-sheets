import type { ChartKind } from './Chart';
import { child, children, parseXml, type XmlElement } from './Xml';

/**
 * Charts, between this sheet and an `.xlsx` — Phase 23.
 *
 * An `.xlsx` chart is three parts and a line: a DrawingML chart part
 * (`xl/charts/chartN.xml`) that says what is drawn, a drawing part
 * (`xl/drawings/drawingN.xml`) that says where, the relationships
 * between them, and a `<drawing>` in the worksheet that names the
 * drawing. Phase 15's chart is a kind, a range, a title, a legend and a
 * rectangle, and the chart part is written from the series that range
 * reads as on screen — each series' name, categories and values spelled
 * out as references, because Excel will not guess.
 *
 * Read the other way, the five kinds this sheet draws are kept — bar
 * and column, clustered or stacked, line, area, pie, scatter — and the
 * range is the rectangle the series' references cover. Anything else a
 * file can chart (a combination, a radar, a 3-D surface, a series on
 * another sheet) is counted as left out rather than drawn as something
 * it is not.
 */

/** One series, as references a formula would write: `'Sales'!$B$2:$B$9`. */
export interface XlsxOutSeries {
  readonly name?: string;
  readonly categories?: string;
  readonly values: string;
}

export interface XlsxOutChart {
  readonly kind: ChartKind;
  readonly title: string;
  readonly legend: boolean;
  readonly series: readonly XlsxOutSeries[];
  /** Where it is, in the sheet's pixels from the top-left of A1. */
  readonly place: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
}

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Pixels to English Metric Units, at 96 pixels an inch. */
const EMU = 9525;

export function chartPartOf(chart: XlsxOutChart): string {
  const pie = chart.kind === 'pie';
  const series = chart.series
    .map((one, at) => {
      const name = one.name === undefined ? '' : `<c:tx><c:strRef><c:f>${escape(one.name)}</c:f></c:strRef></c:tx>`;
      if (chart.kind === 'scatter') {
        const x = one.categories === undefined ? '' : `<c:xVal><c:strRef><c:f>${escape(one.categories)}</c:f></c:strRef></c:xVal>`;
        return `<c:ser><c:idx val="${at}"/><c:order val="${at}"/>${name}<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/></c:marker>${x}<c:yVal><c:numRef><c:f>${escape(one.values)}</c:f></c:numRef></c:yVal></c:ser>`;
      }
      const categories = one.categories === undefined ? '' : `<c:cat><c:strRef><c:f>${escape(one.categories)}</c:f></c:strRef></c:cat>`;
      return `<c:ser><c:idx val="${at}"/><c:order val="${at}"/>${name}${categories}<c:val><c:numRef><c:f>${escape(one.values)}</c:f></c:numRef></c:val></c:ser>`;
    })
    .join('');
  const axes = pie ? '' : '<c:axId val="1"/><c:axId val="2"/>';
  let plot: string;
  switch (chart.kind) {
    case 'column':
    case 'bar':
    case 'stacked':
      plot =
        `<c:barChart><c:barDir val="${chart.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${chart.kind === 'stacked' ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>` +
        `${series}${chart.kind === 'stacked' ? '<c:overlap val="100"/>' : ''}${axes}</c:barChart>`;
      break;
    case 'line':
      plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>`;
      break;
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'scatter':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axes}</c:scatterChart>`;
      break;
  }
  const sideways = chart.kind === 'bar';
  const categoryAxis =
    chart.kind === 'scatter'
      ? `<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>`
      : `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${sideways ? 'l' : 'b'}"/><c:crossAx val="2"/></c:catAx>`;
  const valueAxis = `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${sideways ? 'b' : 'l'}"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>`;
  const title =
    chart.title === ''
      ? '<c:autoTitleDeleted val="1"/>'
      : `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escape(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
  return xml(
    `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${title}<c:plotArea><c:layout/>${plot}${pie ? '' : categoryAxis + valueAxis}</c:plotArea>` +
      `${chart.legend ? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>' : ''}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`
  );
}

/**
 * The drawing that places a sheet's charts: one anchor each, and the
 * id of the relationship each anchor's frame names — `rId1` for the
 * first chart, and on.
 *
 * An absolute anchor, in EMUs from the top-left of the sheet, because
 * that is what this sheet's charts are: a rectangle in pixels, not a
 * span of cells. A reader that prefers cells — Excel, when a row is
 * inserted — works the cells out itself.
 */
export function drawingPartOf(charts: readonly XlsxOutChart[]): string {
  const anchors = charts
    .map(
      (chart, at) =>
        `<xdr:absoluteAnchor><xdr:pos x="${Math.round(chart.place.x * EMU)}" y="${Math.round(chart.place.y * EMU)}"/><xdr:ext cx="${Math.round(chart.place.width * EMU)}" cy="${Math.round(chart.place.height * EMU)}"/>` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${at + 2}" name="Chart ${at + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
        `<a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rId${at + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`
    )
    .join('');
  return xml(`<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${anchors}</xdr:wsDr>`);
}

// ---------------------------------------------------------------------------
// In
// ---------------------------------------------------------------------------

/** Where a chart is, as a file says it: a span of cells with offsets, or a rectangle. */
export type XlsxAnchor =
  | {
      readonly kind: 'cells';
      readonly fromColumn: number;
      readonly fromColumnOffset: number;
      readonly fromRow: number;
      readonly fromRowOffset: number;
      readonly toColumn: number;
      readonly toColumnOffset: number;
      readonly toRow: number;
      readonly toRowOffset: number;
    }
  | { readonly kind: 'pixels'; readonly x: number; readonly y: number; readonly width: number; readonly height: number };

export interface XlsxChart {
  readonly kind: ChartKind;
  readonly title: string;
  readonly legend: boolean;
  /**
   * The sheet its data is on, when that is not the sheet the chart is
   * on — the usual shape of a workbook with a summary sheet of charts.
   */
  readonly sheet?: string;
  /** The rectangle its series' references cover, on that sheet. */
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
  /** Offsets in pixels, as a file's EMUs are converted. */
  readonly anchor: XlsxAnchor;
}

/**
 * A worksheet's charts: the drawing its `<drawing>` names, each frame
 * in it that holds a chart, and each chart part. `read` answers a part
 * by its path; `relsOf` finds a part's own relationships.
 */
export function readCharts(
  worksheet: XmlElement,
  sheetPart: string,
  sheetName: string,
  read: (path: string) => string | null
): { charts: XlsxChart[]; skipped: number } {
  const drawingId = child(worksheet, 'drawing')?.attributes.id;
  if (drawingId === undefined) {
    return { charts: [], skipped: 0 };
  }
  const drawingPart = relationshipTarget(read(relsPathOf(sheetPart)), sheetPart, drawingId);
  const drawingText = drawingPart === null ? null : read(drawingPart);
  if (drawingPart === null || drawingText === null) {
    return { charts: [], skipped: 0 };
  }
  const drawingRels = read(relsPathOf(drawingPart));
  const charts: XlsxChart[] = [];
  let skipped = 0;
  const drawing = parseXml(drawingText);
  for (const anchorNode of drawing.children) {
    if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(anchorNode.name)) {
      continue;
    }
    const frame = child(anchorNode, 'graphicFrame');
    const chartRef = child(child(child(frame, 'graphic'), 'graphicData'), 'chart');
    if (chartRef === null) {
      // A picture or a shape: not a chart, and not this phase's to count.
      continue;
    }
    const chartPart = relationshipTarget(drawingRels, drawingPart, chartRef.attributes.id ?? '');
    const chartText = chartPart === null ? null : read(chartPart);
    const anchor = anchorOf(anchorNode);
    const chart = chartText === null || anchor === null ? null : chartOf(parseXml(chartText), sheetName, anchor);
    if (chart === null) {
      skipped++;
    } else {
      charts.push(chart);
    }
  }
  return { charts, skipped };
}

function chartOf(space: XmlElement, sheetName: string, anchor: XlsxAnchor): XlsxChart | null {
  const chart = child(space, 'chart');
  const plot = child(chart, 'plotArea');
  const kinds = (plot?.children ?? []).filter(node => node.name.endsWith('Chart'));
  // One kind of chart in the plot: a combination is two, and is drawn
  // here as neither.
  if (kinds.length !== 1) {
    return null;
  }
  const node = kinds[0];
  const kind = kindOf(node);
  if (kind === null) {
    return null;
  }
  const refs: string[] = [];
  for (const series of children(node, 'ser')) {
    for (const holder of ['tx', 'cat', 'val', 'xVal', 'yVal']) {
      const reference = child(child(series, holder), 'numRef') ?? child(child(series, holder), 'strRef');
      const formula = child(reference, 'f')?.text.trim();
      if (formula !== undefined && formula !== '') {
        refs.push(formula);
      }
    }
  }
  const box = boxOf(refs, sheetName);
  if (box === null) {
    return null;
  }
  const deleted = child(chart, 'autoTitleDeleted')?.attributes.val;
  const titleNode = child(chart, 'title');
  const title = titleNode === null ? '' : textIn(titleNode).trim();
  return {
    kind,
    title: deleted === '1' || deleted === 'true' ? '' : title,
    legend: child(chart, 'legend') !== null,
    ...box,
    anchor
  };
}

function kindOf(node: XmlElement): ChartKind | null {
  switch (node.name) {
    case 'barChart':
    case 'bar3DChart': {
      const sideways = child(node, 'barDir')?.attributes.val === 'bar';
      const grouping = child(node, 'grouping')?.attributes.val ?? 'clustered';
      if (grouping === 'stacked' || grouping === 'percentStacked') {
        // Stacked is drawn upright here; stacked bars on their side are not.
        return sideways ? null : 'stacked';
      }
      return sideways ? 'bar' : 'column';
    }
    case 'lineChart':
    case 'line3DChart':
      return 'line';
    case 'areaChart':
    case 'area3DChart':
      return 'area';
    case 'pieChart':
    case 'pie3DChart':
    case 'doughnutChart':
      return 'pie';
    case 'scatterChart':
      return 'scatter';
    default:
      return null;
  }
}

/**
 * The rectangle a chart's references cover, and the sheet they are on
 * when it is not the chart's own — when every one of them is a plain
 * range on one sheet, which is the one shape a chart here can read.
 */
function boxOf(
  refs: readonly string[],
  sheetName: string
): { sheet?: string; firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | null {
  if (refs.length === 0) {
    return null;
  }
  let box: { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number } | null = null;
  let on: string | undefined;
  for (const ref of refs) {
    const match = /^(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref);
    if (match === null) {
      return null;
    }
    const named = match[1]?.replace(/''/g, "'") ?? match[2];
    const sheet = named === undefined || named.toUpperCase() === sheetName.toUpperCase() ? undefined : named;
    // Every reference on the same sheet, or the chart reads as neither.
    if (box !== null && (sheet ?? '').toUpperCase() !== (on ?? '').toUpperCase()) {
      return null;
    }
    on = sheet;
    const first = { row: Number(match[4]) - 1, column: columnNumber(match[3]) };
    const last = match[5] === undefined ? first : { row: Number(match[6]) - 1, column: columnNumber(match[5]) };
    const here = {
      firstRow: Math.min(first.row, last.row),
      firstColumn: Math.min(first.column, last.column),
      lastRow: Math.max(first.row, last.row),
      lastColumn: Math.max(first.column, last.column)
    };
    box =
      box === null
        ? here
        : {
            firstRow: Math.min(box.firstRow, here.firstRow),
            firstColumn: Math.min(box.firstColumn, here.firstColumn),
            lastRow: Math.max(box.lastRow, here.lastRow),
            lastColumn: Math.max(box.lastColumn, here.lastColumn)
          };
  }
  return box === null ? null : { ...box, ...(on === undefined ? {} : { sheet: on }) };
}

function anchorOf(node: XmlElement): XlsxAnchor | null {
  const px = (emu: string | undefined): number => Math.round(Number(emu ?? 0) / EMU);
  if (node.name === 'absoluteAnchor') {
    const pos = child(node, 'pos');
    const ext = child(node, 'ext');
    return pos === null || ext === null
      ? null
      : { kind: 'pixels', x: px(pos.attributes.x), y: px(pos.attributes.y), width: px(ext.attributes.cx), height: px(ext.attributes.cy) };
  }
  const from = child(node, 'from');
  if (from === null) {
    return null;
  }
  const cell = (at: XmlElement | null, name: string): number => Number(child(at, name)?.text ?? 0);
  const fromColumn = cell(from, 'col');
  const fromRow = cell(from, 'row');
  const fromColumnOffset = px(child(from, 'colOff')?.text);
  const fromRowOffset = px(child(from, 'rowOff')?.text);
  if (node.name === 'oneCellAnchor') {
    const ext = child(node, 'ext');
    // A span of cells is not known until the widths are, so a one-cell
    // anchor keeps its size in pixels and its corner in cells.
    return {
      kind: 'cells',
      fromColumn,
      fromColumnOffset,
      fromRow,
      fromRowOffset,
      toColumn: -1,
      toColumnOffset: px(ext?.attributes.cx),
      toRow: -1,
      toRowOffset: px(ext?.attributes.cy)
    };
  }
  const to = child(node, 'to');
  return {
    kind: 'cells',
    fromColumn,
    fromColumnOffset,
    fromRow,
    fromRowOffset,
    toColumn: cell(to, 'col'),
    toColumnOffset: px(child(to, 'colOff')?.text),
    toRow: cell(to, 'row'),
    toRowOffset: px(child(to, 'rowOff')?.text)
  };
}

function textIn(node: XmlElement): string {
  if (node.name === 't') {
    return node.text;
  }
  return node.children.map(textIn).join('');
}

function columnNumber(name: string): number {
  let value = 0;
  for (const letter of name.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value - 1;
}

/** Where a part's own relationships are: `xl/drawings/_rels/drawing1.xml.rels`. */
export function relsPathOf(part: string): string {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

function relationshipTarget(text: string | null, part: string, id: string): string | null {
  if (text === null) {
    return null;
  }
  const folder = part.slice(0, part.lastIndexOf('/') + 1);
  for (const relationship of children(parseXml(text), 'Relationship')) {
    if (relationship.attributes.Id === id) {
      const target = relationship.attributes.Target ?? '';
      return normalise(target.startsWith('/') ? target.slice(1) : `${folder}${target}`);
    }
  }
  return null;
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

function xml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

function escape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
