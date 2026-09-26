import { describe, expect, it } from 'vitest';

import { readXlsx } from './Xlsx';

/**
 * Charts read from an `.xlsx` as Excel writes them — Phase 23.
 *
 * A two-cell anchor, which is what Excel writes for a chart nobody has
 * told to float; each kind this sheet draws; and the three things a
 * file can chart that this sheet cannot, counted rather than drawn as
 * something else.
 */
const LIMITS = { rows: 10_000, columns: 100 };
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

function anchored(chartRef: string): string {
  return (
    '<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
    `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="${chartRef}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
  );
}

function book(charts: readonly string[], anchors = charts.map((_, at) => anchored(`rId${at + 1}`)).join('')) {
  const parts: Record<string, string> = {
    'xl/workbook.xml': '<workbook><sheets><sheet name="Sales" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships><Relationship Id="rId5" Target="../drawings/drawing1.xml"/></Relationships>',
    'xl/drawings/drawing1.xml': `<xdr:wsDr>${anchors}</xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': `<Relationships>${charts
      .map((_, at) => `<Relationship Id="rId${at + 1}" Target="../charts/chart${at + 1}.xml"/>`)
      .join('')}</Relationships>`
  };
  charts.forEach((chart, at) => (parts[`xl/charts/chart${at + 1}.xml`] = `<c:chartSpace ${C}><c:chart>${chart}</c:chart></c:chartSpace>`));
  return readXlsx(path => parts[path] ?? null, LIMITS);
}

const series = (values: string, name = '', categories = '') =>
  `<c:ser>${name === '' ? '' : `<c:tx><c:strRef><c:f>${name}</c:f></c:strRef></c:tx>`}${
    categories === '' ? '' : `<c:cat><c:strRef><c:f>${categories}</c:f></c:strRef></c:cat>`
  }<c:val><c:numRef><c:f>${values}</c:f></c:numRef></c:val></c:ser>`;

describe('charts from an .xlsx', () => {
  it('reads a column chart: its title, legend, the range its series cover, and where it is', () => {
    const read = book([
      '<c:title><c:tx><c:rich><a:p><a:r><a:t>Units </a:t></a:r><a:r><a:t>sold</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
        `<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series('Sales!$B$2:$B$9', 'Sales!$B$1', 'Sales!$A$2:$A$9')}${series('Sales!$C$2:$C$9', 'Sales!$C$1', 'Sales!$A$2:$A$9')}</c:barChart></c:plotArea>` +
        '<c:legend><c:legendPos val="r"/></c:legend>'
    ]);
    const [chart] = read.sheets[0].charts;
    expect(chart).toMatchObject({
      kind: 'column',
      title: 'Units sold',
      legend: true,
      firstRow: 0,
      firstColumn: 0,
      lastRow: 8,
      lastColumn: 2,
      anchor: { kind: 'cells', fromColumn: 4, fromColumnOffset: 10, fromRow: 1, toColumn: 9, toRow: 13 }
    });
  });

  it('reads each kind this sheet draws', () => {
    const plot = (inner: string) => `<c:autoTitleDeleted val="1"/><c:plotArea>${inner}</c:plotArea>`;
    const s = series('Sales!$B$2:$B$9');
    const read = book([
      plot(`<c:barChart><c:barDir val="bar"/>${s}</c:barChart>`),
      plot(`<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${s}</c:barChart>`),
      plot(`<c:lineChart>${s}</c:lineChart>`),
      plot(`<c:areaChart>${s}</c:areaChart>`),
      plot(`<c:doughnutChart>${s}</c:doughnutChart>`),
      plot(`<c:scatterChart><c:ser><c:xVal><c:numRef><c:f>Sales!$A$2:$A$9</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sales!$B$2:$B$9</c:f></c:numRef></c:yVal></c:ser></c:scatterChart>`)
    ]);
    expect(read.sheets[0].charts.map(chart => [chart.kind, chart.title, chart.legend])).toEqual([
      ['bar', '', false],
      ['stacked', '', false],
      ['line', '', false],
      ['area', '', false],
      ['pie', '', false],
      ['scatter', '', false]
    ]);
  });

  it('counts a radar, a combination, and a chart of another sheet as left out', () => {
    const read = book([
      `<c:plotArea><c:radarChart>${series('Sales!$B$2:$B$9')}</c:radarChart></c:plotArea>`,
      `<c:plotArea><c:barChart>${series('Sales!$B$2:$B$9')}</c:barChart><c:lineChart>${series('Sales!$C$2:$C$9')}</c:lineChart></c:plotArea>`,
      `<c:plotArea><c:lineChart>${series('Other!$B$2:$B$9')}</c:lineChart></c:plotArea>`
    ]);
    expect(read.sheets[0].charts).toEqual([]);
    expect(read.leftOut.charts).toBe(3);
  });

  it('passes over a picture in the drawing without counting it', () => {
    const picture =
      '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic/><xdr:clientData/></xdr:oneCellAnchor>';
    const read = book([], picture);
    expect(read.sheets[0].charts).toEqual([]);
    expect(read.leftOut.charts).toBe(0);
  });
});
