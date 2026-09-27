import { parseAddress, type RangeRef } from '../sheet/A1';
import type { ChartKind } from '../sheet/Chart';
import {
  GENERAL,
  NO_BORDERS,
  PLAIN,
  type CellBorders,
  type CellFormat,
  type CellPaint,
  type NumberFormat
} from '../sheet/Format';
import { COLUMN_WIDTH, GUTTER_WIDTH, HEADER_HEIGHT, ROW_HEIGHT } from './dimensions';
import type { SheetDocument } from './SheetDocument';

/**
 * Something to open onto: a year of a small trading business, built to
 * show what the sheet can do the moment the page loads.
 *
 * Five sheets, and each is a different kind of spreadsheet:
 *
 * - **Dashboard**, which opens first. Four figures across the top, a
 *   paragraph that wraps, twelve months of revenue whose labels are one
 *   formula spilling down a column (`TEXT` over `EDATE` over
 *   `SEQUENCE`), the regions as `SORT(UNIQUE(...))` counted through
 *   `E8#`, the five largest orders as one `SORT(FILTER(...))` block, and
 *   three charts — one of them drawn from another sheet's cells.
 * - **Sales**, the ledger: twenty-four orders, each priced and staffed
 *   by a lookup into `Reference`, with a commission found by an
 *   approximate `XLOOKUP` into a tier table; names (`=SUM(Revenue)`), a
 *   merged banner, a frozen corner, a three-colour scale, rules that
 *   paint what is under target and what is held, a dropdown that
 *   refuses what it does not know, and notes on the cells worth one.
 * - **Forecast**, the what-if: a scenario picked from a dropdown, the
 *   growth it means looked up, and next year projected by a single
 *   formula that spills twelve months (`ROUND` running across an
 *   array), charted; and the payment on a warehouse loan.
 * - **Reference**, the tables everything reads, where an edit to a
 *   price moves every sheet.
 * - **Summary**, the quarter by region, and the function library at
 *   work — a line each from most of its families.
 *
 * Names are resolved against the sheet the formula is written on, so
 * the ones defined here are used only on `Sales`, and every formula
 * that reaches another sheet writes the sheet out.
 *
 * It is not a fixture: specs build their own sheets. Phase 6 replaces
 * it with whatever the repository loaded, so this runs once in the life
 * of a browser profile.
 */
export function seed(document: SheetDocument): void {
  // Every sheet exists before any formula is written, because a formula
  // reading `Sales!A4` can only point anywhere if `Sales` is there.
  document.renameSheet(0, 'Dashboard');
  document.setSheetColour(0, TAB_PURPLE);
  document.setSheetColour(document.addSheet('Sales'), TAB_BLUE);
  document.setSheetColour(document.addSheet('Forecast'), TAB_RED);
  document.setSheetColour(document.addSheet('Reference'), TAB_ORANGE);
  document.setSheetColour(document.addSheet('Summary'), TAB_GREEN);

  document.activate(SHEET.reference);
  reference(document);
  document.activate(SHEET.sales);
  sales(document);
  document.activate(SHEET.summary);
  summary(document);
  document.activate(SHEET.dashboard);
  dashboard(document);
  document.activate(SHEET.forecast);
  forecast(document);

  // On the regions, whose first cell is the whole of them — the
  // formula bar opens on SORT(UNIQUE(...)) — rather than on a tile,
  // whose note would open over the figure beside it.
  document.activate(SHEET.dashboard);
  document.setSelection(MONTHS_HEAD + 1, 4, MONTHS_HEAD + 1, 4);
}

/** Where each sheet is, in tab order. */
export const SHEET = { dashboard: 0, sales: 1, forecast: 2, reference: 3, summary: 4 } as const;

// -------------------------------------------------------------------
// Sales: the ledger
// -------------------------------------------------------------------

function sales(document: SheetDocument): void {
  setWidths(document, [100, 128, 100, 100, 72, 84, 104, 72, 92, 96, 96]);

  document.setCell(0, 0, 'Northwind Trading — orders');
  document.setCell(
    1,
    0,
    `="${ORDERS.length} orders, "&TEXT(MIN(Ordered),"mmmm yyyy")&" to "&TEXT(MAX(Ordered),"mmmm yyyy")&", priced from the Reference sheet"`
  );

  const headings = ['Region', 'Rep', 'Ordered', 'Due', 'Units', 'Price', 'Revenue', 'Share', 'Status', 'Review', 'Commission'];
  headings.forEach((heading, column) => document.setCell(2, column, heading));

  ORDERS.forEach((order, index) => {
    const row = index + 3;
    const line = row + 1;
    document.setCell(row, 0, order.region);
    document.setCell(row, 1, `=VLOOKUP($A${line},${PRICES},2,FALSE)`);
    document.setCell(row, 2, `=DATE(${order.ordered[0]},${order.ordered[1]},${order.ordered[2]})`);
    // Thirty days' credit, landing on a working day.
    document.setCell(row, 3, `=WORKDAY(C${line},20)`);
    document.setCell(row, 4, String(order.units));
    document.setCell(row, 5, `=XLOOKUP($A${line},${REGION_COLUMN},${PRICE_COLUMN})`);
    document.setCell(row, 6, `=E${line}*F${line}`);
    document.setCell(row, 7, `=G${line}/$G$${TOTAL_LINE}`);
    document.setCell(row, 8, `=IF(G${line}>=VLOOKUP($A${line},${PRICES},4,FALSE),"On target","Below")`);
    document.setCell(row, 9, order.review);
    // The tier an order falls in is the largest threshold under it: an
    // approximate match, next smaller, into a table somebody maintains.
    // The tier table's arithmetic is a function with a name; see `salesNames`.
    document.setCell(row, 10, `=Commission(G${line})`);
  });

  document.setCell(TOTAL_ROW, 0, 'Total');
  document.setCell(TOTAL_ROW, 4, '=SUM(Units)');
  document.setCell(TOTAL_ROW, 5, '=AVERAGE(Price)');
  document.setCell(TOTAL_ROW, 6, '=SUM(Revenue)');
  document.setCell(TOTAL_ROW, 7, `=SUM(H4:H${TOTAL_LINE - 1})`);
  document.setCell(TOTAL_ROW, 8, `=COUNTIF(I4:I${TOTAL_LINE - 1},"Below")&" below"`);
  document.setCell(TOTAL_ROW, 10, `=SUM(K4:K${TOTAL_LINE - 1})`);

  salesFormats(document);

  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 10 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 10 });
  document.frozenRows = 3;
  document.frozenColumns = 1;

  salesNames(document);
  salesRules(document);

  document.setNote(2, 10, 'Found by an approximate XLOOKUP: the largest tier on the Reference sheet that the order reaches.');
  document.setNote(3 + BIGGEST, 6, 'The largest order of the year. The dashboard picks it out with SORT and FILTER.');
  document.setNote(2, 3, 'Twenty working days after the order: WORKDAY skips the weekends.');
}

// -------------------------------------------------------------------
// Reference: what everything reads
// -------------------------------------------------------------------

function reference(document: SheetDocument): void {
  setWidths(document, [100, 150, 96, 108, 120, 24, 108, 104]);

  ['Region', 'Rep', 'List price', 'Order target', 'Year target'].forEach((heading, column) =>
    document.setCell(0, column, heading)
  );
  REGIONS.forEach((region, index) => {
    const row = index + 1;
    document.setCell(row, 0, region.name);
    document.setCell(row, 1, region.rep);
    document.setCell(row, 2, region.price.toFixed(2));
    document.setCell(row, 3, String(region.target));
    document.setCell(row, 4, String(region.year));
  });

  ['Sales from', 'Commission'].forEach((heading, index) => document.setCell(0, 6 + index, heading));
  TIERS.forEach((tier, index) => {
    document.setCell(index + 1, 6, String(tier.from));
    document.setCell(index + 1, 7, String(tier.rate));
  });

  document.setCell(8, 0, 'Edit a price here and every sheet moves: Sales, the dashboard, the forecast.');

  for (let column = 0; column <= 4; column++) {
    document.setFormat(0, column, heading());
  }
  document.setFormat(0, 6, heading());
  document.setFormat(0, 7, heading());
  for (let row = 1; row <= REGIONS.length; row++) {
    document.setFormat(row, 2, cell(MONEY));
    document.setFormat(row, 3, cell(WHOLE_MONEY));
    document.setFormat(row, 4, cell(WHOLE_MONEY));
  }
  for (let row = 1; row <= TIERS.length; row++) {
    document.setFormat(row, 6, cell(WHOLE_MONEY));
    document.setFormat(row, 7, cell(SHARE));
  }
  document.setFormat(8, 0, cell(GENERAL, { italic: true, color: NOTE }));
}

// -------------------------------------------------------------------
// Dashboard: the page that opens
// -------------------------------------------------------------------

function dashboard(document: SheetDocument): void {
  setWidths(document, [96, 104, 96, 104, 104, 112, 96, 104]);

  document.setCell(0, 0, 'Northwind Trading — the year at a glance');
  document.setCell(
    1,
    0,
    'Every figure here is a formula over the Sales sheet. The months down column A are one formula that spills, the regions are SORT(UNIQUE(…)), and the five largest orders are one FILTER. Edit an order, a price on Reference, or the scenario on Forecast, and it all moves.'
  );

  // Four figures across the top, each a label over a value.
  TILES.forEach((tile, index) => {
    const column = index * 2;
    document.setCell(3, column, tile.label);
    document.setCell(4, column, tile.formula);
    document.merges.add({ firstRow: 3, lastRow: 3, firstColumn: column, lastColumn: column + 1 });
    document.merges.add({ firstRow: 4, lastRow: 4, firstColumn: column, lastColumn: column + 1 });
  });

  // Twelve months, down from one formula.
  ['Month', 'Revenue', 'Orders'].forEach((heading, column) => document.setCell(MONTHS_HEAD, column, heading));
  document.setCell(MONTHS_HEAD + 1, 0, `=TEXT(EDATE(DATE(${FIRST_MONTH[0]},${FIRST_MONTH[1]},1),SEQUENCE(12)-1),"mmm yy")`);
  for (let month = 0; month < 12; month++) {
    const row = MONTHS_HEAD + 1 + month;
    const line = row + 1;
    // The month a row stands for, as the text the orders are compared on.
    const key = `TEXT(EDATE(DATE(${FIRST_MONTH[0]},${FIRST_MONTH[1]},1),ROWS($A$${MONTHS_HEAD + 2}:A${line})-1),"yyyymm")`;
    document.setCell(row, 1, `=SUMPRODUCT((TEXT(${SALES_ORDERED},"yyyymm")=${key})*${SALES_REVENUE})`);
    document.setCell(row, 2, `=SUMPRODUCT(--(TEXT(${SALES_ORDERED},"yyyymm")=${key}))`);
  }

  // The regions, found rather than listed.
  ['Region', 'Revenue', 'Of target'].forEach((heading, index) => document.setCell(MONTHS_HEAD, 4 + index, heading));
  document.setCell(MONTHS_HEAD + 1, 4, `=SORT(UNIQUE(${SALES_REGIONS}))`);
  for (let index = 0; index < REGIONS.length; index++) {
    const row = MONTHS_HEAD + 1 + index;
    const line = row + 1;
    document.setCell(row, 5, `=SUMIF(${SALES_REGIONS},E${line},${SALES_REVENUE})`);
    document.setCell(row, 6, `=F${line}/XLOOKUP(E${line},${REGION_COLUMN},${YEAR_TARGET})`);
  }
  const regionsEnd = MONTHS_HEAD + 1 + REGIONS.length;
  document.setCell(regionsEnd + 1, 4, `=COUNTA(E${MONTHS_HEAD + 2}#)&" regions: "&TEXTJOIN(", ",TRUE,E${MONTHS_HEAD + 2}#)`);
  document.merges.add({ firstRow: regionsEnd + 1, lastRow: regionsEnd + 1, firstColumn: 4, lastColumn: 7 });

  // The five largest orders, as one block.
  document.setCell(TOP_HEAD - 1, 0, 'The five largest orders');
  document.merges.add({ firstRow: TOP_HEAD - 1, lastRow: TOP_HEAD - 1, firstColumn: 0, lastColumn: 3 });
  ['Region', 'Rep', 'Ordered', 'Due', 'Units', 'Price', 'Revenue'].forEach((heading, column) =>
    document.setCell(TOP_HEAD, column, heading)
  );
  document.setCell(
    TOP_HEAD + 1,
    0,
    `=SORT(FILTER(Sales!$A$4:$G$${TOTAL_LINE - 1},${SALES_REVENUE}>=LARGE(${SALES_REVENUE},5)),7,-1)`
  );

  dashboardFormats(document);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 7 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 7 });
  // Tall enough for the paragraph to wrap into, as a row somebody sized.
  document.rowHeights.set(1, 64);

  document.addConditional({ range: range(`B${MONTHS_HEAD + 2}:B${MONTHS_HEAD + 13}`), scale: SCALE, test: null });
  document.addConditional({
    range: range(`G${MONTHS_HEAD + 2}:G${regionsEnd}`),
    test: { kind: 'lessThan', value: 1 },
    paint: { color: BELOW, bold: true }
  });
  document.addConditional({
    range: range(`G${MONTHS_HEAD + 2}:G${regionsEnd}`),
    test: { kind: 'greaterThan', value: 0.999 },
    paint: { color: GOOD, bold: true }
  });

  const right = chartsLeft(document, 8);
  chart(document, 'line', 'Revenue by month', range(`A${MONTHS_HEAD + 1}:B${MONTHS_HEAD + 13}`), right, 70, 420, 240);
  chart(document, 'column', 'Revenue by region', range(`E${MONTHS_HEAD + 1}:F${regionsEnd}`), right, 322, 420, 230);
  // Drawn from another sheet's cells: the quarter's rollup on Summary.
  chart(
    document,
    'pie',
    'This quarter, by region',
    { start: { ...range(`A3:A3`).start, sheet: 'Summary' }, end: { ...range(`D${ROLLUP_LINE - 1}:D${ROLLUP_LINE - 1}`).start, sheet: 'Summary' } },
    right,
    564,
    420,
    260
  );

  document.setNote(4, 0, `The whole year: all ${ORDERS.length} orders on the Sales sheet.`);
  document.setNote(MONTHS_HEAD + 1, 0, 'One formula, in A9 alone: TEXT over EDATE over SEQUENCE(12), spilling twelve rows.');
  document.setNote(TOP_HEAD + 1, 0, 'One formula: SORT(FILTER(...)) over the Sales sheet, spilling five rows and seven columns.');
}

// -------------------------------------------------------------------
// Forecast: the what-if
// -------------------------------------------------------------------

function forecast(document: SheetDocument): void {
  setWidths(document, [120, 112, 124, 24, 150, 120]);

  document.setCell(0, 0, 'Next year, three ways');

  document.setCell(2, 0, 'Scenario');
  document.setCell(2, 1, 'Base');
  document.setCell(3, 0, 'Growth a month');
  document.setCell(3, 1, '=XLOOKUP(B3,A7:A9,B7:B9)');
  document.setCell(4, 0, 'From an average of');
  document.setCell(4, 1, `=AVERAGE(Dashboard!B${MONTHS_HEAD + 2}:B${MONTHS_HEAD + 13})`);

  ['Scenario', 'Growth a month'].forEach((heading, column) => document.setCell(5, column, heading));
  SCENARIOS.forEach((scenario, index) => {
    document.setCell(6 + index, 0, scenario.name);
    document.setCell(6 + index, 1, String(scenario.growth));
  });

  ['Month', 'Revenue', 'Running total'].forEach((heading, column) => document.setCell(FORECAST_HEAD, column, heading));
  document.setCell(FORECAST_HEAD + 1, 0, '=TEXT(EDATE(DATE(2026,10,1),SEQUENCE(12)-1),"mmm yy")');
  // Twelve months from one formula: ROUND runs across the array the
  // power makes, and the whole column spills from B12.
  document.setCell(FORECAST_HEAD + 1, 1, '=ROUND($B$5*(1+$B$4)^SEQUENCE(12),0)');
  for (let month = 0; month < 12; month++) {
    const line = FORECAST_HEAD + 2 + month;
    document.setCell(FORECAST_HEAD + 1 + month, 2, `=SUM($B$${FORECAST_HEAD + 2}:B${line})`);
  }
  document.setCell(FORECAST_HEAD + 14, 0, 'Next year');
  document.setCell(FORECAST_HEAD + 14, 1, `=SUM(B${FORECAST_HEAD + 2}#)`);
  document.setCell(FORECAST_HEAD + 15, 0, 'Against this year');
  document.setCell(FORECAST_HEAD + 15, 1, `=B${FORECAST_HEAD + 15}/SUM(Dashboard!B${MONTHS_HEAD + 2}:B${MONTHS_HEAD + 13})-1`);

  document.setCell(2, 4, 'Warehouse loan');
  document.setCell(3, 4, 'Borrowed');
  document.setCell(3, 5, '250000');
  document.setCell(4, 4, 'A year');
  document.setCell(4, 5, '0.061');
  document.setCell(5, 4, 'Years');
  document.setCell(5, 5, '5');
  document.setCell(6, 4, 'A month, repaid');
  document.setCell(6, 5, '=PMT(F5/12,F6*12,-F4)');
  document.setCell(7, 4, 'Interest, all told');
  document.setCell(7, 5, '=F7*F6*12-F4');
  document.setCell(8, 4, 'Of a month’s revenue');
  document.setCell(8, 5, `=F7/B5`);

  forecastFormats(document);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 5 });

  document.addValidation({
    range: range('B3:B3'),
    rule: { kind: 'list', values: SCENARIOS.map(scenario => scenario.name) },
    strict: true,
    message: 'A scenario is Low, Base or High.'
  });
  document.addConditional({ range: range(`B${FORECAST_HEAD + 2}:B${FORECAST_HEAD + 13}`), scale: SCALE, test: null });
  document.addConditional({
    range: range(`B${FORECAST_HEAD + 16}:B${FORECAST_HEAD + 16}`),
    test: { kind: 'greaterThan', value: 0 },
    paint: { color: GOOD, bold: true }
  });
  document.setNote(2, 1, 'Pick Low, Base or High from the list: the growth, the twelve months and the chart all follow.');
  document.setNote(FORECAST_HEAD + 1, 1, 'One formula: ROUND(average × (1 + growth) ^ SEQUENCE(12)), spilling twelve months.');

  chart(document, 'area', 'Next year, month by month', range(`A${FORECAST_HEAD + 1}:B${FORECAST_HEAD + 13}`), chartsLeft(document, 6), 70, 440, 280);
}

// -------------------------------------------------------------------
// Summary: the quarter, and the library
// -------------------------------------------------------------------

function summary(document: SheetDocument): void {
  setWidths(document, [240, 92, 104, 124, 92]);

  document.setCell(0, 0, 'The quarter, July to September');
  ['Region', 'Orders', 'Units', 'Revenue', 'Share'].forEach((heading, column) => document.setCell(2, column, heading));
  REGIONS.forEach((region, index) => {
    const row = index + 3;
    const line = row + 1;
    document.setCell(row, 0, region.name);
    document.setCell(row, 1, `=COUNTIFS(${SALES_REGIONS},$A${line},${SALES_ORDERED},">="&${QUARTER_START})`);
    document.setCell(row, 2, `=SUMIFS(${SALES_UNITS},${SALES_REGIONS},$A${line},${SALES_ORDERED},">="&${QUARTER_START})`);
    document.setCell(row, 3, `=SUMIFS(${SALES_REVENUE},${SALES_REGIONS},$A${line},${SALES_ORDERED},">="&${QUARTER_START})`);
    document.setCell(row, 4, `=IFERROR(D${line}/$D$${ROLLUP_LINE},0)`);
  });
  document.setCell(ROLLUP_ROW, 0, 'Total');
  for (let column = 1; column <= 4; column++) {
    const letter = ['A', 'B', 'C', 'D', 'E'][column];
    document.setCell(ROLLUP_ROW, column, `=SUM(${letter}4:${letter}${ROLLUP_LINE - 1})`);
  }

  document.setCell(GLANCES_HEAD, 0, 'The library, at work');
  GLANCES.forEach((glance, index) => {
    document.setCell(GLANCES_HEAD + 1 + index, 0, glance.label);
    document.setCell(GLANCES_HEAD + 1 + index, 1, glance.formula);
  });

  summaryFormats(document);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 });
  document.addConditional({ range: range(`D4:D${ROLLUP_LINE - 1}`), scale: SCALE, test: null });
}

// -------------------------------------------------------------------
// What the sheets are made of
// -------------------------------------------------------------------

interface Region {
  readonly name: string;
  readonly rep: string;
  readonly price: number;
  /** What one order should reach. */
  readonly target: number;
  /** What the year should come to. */
  readonly year: number;
}

const REGIONS: readonly Region[] = [
  { name: 'North', rep: 'Dana Okonjo', price: 48, target: 12_000, year: 90_000 },
  { name: 'South', rep: 'Priya Raman', price: 52.5, target: 15_000, year: 80_000 },
  { name: 'East', rep: 'Tomas Lindqvist', price: 44.75, target: 9_000, year: 70_000 },
  { name: 'West', rep: 'Aiko Tanaka', price: 61.25, target: 18_000, year: 95_000 },
  { name: 'Central', rep: 'Marcus Bell', price: 39.9, target: 8_000, year: 55_000 }
];

/** Commission by the size of the order: the rate of the largest threshold it reaches. */
const TIERS: readonly { from: number; rate: number }[] = [
  { from: 0, rate: 0.02 },
  { from: 5_000, rate: 0.03 },
  { from: 15_000, rate: 0.045 },
  { from: 30_000, rate: 0.06 }
];

const SCENARIOS: readonly { name: string; growth: number }[] = [
  { name: 'Low', growth: 0.005 },
  { name: 'Base', growth: 0.02 },
  { name: 'High', growth: 0.04 }
];

interface Order {
  readonly region: string;
  /** Year, month, day, as `DATE` takes them. */
  readonly ordered: readonly [number, number, number];
  readonly units: number;
  readonly review: string;
}

/**
 * Two orders a month, October to September.
 *
 * The unit counts are picked so that some orders miss their region's
 * target and some beat it, so that the months rise and dip for the
 * line chart and the colour scale, and so that every commission tier
 * is reached by at least one order.
 */
const ORDERS: readonly Order[] = [
  { region: 'North', ordered: [2025, 10, 6], units: 180, review: 'Approved' },
  { region: 'South', ordered: [2025, 10, 21], units: 260, review: 'Approved' },
  { region: 'East', ordered: [2025, 11, 4], units: 95, review: 'Approved' },
  { region: 'West', ordered: [2025, 11, 18], units: 340, review: 'Approved' },
  { region: 'Central', ordered: [2025, 12, 2], units: 410, review: 'Approved' },
  { region: 'North', ordered: [2025, 12, 15], units: 520, review: 'Approved' },
  { region: 'South', ordered: [2026, 1, 12], units: 150, review: 'Approved' },
  { region: 'East', ordered: [2026, 1, 27], units: 380, review: 'Held' },
  { region: 'West', ordered: [2026, 2, 9], units: 210, review: 'Approved' },
  { region: 'Central', ordered: [2026, 2, 23], units: 175, review: 'Approved' },
  { region: 'North', ordered: [2026, 3, 9], units: 290, review: 'Approved' },
  { region: 'South', ordered: [2026, 3, 24], units: 610, review: 'Approved' },
  { region: 'East', ordered: [2026, 4, 7], units: 230, review: 'Approved' },
  { region: 'West', ordered: [2026, 4, 21], units: 720, review: 'Approved' },
  { region: 'Central', ordered: [2026, 5, 5], units: 330, review: 'Pending' },
  { region: 'North', ordered: [2026, 5, 19], units: 145, review: 'Approved' },
  { region: 'South', ordered: [2026, 6, 2], units: 305, review: 'Approved' },
  { region: 'East', ordered: [2026, 6, 16], units: 470, review: 'Approved' },
  { region: 'West', ordered: [2026, 7, 7], units: 260, review: 'Held' },
  { region: 'Central', ordered: [2026, 7, 21], units: 512, review: 'Approved' },
  { region: 'North', ordered: [2026, 8, 4], units: 340, review: 'Pending' },
  { region: 'South', ordered: [2026, 8, 18], units: 78, review: 'Approved' },
  { region: 'East', ordered: [2026, 9, 1], units: 610, review: 'Held' },
  { region: 'West', ordered: [2026, 9, 15], units: 205, review: 'Pending' }
];

/** The order with the largest revenue, for its note. */
const BIGGEST = ORDERS.reduce(
  (best, order, index) => {
    const price = REGIONS.find(region => region.name === order.region)?.price ?? 0;
    return order.units * price > best.value ? { index, value: order.units * price } : best;
  },
  { index: 0, value: 0 }
).index;

/** The row the totals sit on, and its one-based twin for the text. */
const TOTAL_ROW = ORDERS.length + 3;
const TOTAL_LINE = TOTAL_ROW + 1;

const ROLLUP_ROW = REGIONS.length + 3;
const ROLLUP_LINE = ROLLUP_ROW + 1;
const GLANCES_HEAD = ROLLUP_ROW + 3;

/** The dashboard's layout: the month table's heading, and the top-orders block's. */
const MONTHS_HEAD = 7;
const TOP_HEAD = MONTHS_HEAD + 16;
const FIRST_MONTH = [2025, 10] as const;
const FORECAST_HEAD = 10;

const QUARTER_START = 'DATE(2026,7,1)';

/** The lookup blocks on `Reference`, absolute so a fill cannot move them. */
const PRICES = 'Reference!$A$2:$D$6';
const REGION_COLUMN = 'Reference!$A$2:$A$6';
const PRICE_COLUMN = 'Reference!$C$2:$C$6';
const YEAR_TARGET = 'Reference!$E$2:$E$6';
const TIER_FROM = 'Reference!$G$2:$G$5';
const TIER_RATE = 'Reference!$H$2:$H$5';

const SALES_REGIONS = `Sales!$A$4:$A$${TOTAL_LINE - 1}`;
const SALES_ORDERED = `Sales!$C$4:$C$${TOTAL_LINE - 1}`;
const SALES_UNITS = `Sales!$E$4:$E$${TOTAL_LINE - 1}`;
const SALES_REVENUE = `Sales!$G$4:$G$${TOTAL_LINE - 1}`;

/** The dashboard's four figures. */
const TILES: readonly { label: string; formula: string; number: NumberFormat }[] = [
  { label: 'Revenue, the year', formula: `=SUM(${SALES_REVENUE})`, number: { kind: 'currency', places: 0, symbol: '$' } },
  { label: 'Orders', formula: `=COUNT(${SALES_REVENUE})`, number: { kind: 'number', places: 0, thousands: true } },
  { label: 'Average order', formula: `=AVERAGE(${SALES_REVENUE})`, number: { kind: 'currency', places: 0, symbol: '$' } },
  {
    label: 'Best month',
    formula: `=INDEX(A${MONTHS_HEAD + 2}#,MATCH(MAX(B${MONTHS_HEAD + 2}:B${MONTHS_HEAD + 13}),B${MONTHS_HEAD + 2}:B${MONTHS_HEAD + 13},0))`,
    number: GENERAL
  }
];

/**
 * A line each from most of the function families, each a question
 * somebody would actually ask — and the last one asks a question with
 * no answer, so `IFERROR` has something to catch in front of a person.
 */
const GLANCES: readonly { label: string; formula: string; number: NumberFormat }[] = [
  { label: 'Largest order', formula: `=LARGE(${SALES_REVENUE},1)`, number: { kind: 'currency', places: 2, symbol: '$' } },
  { label: 'Median order', formula: `=MEDIAN(${SALES_REVENUE})`, number: { kind: 'currency', places: 2, symbol: '$' } },
  { label: 'Spread, one deviation', formula: `=STDEV(${SALES_REVENUE})`, number: { kind: 'currency', places: 2, symbol: '$' } },
  {
    label: 'Region that took it',
    formula: `=INDEX(${SALES_REGIONS},MATCH(LARGE(${SALES_REVENUE},1),${SALES_REVENUE},0))`,
    number: GENERAL
  },
  { label: 'Orders over 20,000', formula: `=COUNTIF(${SALES_REVENUE},">20000")`, number: { kind: 'number', places: 0, thousands: true } },
  { label: 'Working days in the quarter', formula: '=NETWORKDAYS(DATE(2026,7,1),DATE(2026,9,30))', number: GENERAL },
  { label: 'First order to last', formula: `=DATEDIF(MIN(${SALES_ORDERED}),MAX(${SALES_ORDERED}),"m")&" months"`, number: GENERAL },
  { label: 'Last order closes', formula: `=TEXT(EOMONTH(MAX(${SALES_ORDERED}),0),"dddd d mmmm yyyy")`, number: GENERAL },
  { label: 'Reps who held an order', formula: `=TEXTJOIN(", ",TRUE,UNIQUE(FILTER(Sales!$B$4:$B$${TOTAL_LINE - 1},Sales!$J$4:$J$${TOTAL_LINE - 1}="Held")))`, number: GENERAL },
  {
    label: 'How the year went',
    formula: `=IFS(SUM(${SALES_REVENUE})>=SUM(${YEAR_TARGET}),"Ahead of target",SUM(${SALES_REVENUE})>=0.9*SUM(${YEAR_TARGET}),"Close to target",TRUE,"Behind target")`,
    number: GENERAL
  },
  { label: 'Quarter of the last order', formula: `=SWITCH(ROUNDUP(MONTH(MAX(${SALES_ORDERED}))/3,0),1,"First",2,"Second",3,"Third","Fourth")&" quarter"`, number: GENERAL },
  {
    label: 'Largest order, against the rest',
    formula: `=LET(top,LARGE(${SALES_REVENUE},1),rest,SUM(${SALES_REVENUE})-top,top/rest)`,
    number: { kind: 'percent', places: 1 }
  },
  {
    label: 'Commission, worked out again by MAP',
    formula: `=SUM(MAP(${SALES_REVENUE},LAMBDA(amount,Commission(amount))))`,
    number: { kind: 'currency', places: 2, symbol: '$' }
  },
  { label: 'A region not on file', formula: `=IFERROR(VLOOKUP("Nowhere",${PRICES},2,FALSE),"not on file")`, number: GENERAL }
];

// -------------------------------------------------------------------
// Paint
// -------------------------------------------------------------------

/**
 * The tab colours, as `SheetTabs` names them.
 *
 * Copied rather than imported: that file is a component in the render
 * worker, and this one runs in the application worker.
 */
const TAB_BLUE = '#4285f4';
const TAB_ORANGE = '#fa7b17';
const TAB_GREEN = '#34a853';
const TAB_PURPLE = '#a142f4';
const TAB_RED = '#ea4335';

const BANNER_FILL = '#1f3a5f';
const HEADING_FILL = '#eef2f7';
const TILE_FILL = '#f4f7fb';
const NOTE = '#5a6b7d';
const BELOW = '#b3261e';
const GOOD = '#1e7b34';
const HELD_FILL = '#ffe8cc';
const INPUT_FILL = '#fff8d6';

const MONEY: NumberFormat = { kind: 'currency', places: 2, symbol: '$' };
const WHOLE_MONEY: NumberFormat = { kind: 'currency', places: 0, symbol: '$' };
const COUNTED: NumberFormat = { kind: 'number', places: 0, thousands: true };
const SHARE: NumberFormat = { kind: 'percent', places: 1 };
const DAY: NumberFormat = { kind: 'date', pattern: 'dmy' };

const RULED_BELOW: CellBorders = { ...NO_BORDERS, bottom: { width: 2, color: '' } };
const RULED_ABOVE: CellBorders = { ...NO_BORDERS, top: { width: 2, color: '' } };
const BOXED: CellBorders = {
  top: { width: 1, color: '#c9d3df' },
  right: { width: 1, color: '#c9d3df' },
  bottom: { width: 1, color: '#c9d3df' },
  left: { width: 1, color: '#c9d3df' }
};

/** A three-colour scale, pale enough to read black text through. */
const SCALE = { from: '#fde2e2', middle: '#fff4cc', to: '#d9efdc' };

function cell(number: NumberFormat, over: Partial<CellPaint> = {}): CellFormat {
  return { number, paint: { ...PLAIN, ...over } };
}

function heading(): CellFormat {
  return cell(GENERAL, { bold: true, fill: HEADING_FILL, align: 'center', borders: RULED_BELOW });
}

function banner(size: number): CellFormat {
  return cell(GENERAL, { bold: true, fontSize: size, color: '#ffffff', fill: BANNER_FILL, align: 'center' });
}

function salesFormats(document: SheetDocument): void {
  document.setFormat(0, 0, banner(17));
  document.setFormat(1, 0, cell(GENERAL, { italic: true, color: NOTE, align: 'center' }));
  for (let column = 0; column <= 10; column++) {
    document.setFormat(2, column, heading());
  }
  for (let row = 3; row < TOTAL_ROW; row++) {
    document.setFormat(row, 2, cell(DAY));
    document.setFormat(row, 3, cell(DAY, { color: NOTE }));
    document.setFormat(row, 4, cell(COUNTED));
    document.setFormat(row, 5, cell(MONEY));
    document.setFormat(row, 6, cell(MONEY));
    document.setFormat(row, 7, cell(SHARE));
    document.setFormat(row, 8, cell(GENERAL, { align: 'center' }));
    document.setFormat(row, 9, cell(GENERAL, { align: 'center' }));
    document.setFormat(row, 10, cell(MONEY, { color: NOTE }));
  }
  const totals: readonly (NumberFormat | null)[] = [GENERAL, null, null, null, COUNTED, MONEY, MONEY, SHARE, GENERAL, null, MONEY];
  totals.forEach((number, column) => {
    if (number !== null) {
      document.setFormat(TOTAL_ROW, column, cell(number, { bold: true, borders: RULED_ABOVE }));
    }
  });
}

function dashboardFormats(document: SheetDocument): void {
  document.setFormat(0, 0, banner(18));
  document.setFormat(1, 0, cell(GENERAL, { italic: true, color: NOTE, wrap: true }));
  TILES.forEach((tile, index) => {
    const column = index * 2;
    document.setFormat(3, column, cell(GENERAL, { color: NOTE, fill: TILE_FILL, align: 'center', borders: BOXED }));
    document.setFormat(4, column, cell(tile.number, { bold: true, fontSize: 18, fill: TILE_FILL, align: 'center', borders: BOXED }));
  });
  for (let column = 0; column <= 2; column++) {
    document.setFormat(MONTHS_HEAD, column, heading());
  }
  for (let column = 4; column <= 6; column++) {
    document.setFormat(MONTHS_HEAD, column, heading());
  }
  for (let month = 0; month < 12; month++) {
    const row = MONTHS_HEAD + 1 + month;
    document.setFormat(row, 0, cell(GENERAL, { align: 'start' }));
    document.setFormat(row, 1, cell(WHOLE_MONEY));
    document.setFormat(row, 2, cell(COUNTED, { align: 'center' }));
  }
  for (let index = 0; index < REGIONS.length; index++) {
    const row = MONTHS_HEAD + 1 + index;
    document.setFormat(row, 5, cell(WHOLE_MONEY));
    document.setFormat(row, 6, cell(SHARE));
  }
  document.setFormat(MONTHS_HEAD + 2 + REGIONS.length, 4, cell(GENERAL, { italic: true, color: NOTE }));
  document.setFormat(TOP_HEAD - 1, 0, cell(GENERAL, { bold: true, fontSize: 14 }));
  for (let column = 0; column <= 6; column++) {
    document.setFormat(TOP_HEAD, column, heading());
  }
  for (let row = TOP_HEAD + 1; row <= TOP_HEAD + 5; row++) {
    document.setFormat(row, 2, cell(DAY));
    document.setFormat(row, 3, cell(DAY, { color: NOTE }));
    document.setFormat(row, 4, cell(COUNTED));
    document.setFormat(row, 5, cell(MONEY));
    document.setFormat(row, 6, cell(MONEY, { bold: true }));
  }
}

function forecastFormats(document: SheetDocument): void {
  document.setFormat(0, 0, banner(17));
  for (const row of [2, 3, 4]) {
    document.setFormat(row, 0, cell(GENERAL, { bold: true }));
  }
  document.setFormat(2, 1, cell(GENERAL, { fill: INPUT_FILL, bold: true, align: 'center', borders: BOXED }));
  document.setFormat(3, 1, cell({ kind: 'percent', places: 1 }));
  document.setFormat(4, 1, cell(WHOLE_MONEY));
  document.setFormat(5, 0, heading());
  document.setFormat(5, 1, heading());
  for (let index = 0; index < SCENARIOS.length; index++) {
    document.setFormat(6 + index, 1, cell({ kind: 'percent', places: 1 }));
  }
  for (let column = 0; column <= 2; column++) {
    document.setFormat(FORECAST_HEAD, column, heading());
  }
  for (let month = 0; month < 12; month++) {
    document.setFormat(FORECAST_HEAD + 1 + month, 1, cell(WHOLE_MONEY));
    document.setFormat(FORECAST_HEAD + 1 + month, 2, cell(WHOLE_MONEY, { color: NOTE }));
  }
  document.setFormat(FORECAST_HEAD + 14, 0, cell(GENERAL, { bold: true, borders: RULED_ABOVE }));
  document.setFormat(FORECAST_HEAD + 14, 1, cell(WHOLE_MONEY, { bold: true, borders: RULED_ABOVE }));
  document.setFormat(FORECAST_HEAD + 15, 0, cell(GENERAL, { color: NOTE }));
  document.setFormat(FORECAST_HEAD + 15, 1, cell({ kind: 'percent', places: 1 }));
  document.setFormat(2, 4, cell(GENERAL, { bold: true, borders: RULED_BELOW }));
  document.setFormat(2, 5, cell(GENERAL, { borders: RULED_BELOW }));
  document.setFormat(3, 5, cell(WHOLE_MONEY, { fill: INPUT_FILL }));
  document.setFormat(4, 5, cell({ kind: 'percent', places: 1 }, { fill: INPUT_FILL }));
  document.setFormat(5, 5, cell(GENERAL, { fill: INPUT_FILL }));
  document.setFormat(6, 5, cell(MONEY, { bold: true }));
  document.setFormat(7, 5, cell(MONEY));
  document.setFormat(8, 5, cell(SHARE));
}

function summaryFormats(document: SheetDocument): void {
  document.setFormat(0, 0, banner(15));
  for (let column = 0; column <= 4; column++) {
    document.setFormat(2, column, heading());
  }
  for (let row = 3; row < ROLLUP_ROW; row++) {
    document.setFormat(row, 1, cell(COUNTED));
    document.setFormat(row, 2, cell(COUNTED));
    document.setFormat(row, 3, cell(MONEY));
    document.setFormat(row, 4, cell(SHARE));
  }
  const rollup: readonly NumberFormat[] = [GENERAL, COUNTED, COUNTED, MONEY, SHARE];
  rollup.forEach((number, column) =>
    document.setFormat(ROLLUP_ROW, column, cell(number, { bold: true, borders: RULED_ABOVE }))
  );
  document.setFormat(GLANCES_HEAD, 0, cell(GENERAL, { bold: true, borders: RULED_BELOW }));
  GLANCES.forEach((glance, index) => {
    document.setFormat(GLANCES_HEAD + 1 + index, 0, cell(GENERAL, { color: NOTE }));
    document.setFormat(GLANCES_HEAD + 1 + index, 1, cell(glance.number));
  });
}

// -------------------------------------------------------------------
// Names, rules, and the things that refuse
// -------------------------------------------------------------------

function salesNames(document: SheetDocument): void {
  const last = TOTAL_LINE - 1;
  document.defineName('Region', range(`A4:A${last}`));
  document.defineName('Ordered', range(`C4:C${last}`));
  document.defineName('Units', range(`E4:E${last}`));
  document.defineName('Price', range(`F4:F${last}`));
  document.defineName('Revenue', range(`G4:G${last}`));
  // A name that holds a function: the commission an amount earns, by the
  // tier it reaches. Every row of column K calls it, and so does a
  // figure on Summary, handing it to MAP.
  document.defineFormulaName(
    'Commission',
    `=LAMBDA(amount, ROUND(amount * XLOOKUP(amount, ${TIER_FROM}, ${TIER_RATE}, 0, -1), 2))`
  );
}

function salesRules(document: SheetDocument): void {
  const last = TOTAL_LINE - 1;

  document.addConditional({ range: range(`G4:G${last}`), scale: SCALE, test: null });
  document.addConditional({
    range: range(`I4:I${last}`),
    test: { kind: 'textContains', text: 'Below' },
    paint: { color: BELOW, bold: true }
  });
  document.addConditional({
    range: range(`J4:J${last}`),
    test: { kind: 'equalTo', value: 'Held' },
    paint: { fill: HELD_FILL }
  });
  // A rule that is a formula: a large order still waiting on a review.
  document.addConditional({
    range: range(`E4:E${last}`),
    test: { kind: 'formula', input: '=AND(E4>=300,J4<>"Approved")' },
    paint: { bold: true, color: BELOW }
  });

  /**
   * The strict one and the lenient one, side by side and on purpose: a
   * review is one of three words, and typing a fourth is a mistake; a
   * unit count out of the usual range might be a typo or a very good
   * month, so it is marked and let stand.
   */
  document.addValidation({
    range: range(`J4:J${last}`),
    rule: { kind: 'list', values: ['Approved', 'Pending', 'Held'] },
    strict: true,
    message: 'A review is Approved, Pending or Held.'
  });
  document.addValidation({
    range: range(`E4:E${last}`),
    rule: { kind: 'number', min: 1, max: 5_000, integer: true },
    message: 'Units are whole numbers, 1 to 5,000.'
  });
}

// -------------------------------------------------------------------
// Charts
// -------------------------------------------------------------------

/** Where the charts start: past the columns the sheet uses, and a little more. */
function chartsLeft(document: SheetDocument, columns: number): number {
  const used = document.columnWidths.slice(0, columns).reduce((sum, width) => sum + width, 0);
  return GUTTER_WIDTH + used + 24;
}

function chart(
  document: SheetDocument,
  kind: ChartKind,
  title: string,
  over: RangeRef,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  document.addChart({ kind, title, range: over, place: { x, y: y + HEADER_HEIGHT - ROW_HEIGHT, width, height }, legend: kind === 'pie' });
}

// -------------------------------------------------------------------
// Small helpers
// -------------------------------------------------------------------

/**
 * The widths this sheet wants, over whatever the document already had.
 * A sheet added here starts with none, and a column with no width is
 * drawn at nothing — so the tail is filled in rather than left as holes,
 * out to every column the sheet has, as a sheet added from the tab strip
 * is. The window's width is the sum of these, and a chart past the last
 * one would be laid out in a row too narrow to hold it, at no width.
 */
function setWidths(document: SheetDocument, wanted: readonly number[]): void {
  const all = [...document.columnWidths];
  const length = Math.max(all.length, wanted.length, document.book.extent?.columns ?? 0);
  for (let column = 0; column < length; column++) {
    all[column] = wanted[column] ?? all[column] ?? COLUMN_WIDTH;
  }
  document.columnWidths = all;
}

function range(address: string): RangeRef {
  const parsed = parseAddress(address);
  if (parsed === null) {
    throw new Error(`The seed cannot parse the address ${address}.`);
  }
  return parsed;
}
