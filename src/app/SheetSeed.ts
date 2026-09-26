import { parseAddress, type RangeRef } from '../sheet/A1';
import {
  GENERAL,
  NO_BORDERS,
  PLAIN,
  type CellBorders,
  type CellFormat,
  type CellPaint,
  type NumberFormat
} from '../sheet/Format';
import { COLUMN_WIDTH } from './dimensions';
import type { SheetDocument } from './SheetDocument';

/**
 * Something to open onto.
 *
 * An empty grid demonstrates nothing, and after fourteen phases a
 * five-row table demonstrates about as little. This is a small
 * workbook with a shape a person would recognise — orders on one
 * sheet, the prices and targets they are looked up against on
 * another, and a rollup that reads across to both — chosen so that
 * every part of the application somebody can see is doing something
 * the moment the page finishes loading.
 *
 * What is on screen, and which phase put it there:
 *
 * - **Formulas that read formulas** (Phase 1). Units times price is
 *   revenue, revenue over the total is a share, and the share column
 *   goes on adding to a hundred after any of it is edited.
 * - **A merged banner and a frozen corner** (Phase 10). Three rows
 *   and the region column stay put while the rest scrolls.
 * - **Number and paint formats** (Phase 9). Currency, percent,
 *   thousands, dates, and per-edge borders under the header and over
 *   the total.
 * - **The function library** (Phase 11). `VLOOKUP` across sheets,
 *   `SUMIF`, `COUNTIF`, `MEDIAN`, `STDEV`, `LARGE`, `INDEX`/`MATCH`,
 *   `NETWORKDAYS`, `EOMONTH`, `TEXT` and an `IFERROR` that catches a
 *   lookup written to fail.
 * - **Named ranges** (Phase 12). `=SUM(Revenue)` rather than
 *   `=SUM(G4:G13)`, which is the difference the name box exists for.
 * - **Many sheets** (Phase 13). Three of them, coloured, with
 *   formulas reading across and recalculating in an order that
 *   crosses sheets.
 * - **Formats that think, and rules about what a cell may hold**
 *   (Phase 14). A three-colour scale down the revenue column, red
 *   bold on anything under target, and a dropdown on the review
 *   column that refuses a word it does not know.
 *
 * Names are resolved against the sheet the formula is written on, so
 * the ones defined here are used only on `Sales` and every formula
 * that reaches another sheet writes the sheet out. That is not a
 * limitation being worked around; it is what makes `=SUM(Units)`
 * mean this sheet's units on a duplicate of it.
 *
 * It is not a fixture: specs build their own sheets. Phase 6 replaces
 * it with whatever the repository loaded, so this runs exactly once
 * in the life of a browser profile.
 */
export function seed(document: SheetDocument): void {
  // Both other sheets exist before `Sales` is written, because a
  // formula reading `Reference!A2` can only be parsed into something
  // that points anywhere if `Reference` is there to point at.
  document.renameSheet(0, 'Sales');
  document.setSheetColour(0, TAB_BLUE);

  document.setSheetColour(document.addSheet('Reference'), TAB_ORANGE);
  reference(document);

  document.setSheetColour(document.addSheet('Summary'), TAB_GREEN);
  summary(document);

  document.activate(0);
  sales(document);

  document.setSelection(3, 4, 3, 4);
}

// -------------------------------------------------------------------
// The sheets
// -------------------------------------------------------------------

/** The orders, and the sheet somebody actually looks at. */
function sales(document: SheetDocument): void {
  setWidths(document, [108, 128, 104, 104, 80, 92, 112, 80, 96, 96]);

  document.setCell(0, 0, 'Northwind Trading');
  document.setCell(
    1,
    0,
    '="Ten orders, booked to "&TEXT(MAX(Ordered),"D MMM YYYY")&", priced from the Reference sheet"'
  );

  const headings = ['Region', 'Rep', 'Ordered', 'Due', 'Units', 'Price', 'Revenue', 'Share', 'Status', 'Review'];
  headings.forEach((heading, column) => document.setCell(2, column, heading));

  ORDERS.forEach((order, index) => {
    const row = index + 3;
    const line = row + 1;
    document.setCell(row, 0, order.region);
    // Every lookup goes through the same absolute block, so a rep, a
    // price and a target are three columns of one table rather than
    // three tables — and dragging any of these sideways still works.
    document.setCell(row, 1, `=VLOOKUP($A${line},${PRICES},2,FALSE)`);
    document.setCell(row, 2, `=DATE(${order.ordered[0]},${order.ordered[1]},${order.ordered[2]})`);
    // Month end from the order date, which is what "due" means here
    // and is one function rather than a calendar somebody maintains.
    document.setCell(row, 3, `=EOMONTH(C${line},0)`);
    document.setCell(row, 4, String(order.units));
    document.setCell(row, 5, `=VLOOKUP($A${line},${PRICES},3,FALSE)`);
    document.setCell(row, 6, `=E${line}*F${line}`);
    // Absolute on the total, relative on the row: the pair Phase 5's
    // fill handle has to tell apart when it extends this downwards.
    document.setCell(row, 7, `=G${line}/$G$${TOTAL_LINE}`);
    document.setCell(row, 8, `=IF(G${line}>=VLOOKUP($A${line},${PRICES},4,FALSE),"On target","Below")`);
    document.setCell(row, 9, order.review);
  });

  const total = TOTAL_ROW;
  document.setCell(total, 0, 'Total');
  document.setCell(total, 4, '=SUM(Units)');
  document.setCell(total, 5, '=AVERAGE(Price)');
  document.setCell(total, 6, '=SUM(Revenue)');
  document.setCell(total, 7, `=SUM(H4:H${TOTAL_LINE - 1})`);
  document.setCell(total, 8, `=COUNTIF(I4:I${TOTAL_LINE - 1},"Below")&" below"`);

  salesFormats(document);

  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 9 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 9 });

  // The banner, the subtitle and the headings stay; so does the
  // region, because a wide table read sideways with the row's own
  // label gone is a table nobody can read.
  document.frozenRows = 3;
  document.frozenColumns = 1;

  salesNames(document);
  salesRules(document);
}

/** The prices and targets every lookup on `Sales` goes through. */
function reference(document: SheetDocument): void {
  setWidths(document, [108, 168, 112, 112]);

  ['Region', 'Rep', 'List price', 'Target'].forEach((heading, column) => document.setCell(0, column, heading));

  REGIONS.forEach((region, index) => {
    const row = index + 1;
    document.setCell(row, 0, region.name);
    document.setCell(row, 1, region.rep);
    document.setCell(row, 2, region.price.toFixed(2));
    document.setCell(row, 3, String(region.target));
  });

  document.setCell(7, 0, 'Edit a price here and the Sales sheet moves.');

  for (let column = 0; column <= 3; column++) {
    document.setFormat(0, column, heading());
  }
  for (let row = 1; row <= REGIONS.length; row++) {
    document.setFormat(row, 2, cell(MONEY));
    document.setFormat(row, 3, cell(WHOLE_MONEY));
  }
  document.setFormat(7, 0, cell(GENERAL, { italic: true, color: NOTE }));
}

/** What the quarter came to, read across from the other two sheets. */
function summary(document: SheetDocument): void {
  setWidths(document, [216, 92, 104, 124, 92]);

  document.setCell(0, 0, 'Quarter at a glance');

  ['Region', 'Orders', 'Units', 'Revenue', 'Share'].forEach((heading, column) =>
    document.setCell(2, column, heading)
  );

  REGIONS.forEach((region, index) => {
    const row = index + 3;
    const line = row + 1;
    document.setCell(row, 0, region.name);
    document.setCell(row, 1, `=COUNTIF(${SALES_REGIONS},$A${line})`);
    document.setCell(row, 2, `=SUMIF(${SALES_REGIONS},$A${line},${SALES_UNITS})`);
    document.setCell(row, 3, `=SUMIF(${SALES_REGIONS},$A${line},${SALES_REVENUE})`);
    document.setCell(row, 4, `=D${line}/$D$${ROLLUP_LINE}`);
  });

  const rollup = ROLLUP_ROW;
  document.setCell(rollup, 0, 'Total');
  for (let column = 1; column <= 4; column++) {
    const letter = ['A', 'B', 'C', 'D', 'E'][column];
    document.setCell(rollup, column, `=SUM(${letter}4:${letter}${ROLLUP_LINE - 1})`);
  }

  document.setCell(10, 0, 'The library, at work');
  GLANCES.forEach((glance, index) => {
    const row = index + 11;
    document.setCell(row, 0, glance.label);
    document.setCell(row, 1, glance.formula);
  });

  summaryFormats(document);

  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 });

  // A scale on a sheet that is not the one in front of somebody, so
  // that the rules are visibly *per sheet* rather than per workbook.
  document.addConditional({ range: range(`D4:D${ROLLUP_LINE - 1}`), scale: SCALE, test: null });
}

// -------------------------------------------------------------------
// What the sheets are made of
// -------------------------------------------------------------------

interface Region {
  readonly name: string;
  readonly rep: string;
  readonly price: number;
  readonly target: number;
}

const REGIONS: readonly Region[] = [
  { name: 'North', rep: 'Dana Okonjo', price: 48, target: 12_000 },
  { name: 'South', rep: 'Priya Raman', price: 52.5, target: 15_000 },
  { name: 'East', rep: 'Tomas Lindqvist', price: 44.75, target: 9_000 },
  { name: 'West', rep: 'Aiko Tanaka', price: 61.25, target: 18_000 },
  { name: 'Central', rep: 'Marcus Bell', price: 39.9, target: 8_000 }
];

interface Order {
  readonly region: string;
  /** Year, month, day, as `DATE` takes them. */
  readonly ordered: readonly [number, number, number];
  readonly units: number;
  readonly review: string;
}

/**
 * Ten orders, over a quarter.
 *
 * The unit counts are picked so that six of them miss their target
 * and four beat it, and so that the revenue column spreads widely
 * enough for a colour scale to be a gradient rather than two shades.
 */
const ORDERS: readonly Order[] = [
  { region: 'North', ordered: [2026, 7, 6], units: 180, review: 'Approved' },
  { region: 'South', ordered: [2026, 7, 14], units: 420, review: 'Approved' },
  { region: 'East', ordered: [2026, 7, 23], units: 95, review: 'Pending' },
  { region: 'West', ordered: [2026, 8, 3], units: 260, review: 'Held' },
  { region: 'Central', ordered: [2026, 8, 11], units: 512, review: 'Approved' },
  { region: 'North', ordered: [2026, 8, 19], units: 340, review: 'Pending' },
  { region: 'South', ordered: [2026, 8, 28], units: 78, review: 'Approved' },
  { region: 'East', ordered: [2026, 9, 4], units: 610, review: 'Held' },
  { region: 'West', ordered: [2026, 9, 15], units: 205, review: 'Approved' },
  { region: 'Central', ordered: [2026, 9, 22], units: 150, review: 'Pending' }
];

/** The row the totals sit on, and its one-based twin for the text. */
const TOTAL_ROW = ORDERS.length + 3;
const TOTAL_LINE = TOTAL_ROW + 1;

const ROLLUP_ROW = REGIONS.length + 3;
const ROLLUP_LINE = ROLLUP_ROW + 1;

/** The lookup block on `Reference`, absolute so a fill cannot move it. */
const PRICES = 'Reference!$A$2:$D$6';

const SALES_REGIONS = `Sales!$A$4:$A$${TOTAL_LINE - 1}`;
const SALES_UNITS = `Sales!$E$4:$E$${TOTAL_LINE - 1}`;
const SALES_REVENUE = `Sales!$G$4:$G$${TOTAL_LINE - 1}`;

/**
 * One line each from seven of the eight function families.
 *
 * Written out rather than generated, because the point of the block
 * is that each line is a different question somebody would actually
 * ask — and the last one asks a question with no answer, so that
 * `IFERROR` has something to catch in front of a person rather than
 * only in a spec.
 */
const GLANCES: readonly { label: string; formula: string }[] = [
  { label: 'Largest order', formula: `=LARGE(${SALES_REVENUE},1)` },
  { label: 'Median order', formula: `=MEDIAN(${SALES_REVENUE})` },
  { label: 'Spread, one deviation', formula: `=STDEV(${SALES_REVENUE})` },
  {
    label: 'Region that took it',
    formula: `=INDEX(${SALES_REGIONS},MATCH(LARGE(${SALES_REVENUE},1),${SALES_REVENUE},0))`
  },
  { label: 'Orders over 20,000', formula: `=COUNTIF(${SALES_REVENUE},">20000")` },
  { label: 'Working days in the quarter', formula: '=NETWORKDAYS(DATE(2026,7,1),DATE(2026,9,30))' },
  { label: 'Last order closes', formula: `=TEXT(EOMONTH(MAX(Sales!$C$4:$C$${TOTAL_LINE - 1}),0),"D MMM YYYY")` },
  { label: 'A region not on file', formula: `=IFERROR(VLOOKUP("Nowhere",${PRICES},2,FALSE),"not on file")` }
];

// -------------------------------------------------------------------
// Paint
// -------------------------------------------------------------------

/**
 * The tab colours, as `SheetTabs` names them.
 *
 * Copied rather than imported: that file is a component in the render
 * worker, and this one runs in the application worker, which has no
 * business pulling the chrome across a thread boundary for three
 * strings.
 */
const TAB_BLUE = '#4285f4';
const TAB_ORANGE = '#fa7b17';
const TAB_GREEN = '#34a853';

const BANNER_FILL = '#1f3a5f';
const HEADING_FILL = '#eef2f7';
const NOTE = '#5a6b7d';
const BELOW = '#b3261e';
const HELD_FILL = '#ffe8cc';

const MONEY: NumberFormat = { kind: 'currency', places: 2, symbol: '$' };
const WHOLE_MONEY: NumberFormat = { kind: 'currency', places: 0, symbol: '$' };
const COUNTED: NumberFormat = { kind: 'number', places: 0, thousands: true };
const SHARE: NumberFormat = { kind: 'percent', places: 1 };
const DAY: NumberFormat = { kind: 'date', pattern: 'dmy' };

const RULED_BELOW: CellBorders = { ...NO_BORDERS, bottom: { width: 2, color: '' } };
const RULED_ABOVE: CellBorders = { ...NO_BORDERS, top: { width: 2, color: '' } };

/**
 * A three-colour scale, pale enough to read black text through.
 *
 * Three rather than two so that the middle stop is exercised, which
 * is the half of `scaleColour` that `SCALE_STEPS` being odd exists
 * for.
 */
const SCALE = { from: '#fde2e2', middle: '#fff4cc', to: '#d9efdc' };

function cell(number: NumberFormat, over: Partial<CellPaint> = {}): CellFormat {
  return { number, paint: { ...PLAIN, ...over } };
}

function heading(): CellFormat {
  return cell(GENERAL, { bold: true, fill: HEADING_FILL, align: 'center', borders: RULED_BELOW });
}

function salesFormats(document: SheetDocument): void {
  document.setFormat(
    0,
    0,
    cell(GENERAL, { bold: true, fontSize: 17, color: '#ffffff', fill: BANNER_FILL, align: 'center' })
  );
  document.setFormat(1, 0, cell(GENERAL, { italic: true, color: NOTE, align: 'center' }));

  for (let column = 0; column <= 9; column++) {
    document.setFormat(2, column, heading());
  }

  const last = TOTAL_ROW - 1;
  for (let row = 3; row <= last; row++) {
    document.setFormat(row, 2, cell(DAY));
    document.setFormat(row, 3, cell(DAY, { color: NOTE }));
    document.setFormat(row, 4, cell(COUNTED));
    document.setFormat(row, 5, cell(MONEY));
    document.setFormat(row, 6, cell(MONEY));
    document.setFormat(row, 7, cell(SHARE));
    document.setFormat(row, 8, cell(GENERAL, { align: 'center' }));
    document.setFormat(row, 9, cell(GENERAL, { align: 'center' }));
  }

  const totals: readonly (NumberFormat | null)[] = [
    GENERAL,
    null,
    null,
    null,
    COUNTED,
    MONEY,
    MONEY,
    SHARE,
    GENERAL,
    null
  ];
  totals.forEach((number, column) => {
    if (number !== null) {
      document.setFormat(TOTAL_ROW, column, cell(number, { bold: true, borders: RULED_ABOVE }));
    }
  });
}

function summaryFormats(document: SheetDocument): void {
  document.setFormat(
    0,
    0,
    cell(GENERAL, { bold: true, fontSize: 15, color: '#ffffff', fill: BANNER_FILL, align: 'center' })
  );
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

  document.setFormat(10, 0, cell(GENERAL, { bold: true, borders: RULED_BELOW }));
  // The block is deliberately mixed: four of these are money, one is
  // a count, one a plain number and two are text, so the alignment
  // rule that sends numbers right and everything else left is doing
  // something visible in a column eight rows tall.
  const glanceNumbers: readonly NumberFormat[] = [MONEY, MONEY, MONEY, GENERAL, COUNTED, COUNTED, GENERAL, GENERAL];
  glanceNumbers.forEach((number, index) => {
    document.setFormat(index + 11, 0, cell(GENERAL, { color: NOTE }));
    document.setFormat(index + 11, 1, cell(number));
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

  /**
   * The strict one and the lenient one, side by side and on purpose.
   *
   * A review can only be one of three words and typing a fourth is a
   * mistake with an obvious fix, so that one refuses. A unit count
   * outside the usual range might be a typo and might be a very good
   * quarter, so that one marks the cell and lets it stand — which is
   * the difference `strict` exists to express.
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
// Small helpers
// -------------------------------------------------------------------

/**
 * The widths this sheet wants, over whatever the document already had.
 *
 * A sheet added here rather than through the service starts with no
 * widths at all, and a column with no width is a column drawn at
 * nothing — so the tail is filled in rather than left as holes.
 */
function setWidths(document: SheetDocument, wanted: readonly number[]): void {
  const all = [...document.columnWidths];
  const length = Math.max(all.length, wanted.length);
  for (let column = 0; column < length; column++) {
    all[column] = wanted[column] ?? all[column] ?? COLUMN_WIDTH;
  }
  document.columnWidths = all;
}

function range(address: string): RangeRef {
  const parsed = parseAddress(address);
  if (parsed === null) {
    // Every address in this file is a literal written a few lines
    // above, so this is a typo in the seed and not a runtime case.
    throw new Error(`The seed cannot parse the address ${address}.`);
  }
  return parsed;
}
