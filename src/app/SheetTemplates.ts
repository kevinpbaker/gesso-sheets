import { GENERAL, type NumberFormat } from '../sheet/Format';
import type { SheetDocument } from './SheetDocument';
import {
  banner,
  BELOW,
  cell,
  chart,
  chartsLeft,
  COUNTED,
  DAY,
  GOOD,
  heading,
  HELD_FILL,
  INPUT_FILL,
  MONEY,
  NOTE,
  range,
  RULED_ABOVE,
  seed,
  SHARE,
  setWidths,
  TAB_BLUE,
  TAB_GREEN,
  TAB_ORANGE,
  TAB_PURPLE
} from './SheetSeed';
import type { SheetTemplateBuild } from './SheetService';
import { TEMPLATES } from './Templates';

/**
 * The workbooks File ▸ New can start from, written cell by cell.
 *
 * Each is a function over a fresh document, as the seed is, and the
 * seed is one of them: Northwind Trading is the example a first visit
 * opens on, and from here it can be opened again any time. Yellow is
 * what the person fills in, as it is on the seed's Forecast sheet;
 * everything else is a formula over it.
 *
 * Keyed by the ids in `Templates.ts`, and checked against them by
 * `Templates.spec.ts`, so a card with nothing behind it fails a spec
 * rather than opening a blank.
 */

const BUILDS: Readonly<Record<string, (document: SheetDocument) => void>> = {
  budget,
  invoice,
  tracker,
  timesheet,
  northwind: seed,
  loan
};

export const TEMPLATE_BUILDS: Readonly<Record<string, SheetTemplateBuild>> = Object.fromEntries(
  TEMPLATES.filter(template => Object.hasOwn(BUILDS, template.id)).map(template => [
    template.id,
    { name: template.name, build: BUILDS[template.id] }
  ])
);

const input = cell(GENERAL, { fill: INPUT_FILL });
const inputMoney = cell(MONEY, { fill: INPUT_FILL });
const hint = cell(GENERAL, { italic: true, color: NOTE });
const total = (number = MONEY) => cell(number, { bold: true, borders: RULED_ABOVE });

// -------------------------------------------------------------------
// Monthly budget
// -------------------------------------------------------------------

const INCOME = [
  ['Salary', 4200, 4200],
  ['Side work', 300, 260],
  ['Other', 0, 0]
] as const;

const SPENDING = [
  ['Rent', 1400, 1400],
  ['Groceries', 450, 512],
  ['Utilities', 180, 164],
  ['Transport', 160, 141],
  ['Phone and internet', 90, 90],
  ['Eating out', 200, 268],
  ['Fun', 150, 95],
  ['Savings', 600, 600]
] as const;

function budget(document: SheetDocument): void {
  document.renameSheet(0, 'Budget');
  document.setSheetColour(0, TAB_GREEN);
  setWidths(document, [180, 110, 110, 110]);

  document.setCell(0, 0, 'Monthly budget');
  document.setCell(1, 0, 'Fill in the yellow cells. The totals, what is left and the chart follow.');

  // Income: lines 4 to 8.
  ['Income', 'Planned', 'Actual', 'Difference'].forEach((text, column) => document.setCell(3, column, text));
  INCOME.forEach(([name, planned, actual], index) => {
    const row = 4 + index;
    document.setCell(row, 0, name);
    document.setCell(row, 1, String(planned));
    document.setCell(row, 2, String(actual));
    document.setCell(row, 3, `=C${row + 1}-B${row + 1}`);
  });
  document.setCell(7, 0, 'Total income');
  for (const [column, letter] of [[1, 'B'], [2, 'C'], [3, 'D']] as const) {
    document.setCell(7, column, `=SUM(${letter}5:${letter}7)`);
  }

  // Spending: lines 10 to 19. Under budget is the good direction here,
  // so the difference is the other way round from income's.
  ['Spending', 'Planned', 'Actual', 'Difference'].forEach((text, column) => document.setCell(9, column, text));
  SPENDING.forEach(([name, planned, actual], index) => {
    const row = 10 + index;
    document.setCell(row, 0, name);
    document.setCell(row, 1, String(planned));
    document.setCell(row, 2, String(actual));
    document.setCell(row, 3, `=B${row + 1}-C${row + 1}`);
  });
  document.setCell(18, 0, 'Total spending');
  for (const [column, letter] of [[1, 'B'], [2, 'C'], [3, 'D']] as const) {
    document.setCell(18, column, `=SUM(${letter}11:${letter}18)`);
  }

  document.setCell(20, 0, 'Left over');
  document.setCell(20, 1, '=B8-B19');
  document.setCell(20, 2, '=C8-C19');
  document.setCell(21, 0, '="Saving "&TEXT((C21+C18)/C8,"0%")&" of what came in"');

  document.setFormat(0, 0, banner(16));
  document.setFormat(1, 0, hint);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 3 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 3 });
  for (const head of [3, 9]) {
    for (let column = 0; column <= 3; column++) {
      document.setFormat(head, column, heading());
    }
  }
  for (const [first, last] of [[4, 6], [10, 17]]) {
    for (let row = first; row <= last; row++) {
      document.setFormat(row, 0, input);
      document.setFormat(row, 1, inputMoney);
      document.setFormat(row, 2, inputMoney);
      document.setFormat(row, 3, cell(MONEY));
    }
  }
  for (const row of [7, 18]) {
    document.setFormat(row, 0, total(GENERAL));
    for (let column = 1; column <= 3; column++) {
      document.setFormat(row, column, total());
    }
  }
  document.setFormat(20, 0, cell(GENERAL, { bold: true }));
  document.setFormat(20, 1, cell(MONEY, { bold: true }));
  document.setFormat(20, 2, cell(MONEY, { bold: true }));
  document.setFormat(21, 0, hint);

  for (const over of ['D5:D7', 'D11:D18', 'B21:C21']) {
    document.addConditional({ range: range(over), test: { kind: 'lessThan', value: 0 }, paint: { color: BELOW, bold: true } });
  }
  // Two ranges, around the spending headings rather than over them.
  for (const amounts of ['B5:C7', 'B11:C18']) {
    document.addValidation({
      range: range(amounts),
      rule: { kind: 'number', min: 0 },
      message: 'An amount is a number, nought or more.'
    });
  }

  chart(document, 'column', 'Spending, planned and actual', range('A10:C18'), chartsLeft(document, 4), 40, 460, 300);
  document.setSelection(4, 2, 4, 2);
}

// -------------------------------------------------------------------
// Invoice
// -------------------------------------------------------------------

const LINE_ITEMS = 10;
const FIRST_ITEM = 12;

function invoice(document: SheetDocument): void {
  document.renameSheet(0, 'Invoice');
  document.setSheetColour(0, TAB_BLUE);
  setWidths(document, [280, 70, 110, 120]);

  document.setCell(0, 0, 'Invoice');

  document.setCell(2, 0, 'From');
  document.setCell(3, 0, 'Your business name');
  document.setCell(4, 0, 'Street, town, postcode');
  document.setCell(5, 0, 'you@example.com');
  document.setCell(7, 0, 'Bill to');
  document.setCell(8, 0, 'Client name');
  document.setCell(9, 0, 'Client address');

  document.setCell(2, 2, 'Invoice no.');
  // Not `0001`, which is the number one.
  document.setCell(2, 3, 'INV-0001');
  document.setCell(3, 2, 'Date');
  document.setCell(3, 3, '=TODAY()');
  document.setCell(4, 2, 'Due');
  document.setCell(4, 3, '=D4+D6');
  document.setCell(5, 2, 'Days to pay');
  document.setCell(5, 3, '30');

  ['Description', 'Qty', 'Unit price', 'Amount'].forEach((text, column) => document.setCell(FIRST_ITEM - 1, column, text));
  const samples = [
    ['Design work, by the hour', 12, 85],
    ['Hosting, one year', 1, 240]
  ] as const;
  for (let index = 0; index < LINE_ITEMS; index++) {
    const row = FIRST_ITEM + index;
    const line = row + 1;
    const sample = samples[index];
    if (sample !== undefined) {
      document.setCell(row, 0, sample[0]);
      document.setCell(row, 1, String(sample[1]));
      document.setCell(row, 2, String(sample[2]));
    }
    // Blank until there is a quantity, so the unused lines stay empty.
    document.setCell(row, 3, `=IF(B${line}="","",B${line}*C${line})`);
  }

  const items = `${FIRST_ITEM + 1}:D${FIRST_ITEM + LINE_ITEMS}`;
  const sums = FIRST_ITEM + LINE_ITEMS + 1;
  document.setCell(sums, 2, 'Subtotal');
  document.setCell(sums, 3, `=SUM(D${items})`);
  document.setCell(sums + 1, 2, 'Tax rate');
  document.setCell(sums + 1, 3, '0.1');
  document.setCell(sums + 2, 2, 'Tax');
  document.setCell(sums + 2, 3, `=ROUND(D${sums + 1}*D${sums + 2},2)`);
  document.setCell(sums + 3, 2, 'Total due');
  document.setCell(sums + 3, 3, `=D${sums + 1}+D${sums + 3}`);
  document.setCell(sums + 5, 0, '="Payment is due by "&TEXT(D5,"d mmmm yyyy")&". Thank you for your business."');

  document.setFormat(0, 0, banner(18));
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 3 });
  document.setFormat(2, 0, cell(GENERAL, { bold: true }));
  document.setFormat(7, 0, cell(GENERAL, { bold: true }));
  for (const row of [3, 4, 5, 8, 9]) {
    document.setFormat(row, 0, input);
  }
  for (let row = 2; row <= 5; row++) {
    document.setFormat(row, 2, cell(GENERAL, { color: NOTE, align: 'end' }));
  }
  document.setFormat(2, 3, cell(GENERAL, { fill: INPUT_FILL, align: 'end' }));
  document.setFormat(3, 3, cell(DAY, { fill: INPUT_FILL }));
  document.setFormat(4, 3, cell(DAY));
  document.setFormat(5, 3, cell(COUNTED, { fill: INPUT_FILL }));
  for (let column = 0; column <= 3; column++) {
    document.setFormat(FIRST_ITEM - 1, column, heading());
  }
  for (let row = FIRST_ITEM; row < FIRST_ITEM + LINE_ITEMS; row++) {
    document.setFormat(row, 0, input);
    document.setFormat(row, 1, cell(COUNTED, { fill: INPUT_FILL }));
    document.setFormat(row, 2, inputMoney);
    document.setFormat(row, 3, cell(MONEY));
  }
  for (let row = sums; row <= sums + 3; row++) {
    document.setFormat(row, 2, cell(GENERAL, { color: NOTE, align: 'end' }));
  }
  document.setFormat(sums, 3, cell(MONEY, { borders: RULED_ABOVE }));
  document.setFormat(sums + 1, 3, cell(SHARE, { fill: INPUT_FILL }));
  document.setFormat(sums + 2, 3, cell(MONEY));
  document.setFormat(sums + 3, 2, cell(GENERAL, { bold: true, align: 'end' }));
  document.setFormat(sums + 3, 3, total());
  document.setFormat(sums + 5, 0, hint);

  document.addValidation({
    range: range(`B${FIRST_ITEM + 1}:C${FIRST_ITEM + LINE_ITEMS}`),
    rule: { kind: 'number', min: 0 },
    message: 'A quantity or a price is a number, nought or more.'
  });
  document.setSelection(3, 0, 3, 0);
}

// -------------------------------------------------------------------
// Project tracker
// -------------------------------------------------------------------

const STATUSES = ['Not started', 'In progress', 'Blocked', 'Done'];
const TASK_ROWS = 30;

function tracker(document: SheetDocument): void {
  document.renameSheet(0, 'Tasks');
  document.setSheetColour(0, TAB_PURPLE);
  setWidths(document, [240, 110, 110, 100, 90, 240]);

  const last = 3 + TASK_ROWS;
  document.setCell(0, 0, 'Project tracker');
  document.setCell(
    1,
    0,
    `=COUNTIF(C4:C${last},"Done")&" of "&COUNTA(A4:A${last})&" done, "&COUNTIF(C4:C${last},"Blocked")&" blocked, "&COUNTIFS(E4:E${last},"<0")&" overdue"`
  );
  ['Task', 'Owner', 'Status', 'Due', 'Days left', 'Notes'].forEach((text, column) => document.setCell(2, column, text));

  const samples = [
    ['Write the brief', 'Sam', 'Done', -10, ''],
    ['Gather requirements', 'Alex', 'Done', -4, ''],
    ['First draft', 'Sam', 'In progress', 3, ''],
    ['Review with the team', 'Jordan', 'Not started', 8, ''],
    ['Sign-off', 'Alex', 'Blocked', -1, 'Waiting on the budget']
  ] as const;
  samples.forEach(([task, owner, status, due, notes], index) => {
    const row = 3 + index;
    document.setCell(row, 0, task);
    document.setCell(row, 1, owner);
    document.setCell(row, 2, status);
    document.setCell(row, 3, due < 0 ? `=TODAY()${due}` : `=TODAY()+${due}`);
    document.setCell(row, 5, notes);
  });
  for (let row = 3; row < last; row++) {
    const line = row + 1;
    document.setCell(row, 4, `=IF(OR(D${line}="",C${line}="Done"),"",D${line}-TODAY())`);
  }

  document.setFormat(0, 0, banner(16));
  document.setFormat(1, 0, hint);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 5 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 5 });
  for (let column = 0; column <= 5; column++) {
    document.setFormat(2, column, heading());
  }
  for (let row = 3; row < last; row++) {
    document.setFormat(row, 3, cell(DAY));
    document.setFormat(row, 4, cell(COUNTED, { color: NOTE }));
  }
  document.frozenRows = 3;

  const status = range(`C4:C${last}`);
  document.addValidation({
    range: status,
    rule: { kind: 'list', values: STATUSES },
    strict: true,
    message: `A status is ${STATUSES.slice(0, -1).join(', ')} or ${STATUSES[STATUSES.length - 1]}.`
  });
  document.addConditional({ range: status, test: { kind: 'equalTo', value: 'Done' }, paint: { color: GOOD, bold: true } });
  document.addConditional({ range: status, test: { kind: 'equalTo', value: 'Blocked' }, paint: { fill: HELD_FILL } });
  // A formula rule rather than "less than nought", so an empty cell,
  // which is not a number, is never overdue.
  document.addConditional({
    range: range(`E4:E${last}`),
    test: { kind: 'formula', input: '=AND(ISNUMBER(E4),E4<0)' },
    paint: { color: BELOW, bold: true }
  });
  document.setNote(2, 4, 'Days until it is due. Below nought is overdue; a task that is done shows nothing.');
  document.setSelection(3 + samples.length, 0, 3 + samples.length, 0);
}

// -------------------------------------------------------------------
// Weekly timesheet
// -------------------------------------------------------------------

const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const SAMPLE_HOURS = ['8', '8', '7.5', '8', '6', '', ''];
const HOURS: NumberFormat = { kind: 'number', places: 2, thousands: false };

function timesheet(document: SheetDocument): void {
  document.renameSheet(0, 'Timesheet');
  document.setSheetColour(0, TAB_ORANGE);
  setWidths(document, [140, 110, 200, 80, 240]);

  document.setCell(0, 0, 'Weekly timesheet');
  document.setCell(2, 0, 'Name');
  document.setCell(2, 1, 'Your name');
  document.setCell(3, 0, 'Week starting');
  // This week's Monday: WEEKDAY's type 3 counts Monday as nought.
  document.setCell(3, 1, '=TODAY()-WEEKDAY(TODAY(),3)');
  document.setCell(4, 0, 'Hourly rate');
  document.setCell(4, 1, '45');

  ['Day', 'Date', 'Project', 'Hours', 'Notes'].forEach((text, column) => document.setCell(6, column, text));
  DAYS.forEach((day, index) => {
    const row = 7 + index;
    document.setCell(row, 0, day);
    document.setCell(row, 1, `=$B$4+${index}`);
    document.setCell(row, 3, SAMPLE_HOURS[index]);
  });
  document.setCell(14, 0, 'Total');
  document.setCell(14, 3, '=SUM(D8:D14)');
  document.setCell(15, 0, 'Overtime, past 40');
  document.setCell(15, 3, '=MAX(0,D15-40)');
  document.setCell(16, 0, 'Pay');
  document.setCell(16, 3, '=(D15-D16)*B5+D16*B5*1.5');
  document.setCell(17, 0, 'Overtime is paid at time and a half.');

  document.setFormat(0, 0, banner(16));
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 });
  for (let row = 2; row <= 4; row++) {
    document.setFormat(row, 0, cell(GENERAL, { color: NOTE }));
  }
  document.setFormat(2, 1, input);
  document.setFormat(3, 1, cell(DAY, { fill: INPUT_FILL }));
  document.setFormat(4, 1, inputMoney);
  for (let column = 0; column <= 4; column++) {
    document.setFormat(6, column, heading());
  }
  for (let row = 7; row <= 13; row++) {
    document.setFormat(row, 1, cell(DAY, { color: NOTE }));
    document.setFormat(row, 2, input);
    document.setFormat(row, 3, cell(HOURS, { fill: INPUT_FILL }));
    document.setFormat(row, 4, input);
  }
  document.setFormat(14, 0, total(GENERAL));
  document.setFormat(14, 3, total(HOURS));
  document.setFormat(15, 3, cell(HOURS));
  document.setFormat(16, 0, cell(GENERAL, { bold: true }));
  document.setFormat(16, 3, cell(MONEY, { bold: true }));
  document.setFormat(17, 0, hint);

  document.addValidation({
    range: range('D8:D14'),
    rule: { kind: 'number', min: 0, max: 24 },
    strict: true,
    message: 'Hours in a day are 0 to 24.'
  });
  document.addConditional({ range: range('D8:D14'), test: { kind: 'greaterThan', value: 10 }, paint: { fill: HELD_FILL } });
  document.setSelection(7, 2, 7, 2);
}

// -------------------------------------------------------------------
// Example: paying off a loan
// -------------------------------------------------------------------

const MONTHS = 60;
const SCHEDULE_HEAD = 10;

function loan(document: SheetDocument): void {
  document.renameSheet(0, 'Loan');
  document.setSheetColour(0, TAB_BLUE);
  setWidths(document, [130, 120, 110, 110, 110]);

  document.setCell(0, 0, 'Paying off a loan');
  document.setCell(
    1,
    0,
    'Change the amount, the rate or the term: the schedule and the chart follow.'
  );
  document.setCell(3, 0, 'Borrowed');
  document.setCell(3, 1, '30000');
  document.setCell(4, 0, 'Rate a year');
  document.setCell(4, 1, '0.069');
  document.setCell(5, 0, 'Years');
  document.setCell(5, 1, '5');
  document.setCell(6, 0, 'A month');
  document.setCell(6, 1, '=PMT(B5/12,B6*12,-B4)');
  document.setCell(7, 0, 'Interest, all told');

  ['Month', 'Balance', 'Payment', 'Interest', 'Principal'].forEach((text, column) =>
    document.setCell(SCHEDULE_HEAD, column, text)
  );
  for (let month = 1; month <= MONTHS; month++) {
    const row = SCHEDULE_HEAD + month;
    const line = row + 1;
    // The balance before this month: what was borrowed, for the first.
    const before = month === 1 ? '$B$4' : `B${line - 1}`;
    document.setCell(row, 0, String(month));
    // After the term the payment is nought and so is everything else.
    document.setCell(row, 2, `=IF(A${line}>$B$6*12,0,$B$7)`);
    document.setCell(row, 3, `=${before}*$B$5/12`);
    document.setCell(row, 4, `=C${line}-D${line}`);
    document.setCell(row, 1, `=MAX(0,${before}-E${line})`);
  }
  // Each month's interest is its own formula over the month before, so
  // the column is summed as a range: there is no spill to point at.
  document.setCell(7, 1, `=SUM(D${SCHEDULE_HEAD + 2}:D${SCHEDULE_HEAD + 1 + MONTHS})`);

  document.setFormat(0, 0, banner(16));
  document.setFormat(1, 0, hint);
  document.merges.add({ firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 4 });
  document.merges.add({ firstRow: 1, lastRow: 1, firstColumn: 0, lastColumn: 4 });
  for (let row = 3; row <= 7; row++) {
    document.setFormat(row, 0, cell(GENERAL, { color: NOTE }));
  }
  document.setFormat(3, 1, inputMoney);
  document.setFormat(4, 1, cell({ kind: 'percent', places: 2 }, { fill: INPUT_FILL }));
  document.setFormat(5, 1, cell(COUNTED, { fill: INPUT_FILL }));
  document.setFormat(6, 1, cell(MONEY, { bold: true }));
  document.setFormat(7, 1, cell(MONEY));
  document.setNote(6, 1, 'PMT: the same payment every month that clears the loan over the term.');
  document.setNote(SCHEDULE_HEAD, 3, 'The balance before this month, at a twelfth of the yearly rate.');
  for (let column = 0; column <= 4; column++) {
    document.setFormat(SCHEDULE_HEAD, column, heading());
  }
  for (let row = SCHEDULE_HEAD + 1; row <= SCHEDULE_HEAD + MONTHS; row++) {
    document.setFormat(row, 0, cell(COUNTED, { color: NOTE, align: 'center' }));
    for (let column = 1; column <= 4; column++) {
      document.setFormat(row, column, cell(MONEY));
    }
  }

  document.addValidation({
    range: range('B6:B6'),
    rule: { kind: 'number', min: 1, max: MONTHS / 12, integer: true },
    strict: true,
    message: `The schedule runs to ${MONTHS / 12} years: a term is 1 to ${MONTHS / 12}.`
  });
  document.addConditional({
    range: range(`B${SCHEDULE_HEAD + 2}:B${SCHEDULE_HEAD + 1 + MONTHS}`),
    scale: { from: '#d9efdc', middle: '#fff4cc', to: '#fde2e2' },
    test: null
  });

  chart(
    document,
    'area',
    'What is still owed',
    range(`A${SCHEDULE_HEAD + 1}:B${SCHEDULE_HEAD + 1 + MONTHS}`),
    chartsLeft(document, 5),
    40,
    460,
    280
  );
  document.setSelection(3, 1, 3, 1);
}
