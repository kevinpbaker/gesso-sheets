/**
 * What File ▸ New offers besides a blank: the templates and the
 * examples, by name.
 *
 * Only the names, and on purpose. The dialog that lists them is in the
 * render worker, which never holds a cell, and the functions that
 * write each one are in `SheetTemplates.ts` on the application worker,
 * which builds the document they write into. This list is the one fact
 * both threads need, as `route.ts` is, so it imports nothing.
 *
 * A template is a skeleton to fill in: the headings, the formulas and
 * a line or two showing what goes where. An example is a finished
 * workbook to look around, opened as a copy of its own so nothing
 * done to it reaches the next one.
 */

export type TemplateKind = 'template' | 'example';

export interface TemplateInfo {
  /** What the route carries: `/new/<id>`. */
  readonly id: string;
  /** The card's title, and the new document's name. */
  readonly name: string;
  /** One line under the title, saying what it is for. */
  readonly blurb: string;
  readonly kind: TemplateKind;
}

export const TEMPLATES: readonly TemplateInfo[] = [
  { id: 'budget', name: 'Monthly budget', blurb: 'Income and spending, planned against actual.', kind: 'template' },
  { id: 'invoice', name: 'Invoice', blurb: 'Line items, tax and a total, with a due date.', kind: 'template' },
  { id: 'tracker', name: 'Project tracker', blurb: 'Tasks, owners and status, with what is overdue.', kind: 'template' },
  { id: 'timesheet', name: 'Weekly timesheet', blurb: 'Hours by day, overtime and pay.', kind: 'template' },
  {
    id: 'northwind',
    name: 'Northwind Trading',
    blurb: 'A year of a small business: a dashboard, a ledger, a forecast and charts.',
    kind: 'example'
  },
  { id: 'loan', name: 'Paying off a loan', blurb: 'A repayment schedule month by month, and its chart.', kind: 'example' }
];
