import { columnName, parseAddress, relativeRef, type RangeRef } from './A1';
import type { ColourScale, ConditionalPaint, ConditionalRule, ConditionalTest } from './Conditional';
import { rewriteFormula } from './Rewrite';
import type { Validation, ValidationRule } from './Validation';
import { child, children, type XmlElement } from './Xml';

/**
 * What cells may hold, between this sheet and an `.xlsx` — Phase 23.
 *
 * Excel's `<dataValidation>` is a type (`list`, `whole`, `decimal`,
 * `date`, `textLength`, `custom`, `time`), an operator, one or two
 * formulas and a range, which is a superset of what Phase 14 keeps: a
 * list, a number between two bounds and maybe whole, a longest text, a
 * date between two days. Everything that maps is kept, and what does
 * not — a custom formula, a time, a `notBetween` — is counted, so the
 * import's sentence can say how many were left out rather than keeping
 * a rule that would enforce something else.
 */

/**
 * A list whose values are a range of cells, to be read once every sheet
 * is — named outright, or by a defined name (`name`), which is how most
 * of the lists in POI's files are written and whose range is only known
 * once the workbook's names are.
 */
export interface PendingList {
  readonly at: number;
  readonly name?: string;
  readonly sheet: string | null;
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

export interface ReadValidations {
  readonly validations: Validation[];
  /** Lists still to be filled from their ranges; `at` indexes `validations`. */
  readonly pending: PendingList[];
  /** Validations that mean something this sheet cannot say. */
  readonly skipped: number;
}

/** Days between 1900's day zero and 1904's, for a date bound in a 1904 file. */
const EPOCH_1904 = 1462;

export function readValidations(root: XmlElement, date1904: boolean): ReadValidations {
  const validations: Validation[] = [];
  const pending: PendingList[] = [];
  let skipped = 0;
  for (const entry of children(child(root, 'dataValidations'), 'dataValidation')) {
    // No type is Excel's "any value": a prompt with nothing to enforce,
    // which is not a rule this sheet is failing to keep.
    if ((entry.attributes.type ?? 'none') === 'none') {
      continue;
    }
    const ranges = (entry.attributes.sqref ?? '')
      .split(/\s+/)
      .filter(part => part !== '')
      .map(part => parseAddress(part))
      .filter((range): range is RangeRef => range !== null);
    const read = ruleOf(entry, date1904);
    if (ranges.length === 0 || read === null) {
      skipped++;
      continue;
    }
    const style = entry.attributes.errorStyle ?? 'stop';
    const shows = entry.attributes.showErrorMessage === '1' || entry.attributes.showErrorMessage === 'true';
    const message = entry.attributes.error?.trim();
    for (const range of ranges) {
      const validation: Validation = {
        range,
        rule: read.rule,
        // Excel refuses a value only with the message on and the style
        // `stop`; a warning or a note lets the value in, as a rule here
        // that is not strict does.
        strict: shows && style === 'stop',
        ...(message === undefined || message === '' ? {} : { message })
      };
      if (read.list !== undefined) {
        pending.push({ at: validations.length, ...read.list });
      }
      validations.push(validation);
    }
  }
  return { validations, pending, skipped };
}

function ruleOf(
  entry: XmlElement,
  date1904: boolean
): { rule: ValidationRule; list?: Omit<PendingList, 'at'> } | null {
  const type = entry.attributes.type ?? 'none';
  const operator = entry.attributes.operator ?? 'between';
  const first = child(entry, 'formula1')?.text.trim() ?? '';
  const second = child(entry, 'formula2')?.text.trim() ?? '';
  switch (type) {
    case 'list': {
      if (first.startsWith('"')) {
        const values = first
          .slice(1, first.endsWith('"') ? -1 : undefined)
          .replace(/""/g, '"')
          .split(',')
          .map(value => value.trim())
          .filter(value => value !== '');
        return values.length === 0 ? null : { rule: { kind: 'list', values } };
      }
      const range = listRange(first);
      if (range !== null) {
        return { rule: { kind: 'list', values: [] }, list: range };
      }
      const named = /^=?([A-Za-z_\\][\w.]*)$/.exec(first);
      return named === null
        ? null
        : { rule: { kind: 'list', values: [] }, list: { name: named[1], sheet: null, firstRow: 0, firstColumn: 0, lastRow: -1, lastColumn: -1 } };
    }
    case 'whole':
    case 'decimal': {
      const bounds = boundsOf(operator, first, second, type === 'whole', 0);
      return bounds === null ? null : { rule: { kind: 'number', ...bounds, ...(type === 'whole' ? { integer: true } : {}) } };
    }
    case 'date': {
      const bounds = boundsOf(operator, first, second, true, date1904 ? EPOCH_1904 : 0);
      return bounds === null ? null : { rule: { kind: 'date', from: bounds.min, to: bounds.max } };
    }
    case 'textLength': {
      // Only an upper bound can be kept: the rule here is a longest text.
      const length = Number(operator === 'between' ? second : first);
      if (!Number.isFinite(length) || !['lessThanOrEqual', 'lessThan', 'between'].includes(operator)) {
        return null;
      }
      return { rule: { kind: 'text', maxLength: operator === 'lessThan' ? length - 1 : length } };
    }
    default:
      return null;
  }
}

/**
 * The bounds an operator puts on a number, or null for one this sheet
 * cannot say. A strict bound on a whole number moves by one, so
 * `greaterThan 0` is `min 1`; on a decimal it is kept as inclusive,
 * which lets through the one value Excel would refuse.
 */
function boundsOf(
  operator: string,
  first: string,
  second: string,
  whole: boolean,
  shift: number
): { min?: number; max?: number } | null {
  const a = Number(first);
  const b = Number(second);
  if (!Number.isFinite(a)) {
    return null;
  }
  const step = whole ? 1 : 0;
  switch (operator) {
    case 'between':
      return Number.isFinite(b) ? { min: Math.min(a, b) + shift, max: Math.max(a, b) + shift } : null;
    case 'equal':
      return { min: a + shift, max: a + shift };
    case 'greaterThanOrEqual':
      return { min: a + shift };
    case 'greaterThan':
      return { min: a + step + shift };
    case 'lessThanOrEqual':
      return { max: a + shift };
    case 'lessThan':
      return { max: a - step + shift };
    default:
      return null;
  }
}

/** `$D$1:$D$5` or `Lists!$A$1:$A$9`, as the range a list's values come from. */
function listRange(text: string): Omit<PendingList, 'at'> | null {
  const match = /^=?(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?$/i.exec(text);
  if (match === null) {
    return null;
  }
  const range = parseAddress(`${match[3]}${match[4] === undefined ? '' : `:${match[4]}`}`);
  if (range === null) {
    return null;
  }
  return {
    sheet: match[1]?.replace(/''/g, "'") ?? match[2] ?? null,
    firstRow: Math.min(range.start.row, range.end.row),
    firstColumn: Math.min(range.start.column, range.end.column),
    lastRow: Math.max(range.start.row, range.end.row),
    lastColumn: Math.max(range.start.column, range.end.column)
  };
}

// ---------------------------------------------------------------------------
// Out
// ---------------------------------------------------------------------------

/**
 * The sheet's validations as `<dataValidations>`, or empty. A rule that
 * cannot be written is left out and counted in `unwritten`.
 */
export function writeValidations(validations: readonly Validation[], date1904 = false): { xml: string; unwritten: number } {
  let unwritten = 0;
  const written: string[] = [];
  for (const validation of validations) {
    const body = xmlOf(validation.rule, date1904 ? -EPOCH_1904 : 0);
    if (body === null) {
      unwritten++;
      continue;
    }
    const strict = validation.strict === true;
    const message = validation.message === undefined ? '' : ` error="${escapeAttribute(validation.message)}"`;
    written.push(
      `<dataValidation ${body.attributes} allowBlank="1" showErrorMessage="1" errorStyle="${strict ? 'stop' : 'warning'}"${message} sqref="${sqrefOf(validation.range)}">${body.formulas}</dataValidation>`
    );
  }
  return {
    xml: written.length === 0 ? '' : `<dataValidations count="${written.length}">${written.join('')}</dataValidations>`,
    unwritten
  };
}

function xmlOf(rule: ValidationRule, shift: number): { attributes: string; formulas: string } | null {
  switch (rule.kind) {
    case 'list': {
      const joined = rule.values.join(',');
      // Excel's own limit on a list written into the rule, and a comma
      // inside a value cannot be told from the one between values.
      if (rule.values.length === 0 || joined.length > 255 || rule.values.some(value => value.includes(','))) {
        return null;
      }
      return { attributes: 'type="list"', formulas: `<formula1>"${escapeText(joined.replace(/"/g, '""'))}"</formula1>` };
    }
    case 'number':
      return bounded(rule.integer === true ? 'whole' : 'decimal', rule.min, rule.max, 0);
    case 'date':
      return bounded('date', rule.from, rule.to, shift);
    case 'text':
      return rule.maxLength === undefined
        ? null
        : { attributes: 'type="textLength" operator="lessThanOrEqual"', formulas: `<formula1>${rule.maxLength}</formula1>` };
  }
}

function bounded(type: string, min: number | undefined, max: number | undefined, shift: number): { attributes: string; formulas: string } {
  if (min !== undefined && max !== undefined) {
    return {
      attributes: `type="${type}" operator="between"`,
      formulas: `<formula1>${min + shift}</formula1><formula2>${max + shift}</formula2>`
    };
  }
  if (min !== undefined) {
    return { attributes: `type="${type}" operator="greaterThanOrEqual"`, formulas: `<formula1>${min + shift}</formula1>` };
  }
  if (max !== undefined) {
    return { attributes: `type="${type}" operator="lessThanOrEqual"`, formulas: `<formula1>${max + shift}</formula1>` };
  }
  // Any number at all: Excel has no rule without a bound, so the widest.
  return { attributes: `type="${type}" operator="between"`, formulas: '<formula1>-1E+307</formula1><formula2>1E+307</formula2>' };
}

/** A range as `sqref` writes it: `B2:B9`, or `A:A` for a whole column. */
export function sqrefOf(range: RangeRef): string {
  const start = range.start;
  const end = range.end;
  if (range.wholeColumn === true) {
    return `${columnName(start.column)}:${columnName(end.column)}`;
  }
  const first = `${columnName(Math.min(start.column, end.column))}${Math.min(start.row, end.row) + 1}`;
  const last = `${columnName(Math.max(start.column, end.column))}${Math.max(start.row, end.row) + 1}`;
  return first === last ? first : `${first}:${last}`;
}

/** A rectangle as a range this sheet's rules hold. */
export function rangeOfRect(firstRow: number, firstColumn: number, lastRow: number, lastColumn: number): RangeRef {
  return { start: relativeRef(firstRow, firstColumn), end: relativeRef(lastRow, lastColumn) };
}

function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttribute(text: string): string {
  return escapeText(text).replace(/"/g, '&quot;').replace(/\n/g, '&#10;');
}

// ---------------------------------------------------------------------------
// Conditional formats
// ---------------------------------------------------------------------------

/**
 * Excel's `<conditionalFormatting>` as this sheet's rules — Phase 23.
 *
 * A `cfRule` is a type, maybe an operator, one or more formulas, and a
 * `dxfId` into the styles' differential formats for what it paints.
 * The shorthands Phase 14 has — greater, less, between, equal, contains
 * text, blank, not blank — are read as themselves, because they are
 * answered without an evaluator and a formula would cost the frame
 * budget they exist to save. Almost everything else Excel writes comes
 * with the formula it means, even the rules with names of their own
 * (`beginsWith`, `containsErrors`, `notContainsText`), and is kept as
 * that formula. A two- or three-colour scale is a scale. What has no
 * formula and no counterpart — data bars, icon sets, top ten, above
 * average, duplicates — is counted and left out.
 */
export function readConditionals(
  root: XmlElement,
  dxfs: readonly ConditionalPaint[],
  inputOf: (excel: string) => string
): { rules: ConditionalRule[]; skipped: number } {
  const found: { priority: number; order: number; rule: ConditionalRule }[] = [];
  let skipped = 0;
  let order = 0;
  for (const block of children(root, 'conditionalFormatting')) {
    const ranges = (block.attributes.sqref ?? '')
      .split(/\s+/)
      .filter(part => part !== '')
      .map(part => parseAddress(part))
      .filter((range): range is RangeRef => range !== null);
    if (ranges.length === 0) {
      skipped += children(block, 'cfRule').length;
      continue;
    }
    const anchor = ranges[0].start;
    for (const cfRule of children(block, 'cfRule')) {
      const read = cfRuleOf(cfRule, dxfs, inputOf, `${columnName(anchor.column)}${anchor.row + 1}`);
      if (read === null) {
        skipped++;
        continue;
      }
      const priority = Number(cfRule.attributes.priority ?? Number.MAX_SAFE_INTEGER);
      for (const range of ranges) {
        // A formula is written for the first range's first cell, so a
        // second range is given it moved to its own first cell.
        const test =
          read.test?.kind === 'formula'
            ? { kind: 'formula' as const, input: rewriteFormula(read.test.input, range.start.row - anchor.row, range.start.column - anchor.column) }
            : read.test;
        found.push({ priority, order: order++, rule: { range, test, ...(read.paint === undefined ? {} : { paint: read.paint }), ...(read.scale === undefined ? {} : { scale: read.scale }) } });
      }
    }
  }
  // Excel's priority is the order the rules are tried in, the lowest
  // first — which is this sheet's order too.
  found.sort((a, b) => a.priority - b.priority || a.order - b.order);
  return { rules: found.map(entry => entry.rule), skipped };
}

function cfRuleOf(
  cfRule: XmlElement,
  dxfs: readonly ConditionalPaint[],
  inputOf: (excel: string) => string,
  anchor: string
): { test: ConditionalTest | null; paint?: ConditionalPaint; scale?: ColourScale } | null {
  const type = cfRule.attributes.type ?? '';
  if (type === 'colorScale') {
    const colours = children(child(cfRule, 'colorScale'), 'color').map(rgbOf);
    if (colours.length < 2 || colours.some(colour => colour === null)) {
      return null;
    }
    const [from, middle, to] = colours.length === 2 ? [colours[0], undefined, colours[1]] : colours;
    return { test: null, scale: { from: from!, to: to!, ...(middle == null ? {} : { middle }) } };
  }
  const dxf = dxfs[Number(cfRule.attributes.dxfId ?? -1)];
  const paint = dxf === undefined || Object.keys(dxf).length === 0 ? undefined : dxf;
  if (paint === undefined) {
    // A rule that paints nothing does nothing here.
    return null;
  }
  const formulas = children(cfRule, 'formula').map(formula => formula.text.trim());
  const test = testOf(type, cfRule, formulas, inputOf, anchor);
  return test === null ? null : { test, paint };
}

function testOf(
  type: string,
  cfRule: XmlElement,
  formulas: readonly string[],
  inputOf: (excel: string) => string,
  anchor: string
): ConditionalTest | null {
  const [first = '', second = ''] = formulas;
  const number = (text: string): number | null => (/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(text) ? Number(text) : null);
  switch (type) {
    case 'cellIs': {
      const operator = cfRule.attributes.operator ?? '';
      const a = number(first);
      const b = number(second);
      if (operator === 'greaterThan' && a !== null) {
        return { kind: 'greaterThan', value: a };
      }
      if (operator === 'lessThan' && a !== null) {
        return { kind: 'lessThan', value: a };
      }
      if (operator === 'between' && a !== null && b !== null) {
        return { kind: 'between', low: Math.min(a, b), high: Math.max(a, b) };
      }
      if (operator === 'equal') {
        if (a !== null) {
          return { kind: 'equalTo', value: a };
        }
        if (/^"(?:[^"]|"")*"$/.test(first)) {
          return { kind: 'equalTo', value: first.slice(1, -1).replace(/""/g, '"') };
        }
      }
      // Anything else a comparison can say, as the comparison.
      const comparisons: Record<string, (x: string, y: string) => string> = {
        greaterThan: x => `${anchor}>${x}`,
        lessThan: x => `${anchor}<${x}`,
        greaterThanOrEqual: x => `${anchor}>=${x}`,
        lessThanOrEqual: x => `${anchor}<=${x}`,
        equal: x => `${anchor}=${x}`,
        notEqual: x => `${anchor}<>${x}`,
        between: (x, y) => `AND(${anchor}>=MIN(${x},${y}),${anchor}<=MAX(${x},${y}))`,
        notBetween: (x, y) => `OR(${anchor}<MIN(${x},${y}),${anchor}>MAX(${x},${y}))`
      };
      const write = comparisons[operator];
      return write === undefined || first === '' ? null : { kind: 'formula', input: inputOf(write(first, second)) };
    }
    case 'containsText': {
      const text = cfRule.attributes.text;
      return text === undefined || text === '' ? null : { kind: 'textContains', text };
    }
    case 'containsBlanks':
      return { kind: 'isEmpty' };
    case 'notContainsBlanks':
      return { kind: 'notEmpty' };
    default:
      // `expression`, and every named kind that carries the formula it
      // means. One without a formula is one this sheet cannot answer.
      return first === '' ? null : { kind: 'formula', input: inputOf(first) };
  }
}

/** An `rgb="FFRRGGBB"` colour as `#rrggbb`, or null for a theme colour a scale here cannot name. */
function rgbOf(element: XmlElement): string | null {
  const rgb = element.attributes.rgb;
  return rgb !== undefined && /^[0-9a-f]{8}$/i.test(rgb) ? `#${rgb.slice(2).toLowerCase()}` : null;
}

/**
 * The sheet's rules as `<conditionalFormatting>` blocks, one per rule,
 * with their paints' ids into the workbook's `dxfs`. `priority` counts
 * across the workbook, which is how Excel numbers it.
 */
export function writeConditionals(
  rules: readonly ConditionalRule[],
  dxfIdOf: (paint: ConditionalPaint) => number,
  excelOf: (input: string) => string,
  priority: { next: number }
): string {
  return rules
    .map(rule => {
      const sqref = sqrefOf(rule.range);
      const anchor = `${columnName(Math.min(rule.range.start.column, rule.range.end.column))}${Math.min(rule.range.start.row, rule.range.end.row) + 1}`;
      const at = priority.next++;
      if (rule.scale !== undefined) {
        const scale = rule.scale;
        const middle = scale.middle === undefined ? '' : '<cfvo type="percentile" val="50"/>';
        const colours = [scale.from, ...(scale.middle === undefined ? [] : [scale.middle]), scale.to]
          .map(colour => `<color rgb="${argbOf(colour)}"/>`)
          .join('');
        return `<conditionalFormatting sqref="${sqref}"><cfRule type="colorScale" priority="${at}"><colorScale><cfvo type="min"/>${middle}<cfvo type="max"/>${colours}</colorScale></cfRule></conditionalFormatting>`;
      }
      if (rule.test === null || rule.paint === undefined) {
        return '';
      }
      const dxf = dxfIdOf(rule.paint);
      const head = (type: string, extra = '') => `<cfRule type="${type}" dxfId="${dxf}" priority="${at}"${extra}>`;
      const formula = (text: string) => `<formula>${escapeText(text)}</formula>`;
      const test = rule.test;
      let body: string;
      switch (test.kind) {
        case 'greaterThan':
        case 'lessThan':
          body = `${head('cellIs', ` operator="${test.kind}"`)}${formula(String(test.value))}</cfRule>`;
          break;
        case 'between':
          body = `${head('cellIs', ' operator="between"')}${formula(String(test.low))}${formula(String(test.high))}</cfRule>`;
          break;
        case 'equalTo':
          body = `${head('cellIs', ' operator="equal"')}${formula(
            typeof test.value === 'number' ? String(test.value) : `"${test.value.replace(/"/g, '""')}"`
          )}</cfRule>`;
          break;
        case 'textContains':
          body = `${head('containsText', ` operator="containsText" text="${escapeAttribute(test.text)}"`)}${formula(
            `NOT(ISERROR(SEARCH("${test.text.replace(/"/g, '""')}",${anchor})))`
          )}</cfRule>`;
          break;
        case 'isEmpty':
          body = `${head('containsBlanks')}${formula(`LEN(TRIM(${anchor}))=0`)}</cfRule>`;
          break;
        case 'notEmpty':
          body = `${head('notContainsBlanks')}${formula(`LEN(TRIM(${anchor}))>0`)}</cfRule>`;
          break;
        case 'formula':
          body = `${head('expression')}${formula(excelOf(test.input))}</cfRule>`;
          break;
      }
      return `<conditionalFormatting sqref="${sqref}">${body}</conditionalFormatting>`;
    })
    .join('');
}

/** A paint as a differential format: only what it changes. */
export function dxfOf(paint: ConditionalPaint): string {
  const font = [
    paint.bold === true ? '<b/>' : '',
    paint.italic === true ? '<i/>' : '',
    paint.color === undefined || paint.color === '' ? '' : `<color rgb="${argbOf(paint.color)}"/>`
  ].join('');
  const fill =
    paint.fill === undefined || paint.fill === ''
      ? ''
      : `<fill><patternFill patternType="solid"><bgColor rgb="${argbOf(paint.fill)}"/></patternFill></fill>`;
  return `<dxf>${font === '' ? '' : `<font>${font}</font>`}${fill}</dxf>`;
}

/** `#rrggbb` as Excel's opaque `FFRRGGBB`. */
function argbOf(colour: string): string {
  const hex = /^#?([0-9a-f]{6})$/i.exec(colour.trim())?.[1] ?? '000000';
  return `FF${hex.toUpperCase()}`;
}
