import { formatNumber, isError, type CellValue } from './Values';

/**
 * What a cell looks like, and what its value reads as.
 *
 * Two halves that travel differently, which is why they are two
 * fields rather than twelve. The **number format** is applied on the
 * application worker, because turning 1234.5 into `$1,234.50` is the
 * last step before the wire and the render worker must never learn a
 * locale — a thousands separator computed on the frame path is
 * exactly the work this architecture exists to move off it. The
 * **paint** is the opposite: it is the only half the render worker
 * needs, and it crosses as a palette entry.
 *
 * Headless, like everything in `src/sheet`. `boundaries.spec.ts`
 * fails the build if this file ever reaches for the framework.
 */

/** How a value is turned into the text a person reads. */
export type NumberFormat =
  /** Whatever the value is, printed as itself. */
  | { readonly kind: 'general' }
  | { readonly kind: 'number'; readonly places: number; readonly thousands: boolean }
  | { readonly kind: 'currency'; readonly places: number; readonly symbol: string }
  | { readonly kind: 'percent'; readonly places: number }
  | { readonly kind: 'scientific'; readonly places: number }
  | { readonly kind: 'date'; readonly pattern: DatePattern }
  | { readonly kind: 'time'; readonly pattern: TimePattern }
  /**
   * Never a number, however much it looks like one.
   *
   * The only format that changes what *typing* means rather than what
   * showing means: `007` in a text cell stays `007`, and a phone
   * number keeps its leading zero. That is the whole reason the
   * format exists, and a Text format that only changed the display
   * would be a menu item that does nothing anybody wanted.
   */
  | { readonly kind: 'text' };

/**
 * Named rather than a pattern string.
 *
 * A pattern language — Excel's `dd/mm/yyyy` and its forty cousins —
 * is a parser, a spec and a class of bugs, and none of it is needed
 * to put a date on a screen. Three unambiguous orders cover what
 * people actually pick, and `2026-09-24` is never read as a day in
 * September or the ninth of the twenty-fourth month.
 */
export type DatePattern = 'ymd' | 'dmy' | 'mdy';
export type TimePattern = 'hm' | 'hms';

/**
 * One edge of a cell.
 *
 * A width and a colour rather than a boolean, because the useful
 * borders in a spreadsheet are not all the same: a hairline under a
 * header and a heavy rule above a total are the two everybody draws,
 * and a border model that could only say "on" would draw them the
 * same.
 *
 * Width 0 is no border at all, which is why this is not optional —
 * a cell always has four edges and most of them are nothing.
 */
export interface CellEdge {
  readonly width: number;
  /** A colour, or '' for the theme's own border colour. */
  readonly color: string;
}

export interface CellBorders {
  readonly top: CellEdge;
  readonly right: CellEdge;
  readonly bottom: CellEdge;
  readonly left: CellEdge;
}

export const NO_EDGE: CellEdge = { width: 0, color: '' };
export const NO_BORDERS: CellBorders = { top: NO_EDGE, right: NO_EDGE, bottom: NO_EDGE, left: NO_EDGE };

/** How a cell is painted. Everything here crosses to the render worker. */
export interface CellPaint {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  /** Points, or 0 for the sheet's own size. */
  readonly fontSize: number;
  /** A colour, or '' for the theme's text colour. */
  readonly color: string;
  /** A colour, or '' for no fill at all. */
  readonly fill: string;
  /**
   * `auto` is the spreadsheet rule: numbers right, everything else
   * left. It is a real alignment and not an absent one, because a
   * column of numbers that a person has not touched still has to line
   * up on its decimal point.
   */
  readonly align: 'auto' | 'start' | 'center' | 'end';
  readonly wrap: boolean;
  /**
   * The four edges.
   *
   * Drawn by the grid as four thin rectangles rather than as the
   * node's own border, because `borderWidth` in the engine is one
   * number for all four sides. A border there is paint-only — it
   * touches no layout — and the `decorated` modifier already takes a
   * list of arbitrary coloured rectangles drawn in the node's own
   * paint pass, so four edges cost four draw instances and *no extra
   * nodes*. See `CellBorders` in `Grid.tsx`.
   */
  readonly borders: CellBorders;
}

export interface CellFormat {
  readonly number: NumberFormat;
  readonly paint: CellPaint;
}

export const GENERAL: NumberFormat = { kind: 'general' };

export const PLAIN: CellPaint = {
  bold: false,
  italic: false,
  underline: false,
  fontSize: 0,
  color: '',
  fill: '',
  align: 'auto',
  wrap: false,
  borders: NO_BORDERS
};

export const DEFAULT_FORMAT: CellFormat = { number: GENERAL, paint: PLAIN };

/** Whether a format is the one every unformatted cell has. */
export function isDefault(format: CellFormat): boolean {
  return keyOf(format) === keyOf(DEFAULT_FORMAT);
}

/**
 * A format as one string, for interning.
 *
 * Built field by field rather than with `JSON.stringify`, because
 * stringify's output depends on the order the keys were written and
 * two formats that are equal would intern as two entries the first
 * time a field was set in a different order. A palette that grows a
 * duplicate per edit is a palette that is on the wire forever.
 */
export function keyOf(format: CellFormat): string {
  const n = format.number;
  const number =
    n.kind === 'number'
      ? `number:${n.places}:${n.thousands}`
      : n.kind === 'currency'
        ? `currency:${n.places}:${n.symbol}`
        : n.kind === 'percent'
          ? `percent:${n.places}`
          : n.kind === 'scientific'
            ? `scientific:${n.places}`
            : n.kind === 'date'
              ? `date:${n.pattern}`
              : n.kind === 'time'
                ? `time:${n.pattern}`
                : n.kind;
  const p = format.paint;
  return [
    number,
    p.bold ? 'b' : '',
    p.italic ? 'i' : '',
    p.underline ? 'u' : '',
    p.fontSize,
    p.color,
    p.fill,
    p.align,
    p.wrap ? 'w' : '',
    edgeKey(p.borders.top),
    edgeKey(p.borders.right),
    edgeKey(p.borders.bottom),
    edgeKey(p.borders.left)
  ].join('|');
}

function edgeKey(edge: CellEdge): string {
  return edge.width === 0 ? '' : `${edge.width}:${edge.color}`;
}

/**
 * What the screen shows for a value under a format.
 *
 * The rules that are not obvious, and that every spreadsheet shares:
 *
 *   - **An error ignores the format.** `#DIV/0!` under a currency
 *     format is `#DIV/0!` and not `$#DIV/0!`. The format describes a
 *     number and there is no number.
 *   - **Text ignores a numeric format.** A column formatted as
 *     currency with the word `Total` in it shows `Total`. Formatting
 *     a range must never make what is in it unreadable.
 *   - **An empty cell shows nothing**, whatever it is formatted as.
 *     A blank cell in a currency column is blank, not `$0.00`.
 */
export function formatWith(value: CellValue, format: NumberFormat): string {
  if (isError(value)) {
    return value.code;
  }
  if (value === null) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (typeof value === 'string') {
    return value;
  }
  switch (format.kind) {
    case 'general':
    case 'text':
      return formatNumber(value);
    case 'number':
      return fixed(value, format.places, format.thousands);
    case 'currency':
      return currency(value, format.places, format.symbol);
    case 'percent':
      return `${fixed(value * 100, format.places, false)}%`;
    case 'scientific':
      return scientific(value, format.places);
    case 'date':
      return date(value, format.pattern);
    case 'time':
      return time(value, format.pattern);
  }
}

/**
 * A number to a fixed number of places.
 *
 * `toFixed` rather than an intermediate rounding, because rounding to
 * two places and then printing can land on `1.005 → 1.00` by a route
 * that is harder to explain than the one every other spreadsheet
 * takes.
 */
function fixed(value: number, places: number, thousands: boolean): string {
  if (!Number.isFinite(value)) {
    return formatNumber(value);
  }
  const text = Math.abs(value).toFixed(clampPlaces(places));
  const sign = value < 0 || Object.is(value, -0) ? '-' : '';
  return sign + (thousands ? group(text) : text);
}

/**
 * Negative money in brackets, which is the accounting convention and
 * what a column of figures is read with.
 */
function currency(value: number, places: number, symbol: string): string {
  const body = symbol + fixed(Math.abs(value), places, true);
  return value < 0 ? `(${body})` : body;
}

function scientific(value: number, places: number): string {
  if (!Number.isFinite(value)) {
    return formatNumber(value);
  }
  // `toExponential` writes `1.50e+3`; spreadsheets write `1.50E+03`,
  // with the exponent padded to two digits.
  const [mantissa, exponent] = value.toExponential(clampPlaces(places)).split('e');
  const sign = exponent.startsWith('-') ? '-' : '+';
  const digits = exponent.replace(/[+-]/, '').padStart(2, '0');
  return `${mantissa}E${sign}${digits}`;
}

/** Thousands separators, applied to the integer part only. */
function group(text: string): string {
  const [whole, fraction] = text.split('.');
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return fraction === undefined ? grouped : `${grouped}.${fraction}`;
}

/** `toFixed` throws past 100, and nobody means it past about 15. */
function clampPlaces(places: number): number {
  return Math.min(Math.max(Math.trunc(places), 0), 15);
}

/**
 * The epoch spreadsheets count days from.
 *
 * 1899-12-30, not 1900-01-01, because Lotus 1-2-3 believed 1900 was a
 * leap year and every spreadsheet since has kept the bug so that
 * serial numbers move between them. Serial 45,000 is 2023-03-15 here
 * and in Excel, which is the only property worth having.
 *
 * The bug itself is a single phantom day: Excel's serial 60 is
 * 1900-02-29, a date that did not happen. Serials 1 to 59 are
 * therefore a day *ahead* of what this epoch alone would give, which
 * `SHIFT_BEFORE` puts back, and serial 60 is a day this application
 * declines to invent — it shows 1900-02-28, the same as 59. That is
 * one wrong day in 1900 against reproducing a calendar error on
 * purpose, and it is the trade every implementation has to pick.
 */
const EPOCH_UTC = Date.UTC(1899, 11, 30);
/** Serials below this predate the phantom day and need the shift. */
const SHIFT_BEFORE = 60;
const DAY_MS = 86_400_000;

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * A serial number as a date.
 *
 * Built in UTC throughout. A local-time construction would shift the
 * displayed day for anybody west of Greenwich for part of the year,
 * so a sheet would show a different date to two people looking at the
 * same file — which is the one thing a date must never do.
 *
 * Phase 9 can *show* a date and cannot yet *make* one: nothing parses
 * `2026-09-24` into a serial, and no function returns one. That is
 * Phase 11, and it is the single point where the engine and this file
 * meet.
 */
function date(serial: number, pattern: DatePattern): string {
  if (!Number.isFinite(serial)) {
    return formatNumber(serial);
  }
  const days = Math.floor(serial);
  const at = new Date(EPOCH_UTC + (days < SHIFT_BEFORE ? days + 1 : days) * DAY_MS);
  const year = at.getUTCFullYear();
  const month = at.getUTCMonth();
  const day = at.getUTCDate();
  switch (pattern) {
    case 'ymd':
      return `${pad(year, 4)}-${pad(month + 1, 2)}-${pad(day, 2)}`;
    case 'dmy':
      return `${day} ${MONTHS[month]} ${year}`;
    case 'mdy':
      return `${MONTHS[month]} ${day}, ${year}`;
  }
}

/** The fraction of a day as a clock time. */
function time(serial: number, pattern: TimePattern): string {
  if (!Number.isFinite(serial)) {
    return formatNumber(serial);
  }
  // Rounded to the second before splitting, so 23:59:59.7 reads as
  // the next midnight rather than as 23:59:60.
  const ofDay = serial - Math.floor(serial);
  const seconds = Math.round(ofDay * 86_400) % 86_400;
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  return pattern === 'hm'
    ? `${pad(hours, 2)}:${pad(minutes, 2)}`
    : `${pad(hours, 2)}:${pad(minutes, 2)}:${pad(seconds % 60, 2)}`;
}

function pad(value: number, width: number): string {
  return String(value).padStart(width, '0');
}

/** More decimal places, or fewer, for the two toolbar buttons. */
export function withPlaces(format: NumberFormat, by: number): NumberFormat {
  switch (format.kind) {
    case 'number':
    case 'percent':
    case 'scientific':
      return { ...format, places: clampPlaces(format.places + by) };
    case 'currency':
      return { ...format, places: clampPlaces(format.places + by) };
    /**
     * General has no places to add to, so the first press turns it
     * into a number format that has. Pressing "more decimals" on an
     * untouched cell is how most people first format anything, and a
     * button that did nothing there would be the wrong answer to the
     * commonest use of it.
     */
    case 'general':
      return by > 0 ? { kind: 'number', places: 1, thousands: false } : format;
    default:
      return format;
  }
}

/** How many places a format shows, for a toolbar that wants to say. */
export function placesOf(format: NumberFormat): number | null {
  switch (format.kind) {
    case 'number':
    case 'currency':
    case 'percent':
    case 'scientific':
      return format.places;
    default:
      return null;
  }
}
