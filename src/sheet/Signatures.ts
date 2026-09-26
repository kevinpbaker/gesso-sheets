/**
 * What each function takes, for the editor to show while it is typed.
 *
 * A separate table from the implementations, and that is a real cost:
 * two places to change when a function's arguments change, and a spec
 * at the bottom of `Signatures.spec.ts` that fails when one of them
 * is forgotten. The alternative is metadata on every entry in
 * `Functions.ts`, which would put prose about `SUMIF`'s argument
 * order into the file that has to stay readable as *arithmetic*.
 *
 * The names are what the editor shows, so they are written for
 * somebody reading a popup at speed: `range`, `criteria`, `[places]`
 * — square brackets meaning optional, `…` meaning the last one
 * repeats. That is the notation every spreadsheet's formula hint uses
 * and the one people already read.
 */

export interface Signature {
  /** The arguments, in order, as the hint shows them. */
  readonly args: readonly string[];
  /** One line, shown under the arguments. */
  readonly summary: string;
  /**
   * Whether the last argument repeats.
   *
   * `SUM(number, …)` highlights its last argument for every argument
   * past the first, rather than running out of things to point at.
   */
  readonly repeats?: boolean;
}

export const SIGNATURES: Readonly<Record<string, Signature>> = {
  // Aggregates
  SUM: { args: ['number', '…'], summary: 'Adds up the numbers.', repeats: true },
  AVERAGE: { args: ['number', '…'], summary: 'The mean of the numbers.', repeats: true },
  MIN: { args: ['number', '…'], summary: 'The smallest number.', repeats: true },
  MAX: { args: ['number', '…'], summary: 'The largest number.', repeats: true },
  COUNT: { args: ['value', '…'], summary: 'How many of these are numbers.', repeats: true },
  PRODUCT: { args: ['number', '…'], summary: 'The numbers multiplied together.', repeats: true },
  ROUND: { args: ['number', '[places]'], summary: 'Rounds, away from zero on a tie.' },
  ROUNDUP: { args: ['number', '[places]'], summary: 'Rounds away from zero.' },
  ROUNDDOWN: { args: ['number', '[places]'], summary: 'Rounds towards zero.' },
  ABS: { args: ['number'], summary: 'The number without its sign.' },
  CONCAT: { args: ['text', '…'], summary: 'Joins the text together.', repeats: true },
  CONCATENATE: { args: ['text', '…'], summary: 'Joins the text together.', repeats: true },

  // Logic
  ISERR: { args: ['value'], summary: 'Whether it is an error other than #N/A.' },
  ISNA: { args: ['value'], summary: 'Whether it is #N/A.' },
  ISLOGICAL: { args: ['value'], summary: 'Whether it is TRUE or FALSE.' },
  ISNONTEXT: { args: ['value'], summary: 'Whether it is anything but text.' },
  N: { args: ['value'], summary: 'It as a number: TRUE is 1, text is 0.' },
  T: { args: ['value'], summary: 'It if it is text, and nothing if not.' },
  IF: { args: ['test', 'then', '[otherwise]'], summary: 'One value or the other.' },
  IFS: { args: ['test', 'then', '…'], summary: 'The first test that passes.', repeats: true },
  SWITCH: { args: ['value', 'case', 'then', '…'], summary: 'Matches a value against cases.', repeats: true },
  IFERROR: { args: ['value', 'otherwise'], summary: 'The value, or the other one if it broke.' },
  IFNA: { args: ['value', 'otherwise'], summary: 'The value, or the other one if it is #N/A.' },
  AND: { args: ['test', '…'], summary: 'True when all of them are.', repeats: true },
  OR: { args: ['test', '…'], summary: 'True when any of them is.', repeats: true },
  XOR: { args: ['test', '…'], summary: 'True when an odd number of them are.', repeats: true },
  NOT: { args: ['test'], summary: 'The opposite.' },
  TRUE: { args: [], summary: 'True.' },
  FALSE: { args: [], summary: 'False.' },
  NA: { args: [], summary: '#N/A, written deliberately.' },
  ISERROR: { args: ['value'], summary: 'Whether it is an error.' },
  ISBLANK: { args: ['value'], summary: 'Whether the cell is empty.' },
  ISNUMBER: { args: ['value'], summary: 'Whether it is a number.' },
  ISTEXT: { args: ['value'], summary: 'Whether it is text.' },

  // Maths
  ISEVEN: { args: ['number'], summary: 'Whether it is even.' },
  ISODD: { args: ['number'], summary: 'Whether it is odd.' },
  MROUND: { args: ['number', 'multiple'], summary: 'Rounds to the nearest multiple.' },
  QUOTIENT: { args: ['numerator', 'denominator'], summary: 'The whole part of a division.' },
  EVEN: { args: ['number'], summary: 'Rounds away from zero to an even number.' },
  ODD: { args: ['number'], summary: 'Rounds away from zero to an odd number.' },
  ROW: { args: ['[reference]'], summary: 'The row number, of a reference or of this cell.' },
  COLUMN: { args: ['[reference]'], summary: 'The column number, of a reference or of this cell.' },
  SUBTOTAL: { args: ['function', 'range', '…'], summary: 'An aggregate by number, 9 being SUM, skipping filtered rows.', repeats: true },
  PMT: { args: ['rate', 'periods', 'present', '[future]', '[due]'], summary: 'The payment each period on a loan.' },
  FV: { args: ['rate', 'periods', 'payment', '[present]', '[due]'], summary: 'What payments will be worth at the end.' },
  PV: { args: ['rate', 'periods', 'payment', '[future]', '[due]'], summary: 'What payments are worth now.' },
  NPV: { args: ['rate', 'flow', '…'], summary: 'Cash flows discounted to now.', repeats: true },
  SQRT: { args: ['number'], summary: 'The square root.' },
  POWER: { args: ['number', 'exponent'], summary: 'The number raised to a power.' },
  MOD: { args: ['number', 'divisor'], summary: 'The remainder, signed like the divisor.' },
  INT: { args: ['number'], summary: 'Rounds down, towards minus infinity.' },
  TRUNC: { args: ['number', '[places]'], summary: 'Cuts the decimals off, towards zero.' },
  CEILING: { args: ['number', '[step]'], summary: 'Up to the next multiple.' },
  FLOOR: { args: ['number', '[step]'], summary: 'Down to the multiple below.' },
  SIGN: { args: ['number'], summary: 'Minus one, zero or one.' },
  EXP: { args: ['number'], summary: 'e raised to the number.' },
  LN: { args: ['number'], summary: 'The natural logarithm.' },
  LOG: { args: ['number', '[base]'], summary: 'The logarithm, base ten unless told.' },
  LOG10: { args: ['number'], summary: 'The logarithm, base ten.' },
  FACT: { args: ['number'], summary: 'The factorial.' },
  PI: { args: [], summary: 'π.' },
  RAND: { args: [], summary: 'A number between zero and one. Recalculates on every edit.' },
  RANDBETWEEN: { args: ['low', 'high'], summary: 'A whole number in the range. Recalculates on every edit.' },
  SUMPRODUCT: { args: ['range', '…'], summary: 'The ranges multiplied cell by cell, then added.', repeats: true },

  // Statistics
  MEDIAN: { args: ['number', '…'], summary: 'The middle value.', repeats: true },
  MODE: { args: ['number', '…'], summary: 'The value that appears most often.', repeats: true },
  STDEV: { args: ['number', '…'], summary: 'The standard deviation of a sample.', repeats: true },
  STDEVP: { args: ['number', '…'], summary: 'The standard deviation of a population.', repeats: true },
  VAR: { args: ['number', '…'], summary: 'The variance of a sample.', repeats: true },
  VARP: { args: ['number', '…'], summary: 'The variance of a population.', repeats: true },
  COUNTA: { args: ['value', '…'], summary: 'How many cells hold anything.', repeats: true },
  COUNTBLANK: { args: ['range', '…'], summary: 'How many cells hold nothing.', repeats: true },
  LARGE: { args: ['range', 'k'], summary: 'The k-th largest value.' },
  SMALL: { args: ['range', 'k'], summary: 'The k-th smallest value.' },
  RANK: { args: ['number', 'range', '[ascending]'], summary: 'Where it comes in the list; one is largest.' },
  PERCENTILE: { args: ['range', 'fraction'], summary: 'The value that far through the data.' },
  QUARTILE: { args: ['range', 'quarter'], summary: 'The value at a quarter of the way through.' },

  // Conditional aggregates
  SUMIF: { args: ['range', 'criteria', '[sum range]'], summary: 'Adds the cells that match.' },
  SUMIFS: { args: ['sum range', 'range', 'criteria', '…'], summary: 'Adds the cells matching every pair.', repeats: true },
  COUNTIF: { args: ['range', 'criteria'], summary: 'Counts the cells that match.' },
  COUNTIFS: { args: ['range', 'criteria', '…'], summary: 'Counts the cells matching every pair.', repeats: true },
  AVERAGEIF: { args: ['range', 'criteria', '[average range]'], summary: 'The mean of the cells that match.' },
  AVERAGEIFS: {
    args: ['average range', 'range', 'criteria', '…'],
    summary: 'The mean of the cells matching every pair.',
    repeats: true
  },
  MAXIFS: { args: ['range', 'range', 'criteria', '…'], summary: 'The largest of the cells that match.', repeats: true },
  MINIFS: { args: ['range', 'range', 'criteria', '…'], summary: 'The smallest of the cells that match.', repeats: true },

  // Text
  CLEAN: { args: ['text'], summary: 'Takes out the control characters.' },
  HYPERLINK: { args: ['address', '[label]'], summary: 'Shows the label, or the address.' },
  LEFT: { args: ['text', '[count]'], summary: 'The first characters.' },
  RIGHT: { args: ['text', '[count]'], summary: 'The last characters.' },
  MID: { args: ['text', 'start', 'count'], summary: 'Characters from the middle, counting from one.' },
  LEN: { args: ['text'], summary: 'How many characters.' },
  FIND: { args: ['needle', 'haystack', '[start]'], summary: 'Where it is, minding case.' },
  SEARCH: { args: ['needle', 'haystack', '[start]'], summary: 'Where it is, ignoring case.' },
  TRIM: { args: ['text'], summary: 'Spaces off the ends and runs inside collapsed.' },
  UPPER: { args: ['text'], summary: 'In capitals.' },
  LOWER: { args: ['text'], summary: 'In lower case.' },
  PROPER: { args: ['text'], summary: 'Each word capitalised.' },
  SUBSTITUTE: { args: ['text', 'from', 'to', '[which]'], summary: 'Replaces text by what it says.' },
  REPLACE: { args: ['text', 'start', 'count', 'with'], summary: 'Replaces text by where it is.' },
  REPT: { args: ['text', 'times'], summary: 'The text repeated.' },
  TEXTJOIN: { args: ['separator', 'skip empty', 'text', '…'], summary: 'Joins with a separator.', repeats: true },
  VALUE: { args: ['text'], summary: 'Text read back as a number.' },
  TEXT: { args: ['value', 'format'], summary: 'A value as text, in a named format.' },
  CHAR: { args: ['code'], summary: 'The character with that code.' },
  CODE: { args: ['text'], summary: 'The code of the first character.' },
  EXACT: { args: ['text', 'text'], summary: 'Whether they are the same, minding case.' },

  // Lookup
  VLOOKUP: {
    args: ['value', 'table', 'column', '[approximate]'],
    summary: 'Finds a row by its first column. Exact needs FALSE.'
  },
  HLOOKUP: { args: ['value', 'table', 'row', '[approximate]'], summary: 'Finds a column by its first row.' },
  INDEX: { args: ['range', 'row', '[column]'], summary: 'The cell at a position, counting from one.' },
  MATCH: { args: ['value', 'range', '[mode]'], summary: 'Where a value is. Zero means exact.' },
  XLOOKUP: {
    args: ['value', 'searched', 'returned', '[if missing]', '[mode]'],
    summary: 'Finds a value. Exact by default.'
  },
  CHOOSE: { args: ['which', 'value', '…'], summary: 'The n-th of these, counting from one.', repeats: true },
  ROWS: { args: ['range'], summary: 'How many rows it covers.' },
  COLUMNS: { args: ['range'], summary: 'How many columns it covers.' },
  INDIRECT: { args: ['address'], summary: 'The cell an address names, worked out as it runs.' },
  OFFSET: {
    args: ['from', 'rows', 'columns', '[height]', '[width]'],
    summary: 'A range moved, and resized, from a corner.'
  },

  // Dates
  DATEVALUE: { args: ['text'], summary: 'A date typed as text, as a date.' },
  TIMEVALUE: { args: ['text'], summary: 'A time typed as text, as a time.' },
  DAYS: { args: ['end', 'start'], summary: 'Days from the start to the end.' },
  WORKDAY: { args: ['start', 'days', '[holidays]'], summary: 'The date that many working days away.' },
  WEEKNUM: { args: ['date', '[type]'], summary: 'The week of the year: 1 starts on Sunday, 2 on Monday.' },
  TODAY: { args: [], summary: "Today's date. Recalculates on every edit." },
  NOW: { args: [], summary: 'The date and time. Recalculates on every edit.' },
  DATE: { args: ['year', 'month', 'day'], summary: 'A date from its parts.' },
  TIME: { args: ['hours', 'minutes', 'seconds'], summary: 'A time of day.' },
  YEAR: { args: ['date'], summary: 'The year.' },
  MONTH: { args: ['date'], summary: 'The month, one to twelve.' },
  DAY: { args: ['date'], summary: 'The day of the month.' },
  HOUR: { args: ['time'], summary: 'The hour.' },
  MINUTE: { args: ['time'], summary: 'The minutes.' },
  SECOND: { args: ['time'], summary: 'The seconds.' },
  WEEKDAY: { args: ['date', '[type]'], summary: 'The day of the week; Sunday is one.' },
  EDATE: { args: ['date', 'months'], summary: 'The same day, months away.' },
  EOMONTH: { args: ['date', 'months'], summary: 'The last day of the month, months away.' },
  NETWORKDAYS: { args: ['from', 'to', '[holidays]'], summary: 'Working days between two dates.' },
  DATEDIF: { args: ['from', 'to', 'unit'], summary: 'The gap in whole units: "Y", "M", "D", "YM", "MD", "YD".' }
};

/** The signature for a name, or null when the sheet has no such function. */
export function signatureOf(name: string): Signature | null {
  return SIGNATURES[name.toUpperCase()] ?? null;
}

/**
 * The names that start with what somebody has typed.
 *
 * Prefix rather than substring, because a list that offered `COUNTIF`
 * for `IF` would bury the function actually being typed under its
 * cousins. Sorted, so the same prefix always offers the same list in
 * the same order and the third item stays the third item.
 */
export function completionsFor(prefix: string): readonly string[] {
  if (prefix === '') {
    return [];
  }
  const upper = prefix.toUpperCase();
  return Object.keys(SIGNATURES)
    .filter(name => name.startsWith(upper))
    .sort();
}
