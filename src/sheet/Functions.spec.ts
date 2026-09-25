import { describe, expect, it } from 'vitest';

import { serialOfDate, serialOfTime } from './Dates';
import { functionNames } from './Functions';
import { Sheet } from './Sheet';
import { Workbook } from './Workbook';
import { formatValue, type CellValue } from './Values';

/**
 * The conformance table: Phase 11's exit criterion.
 *
 * Every function the sheet knows, with its answer asserted as a
 * **literal** rather than as a recomputation. A table that said
 * `expect(SUM(a, b)).toBe(a + b)` would pass for a `SUM` that
 * multiplied, and this file exists precisely to be the place that
 * cannot happen.
 *
 * It runs through the whole engine — parse, graph, recalculate, read —
 * rather than against the function table directly, because that is how
 * a formula reaches a function in the application and the wiring
 * between them is where most of the interesting mistakes live.
 *
 * The last spec in this file fails the build if a name is added to the
 * library with no case here.
 */

/** A sheet with a small table in it, which the cases refer to. */
function sheetWith(cells: Readonly<Record<string, string>>): Sheet {
  const sheet = new Workbook().sheet(0);
  for (const [address, input] of Object.entries(cells)) {
    const match = /^([A-Z]+)(\d+)$/.exec(address)!;
    const column = match[1].split('').reduce((at, letter) => at * 26 + (letter.charCodeAt(0) - 64), 0) - 1;
    sheet.setCell(Number(match[2]) - 1, column, input);
  }
  return sheet;
}

/**
 * The fixture every case is evaluated against.
 *
 *   A: numbers      B: text        C: mixed / blanks
 *   1: 10           North          3
 *   2: 20           South          (blank)
 *   3: 30           East           7
 *   4: 40           North          1
 *   5: 5            west           9
 */
const FIXTURE: Readonly<Record<string, string>> = {
  A1: '10',
  A2: '20',
  A3: '30',
  A4: '40',
  A5: '5',
  B1: 'North',
  B2: 'South',
  B3: 'East',
  B4: 'North',
  B5: 'west',
  C1: '3',
  C3: '7',
  C4: '1',
  C5: '9',
  D1: '2026-09-24',
  D2: '2026-01-31',
  E1: '=1/0'
};

/** The answer a formula gives, as the text a cell would show. */
function answer(formula: string, extra: Readonly<Record<string, string>> = {}): string {
  const sheet = sheetWith({ ...FIXTURE, ...extra });
  // A fixed clock, so `TODAY` and `NOW` can be asserted at all.
  sheet.clock = () => serialOfDate(2026, 9, 24) + serialOfTime(14, 30, 0);
  sheet.dice = () => 0.5;
  sheet.setCell(20, 20, `=${formula}`);
  sheet.recalculate();
  return formatValue(sheet.value(20, 20));
}

/** Which functions a table has a case for, so the guard can check. */
const covered = new Set<string>();

function table(family: string, cases: readonly (readonly [string, string])[]): void {
  describe(family, () => {
    for (const [formula, expected] of cases) {
      const name = /^([A-Z][A-Z0-9.]*)\(/.exec(formula)?.[1];
      if (name !== undefined) {
        covered.add(name);
      }
      it(`${formula} is ${expected}`, () => {
        expect(answer(formula)).toBe(expected);
      });
    }
  });
}

table('aggregates', [
  ['SUM(A1:A5)', '105'],
  ['SUM(A1, 5)', '15'],
  ['SUM(A1:A5, 100)', '205'],
  // Text in a range is skipped; text passed directly is a claim.
  ['SUM(B1:B5)', '0'],
  ['SUM("x")', '#VALUE!'],
  ['SUM(E1)', '#DIV/0!'],
  ['AVERAGE(A1:A5)', '21'],
  ['AVERAGE(B1:B5)', '#DIV/0!'],
  ['MIN(A1:A5)', '5'],
  ['MAX(A1:A5)', '40'],
  ['COUNT(A1:A5)', '5'],
  ['COUNT(C1:C5)', '4'],
  // Counting survives a broken cell; summing does not.
  ['COUNT(E1)', '0'],
  ['ROUND(2.5, 0)', '3'],
  ['ROUND(-2.5, 0)', '-3'],
  ['ROUND(3.14159, 2)', '3.14'],
  ['ABS(-7)', '7'],
  ['CONCAT(B1, "-", B2)', 'North-South']
]);

table('logic', [
  ['AND(TRUE, TRUE)', 'TRUE'],
  ['AND(TRUE, FALSE)', 'FALSE'],
  ['AND(A1>5, A2>5)', 'TRUE'],
  ['OR(FALSE, TRUE)', 'TRUE'],
  ['OR(FALSE, FALSE)', 'FALSE'],
  ['XOR(TRUE, TRUE)', 'FALSE'],
  ['XOR(TRUE, FALSE)', 'TRUE'],
  ['NOT(TRUE)', 'FALSE'],
  ['IF(A1>5, "big", "small")', 'big'],
  ['IF(A1>50, "big", "small")', 'small'],
  // The whole reason `IF` is lazy: the branch not taken is not run.
  ['IF(A1>5, "safe", 1/0)', 'safe'],
  ['IFS(A1>50, "huge", A1>5, "big")', 'big'],
  ['IFS(A1>50, "huge", A1>500, "vast")', '#N/A'],
  ['IFS(A1>50, 1/0, A1>5, "big")', 'big'],
  ['SWITCH(2, 1, "one", 2, "two", "other")', 'two'],
  ['SWITCH(9, 1, "one", 2, "two", "other")', 'other'],
  ['SWITCH(9, 1, "one", 2, "two")', '#N/A'],
  ['IFERROR(1/0, "oops")', 'oops'],
  ['IFERROR(A1, "oops")', '10'],
  ['IFNA(NA(), "missing")', 'missing'],
  // `IFNA` catches only the one code, which is the point of it.
  ['IFNA(1/0, "missing")', '#DIV/0!'],
  ['ISERROR(E1)', 'TRUE'],
  ['ISERROR(A1)', 'FALSE'],
  ['ISBLANK(C2)', 'TRUE'],
  ['ISBLANK(C1)', 'FALSE'],
  ['ISNUMBER(A1)', 'TRUE'],
  ['ISNUMBER(B1)', 'FALSE'],
  ['ISTEXT(B1)', 'TRUE'],
  ['TRUE()', 'TRUE'],
  ['FALSE()', 'FALSE'],
  ['NA()', '#N/A']
]);

table('maths', [
  ['SQRT(16)', '4'],
  ['SQRT(-1)', '#VALUE!'],
  ['POWER(2, 10)', '1024'],
  ['MOD(7, 3)', '1'],
  // The sign follows the divisor, which JavaScript's `%` does not.
  ['MOD(-3, 2)', '1'],
  ['MOD(3, -2)', '-1'],
  ['MOD(7, 0)', '#DIV/0!'],
  ['INT(-2.5)', '-3'],
  ['TRUNC(-2.5)', '-2'],
  ['TRUNC(3.14159, 2)', '3.14'],
  ['CEILING(4.2, 1)', '5'],
  ['CEILING(4.2, 0.5)', '4.5'],
  ['FLOOR(4.7, 1)', '4'],
  ['FLOOR(-4.2, 1)', '-4'],
  ['SIGN(-9)', '-1'],
  ['SIGN(0)', '0'],
  ['EXP(0)', '1'],
  ['LN(1)', '0'],
  ['LN(0)', '#VALUE!'],
  ['LOG(100)', '2'],
  ['LOG(8, 2)', '3'],
  ['LOG10(1000)', '3'],
  ['ROUNDUP(2.1, 0)', '3'],
  ['ROUNDUP(-2.1, 0)', '-3'],
  ['ROUNDDOWN(2.9, 0)', '2'],
  ['SUMPRODUCT(A1:A3, C1:C3)', '240'],
  ['SUMPRODUCT(A1:A3, C1:C4)', '#VALUE!'],
  ['FACT(5)', '120'],
  ['FACT(0)', '1'],
  ['PRODUCT(A1:A3)', '6000'],
  ['PI()', '3.14159265358979'],
  // Fixed dice, so the answer can be asserted rather than described.
  ['RAND()', '0.5'],
  ['RANDBETWEEN(1, 10)', '6']
]);

table('statistics', [
  ['MEDIAN(A1:A5)', '20'],
  ['MEDIAN(A1:A4)', '25'],
  ['MODE(C1:C5)', '#N/A'],
  ['MODE(A1:A5, 10)', '10'],
  ['STDEV(A1:A5)', '14.3178210632764'],
  ['VAR(A1:A5)', '205'],
  ['STDEV(A1)', '#DIV/0!'],
  ['STDEVP(A1:A5)', '12.8062484748657'],
  ['VARP(A1:A5)', '164'],
  ['COUNTA(B1:B5)', '5'],
  ['COUNTA(C1:C5)', '4'],
  ['COUNTBLANK(C1:C5)', '1'],
  ['LARGE(A1:A5, 1)', '40'],
  ['LARGE(A1:A5, 2)', '30'],
  ['SMALL(A1:A5, 1)', '5'],
  ['LARGE(A1:A5, 9)', '#N/A'],
  // Rank one is the largest, which is the spreadsheet convention.
  ['RANK(30, A1:A5)', '2'],
  ['RANK(30, A1:A5, 1)', '4'],
  ['RANK(99, A1:A5)', '#N/A'],
  ['PERCENTILE(A1:A5, 0.5)', '20'],
  ['PERCENTILE(A1:A5, 0)', '5'],
  ['PERCENTILE(A1:A5, 1)', '40'],
  ['QUARTILE(A1:A5, 2)', '20']
]);

table('conditional aggregates', [
  ['SUMIF(A1:A5, ">15")', '90'],
  ['SUMIF(B1:B5, "North", A1:A5)', '50'],
  // Case-insensitive, like every other comparison in the engine.
  ['SUMIF(B1:B5, "WEST", A1:A5)', '5'],
  ['SUMIF(B1:B5, "N*", A1:A5)', '50'],
  ['COUNTIF(A1:A5, ">15")', '3'],
  ['COUNTIF(B1:B5, "North")', '2'],
  ['COUNTIF(C1:C5, "<>")', '4'],
  ['AVERAGEIF(A1:A5, ">15")', '30'],
  ['AVERAGEIF(A1:A5, ">100")', '#DIV/0!'],
  ['SUMIFS(A1:A5, B1:B5, "North")', '50'],
  ['SUMIFS(A1:A5, B1:B5, "North", A1:A5, ">15")', '40'],
  ['COUNTIFS(B1:B5, "North", A1:A5, ">15")', '1'],
  ['AVERAGEIFS(A1:A5, B1:B5, "North")', '25'],
  ['MAXIFS(A1:A5, B1:B5, "North")', '40'],
  ['MINIFS(A1:A5, B1:B5, "North")', '10']
]);

table('text', [
  ['LEFT("spreadsheet", 6)', 'spread'],
  ['LEFT("ab", 9)', 'ab'],
  ['RIGHT("spreadsheet", 5)', 'sheet'],
  ['RIGHT("ab", 0)', ''],
  ['MID("spreadsheet", 7, 5)', 'sheet'],
  ['MID("abc", 0, 2)', '#VALUE!'],
  ['LEN("sheet")', '5'],
  ['LEN(A1)', '2'],
  ['FIND("sheet", "spreadsheet")', '7'],
  ['FIND("SHEET", "spreadsheet")', '#VALUE!'],
  ['SEARCH("SHEET", "spreadsheet")', '7'],
  ['TRIM("  a   b  ")', 'a b'],
  ['UPPER("north")', 'NORTH'],
  ['LOWER("North")', 'north'],
  ['PROPER("north by northwest")', 'North By Northwest'],
  ['SUBSTITUTE("a-b-c", "-", "+")', 'a+b+c'],
  ['SUBSTITUTE("a-b-c", "-", "+", 2)', 'a-b+c'],
  ['REPLACE("abcdef", 2, 3, "X")', 'aXef'],
  ['REPT("ab", 3)', 'ababab'],
  ['TEXTJOIN(", ", TRUE, B1, "", B2)', 'North, South'],
  ['TEXTJOIN(", ", FALSE, B1, "", B2)', 'North, , South'],
  ['VALUE("42")', '42'],
  ['VALUE("x")', '#VALUE!'],
  ['TEXT(1234.5, "#,##0.00")', '1,234.50'],
  ['TEXT(0.25, "0%")', '25%'],
  ['TEXT(D1, "YYYY-MM-DD")', '2026-09-24'],
  ['TEXT(1, "nonsense")', '#VALUE!'],
  ['CHAR(65)', 'A'],
  ['CODE("A")', '65'],
  ['EXACT("a", "a")', 'TRUE'],
  ['EXACT("a", "A")', 'FALSE'],
  ['CONCATENATE(B1, "-", B2)', 'North-South']
]);

table('lookup', [
  ['VLOOKUP(30, A1:B5, 2, FALSE)', 'East'],
  ['VLOOKUP(99, A1:B5, 2, FALSE)', '#N/A'],
  ['VLOOKUP(30, A1:B5, 3, FALSE)', '#VALUE!'],
  // Approximate on unsorted data: a wrong answer, quietly, which is
  // exactly why the exact mode is the one to reach for.
  ['VLOOKUP(35, A1:B5, 2)', 'East'],
  ['HLOOKUP(10, A1:C1, 1, FALSE)', '10'],
  ['INDEX(A1:A5, 3)', '30'],
  ['INDEX(A1:B5, 3, 2)', 'East'],
  ['INDEX(A1:A5, 9)', '#VALUE!'],
  ['MATCH(30, A1:A5, 0)', '3'],
  ['MATCH(99, A1:A5, 0)', '#N/A'],
  ['XLOOKUP(30, A1:A5, B1:B5)', 'East'],
  ['XLOOKUP(99, A1:A5, B1:B5, "none")', 'none'],
  ['XLOOKUP(99, A1:A5, B1:B5)', '#N/A'],
  ['CHOOSE(2, "a", "b", "c")', 'b'],
  ['CHOOSE(9, "a", "b")', '#VALUE!'],
  ['ROWS(A1:A5)', '5'],
  ['COLUMNS(A1:C1)', '3'],
  ['INDIRECT("A3")', '30'],
  ['INDIRECT("A" & 3)', '30'],
  ['INDIRECT("nonsense")', '#REF!'],
  ['SUM(OFFSET(A1, 0, 0, 3, 1))', '60'],
  ['OFFSET(A1, 2, 0)', '30'],
  ['OFFSET(A1, -5, 0)', '#REF!']
]);

table('dates', [
  ['TODAY()', '46289'],
  ['NOW()', '46289.6041666667'],
  ['DATE(2026, 9, 24)', '46289'],
  ['DATE(2026, 13, 1)', '46388'],
  ['TIME(9, 30, 0)', '0.395833333333333'],
  ['YEAR(D1)', '2026'],
  ['MONTH(D1)', '9'],
  ['DAY(D1)', '24'],
  ['HOUR(0.5)', '12'],
  ['MINUTE(0.395833333333333)', '30'],
  ['SECOND(0.5)', '0'],
  // 2026-09-24 is a Thursday.
  ['WEEKDAY(D1)', '5'],
  ['WEEKDAY(D1, 2)', '4'],
  // The 31st of January plus one month is the 28th of February.
  ['EDATE(D2, 1)', '46081'],
  ['EOMONTH(D2, 0)', '46053'],
  ['EOMONTH(D2, 1)', '46081'],
  ['DATEDIF(D2, D1, "D")', '236'],
  ['DATEDIF(D2, D1, "M")', '7'],
  ['DATEDIF(D2, D1, "Y")', '0'],
  ['DATEDIF(D2, D1, "YM")', '7'],
  ['DATEDIF(D1, D2, "D")', '#VALUE!'],
  ['NETWORKDAYS(D2, D1)', '169']
]);

/**
 * The guard that keeps this file honest.
 *
 * A function added to the library with no case above is a function
 * nobody has ever asserted the behaviour of, and the point of a
 * conformance table is that there is no such thing.
 */
describe('the table covers the library', () => {
  it('has a case for every function the sheet knows', () => {
    const missing = functionNames().filter(name => !covered.has(name));
    expect(missing).toEqual([]);
  });

  it('has enough cases to be a conformance table', () => {
    // A guard against the guard passing because the table is empty.
    expect(covered.size).toBeGreaterThan(70);
  });
});

/** An unknown name is `#NAME?`, which is how a typo reads. */
describe('a name the sheet does not know', () => {
  it('is #NAME?', () => {
    expect(answer('NOSUCHFUNCTION(1)')).toBe('#NAME?');
  });

  it('is #NAME? as a bare word too', () => {
    expect(answer('NOSUCHNAME')).toBe('#NAME?');
  });
});

/** Errors travel, rather than being swallowed by the function above. */
describe('an error in an argument', () => {
  const cases: readonly string[] = [
    'SQRT(E1)',
    'LEFT(E1, 1)',
    'VLOOKUP(E1, A1:B5, 2, FALSE)',
    'YEAR(E1)',
    'AND(E1)',
    'MEDIAN(E1)'
  ];
  it.each(cases)('travels out of %s', formula => {
    expect(answer(formula)).toBe('#DIV/0!');
  });
});

/** What a cell shows is the value; these assert the value is right. */
function valueOf(formula: string): CellValue {
  const sheet = sheetWith(FIXTURE);
  sheet.setCell(20, 20, `=${formula}`);
  sheet.recalculate();
  return sheet.value(20, 20);
}

describe('the kinds values come back as', () => {
  it('returns numbers as numbers, not as text', () => {
    expect(valueOf('SUM(A1:A5)')).toBe(105);
  });

  it('returns booleans as booleans', () => {
    expect(valueOf('AND(TRUE, TRUE)')).toBe(true);
  });

  it('returns text as text', () => {
    expect(valueOf('UPPER("a")')).toBe('A');
  });
});
