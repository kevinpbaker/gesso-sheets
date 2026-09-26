import { describe, expect, it } from 'vitest';

import { readXlsx, type XlsxBook } from './Xlsx';

/**
 * Validations read from an `.xlsx` as Excel writes them — Phase 23.
 *
 * The XML here is the shape Excel's own files have, from POI's corpus:
 * a list typed into the rule and a list that names a range on another
 * sheet, the operators as Excel spells them, and the kinds this sheet
 * cannot keep, which are counted rather than kept as something else.
 */
const LIMITS = { rows: 10_000, columns: 100 };

function book(validations: string, extraSheet = ''): XlsxBook {
  const parts: Record<string, string> = {
    'xl/workbook.xml':
      '<workbook><sheets><sheet name="One" r:id="rId1"/><sheet name="Lists" r:id="rId2"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': `<worksheet><sheetData/><dataValidations>${validations}</dataValidations></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet><sheetData>${extraSheet}</sheetData></worksheet>`
  };
  return readXlsx(path => parts[path] ?? null, LIMITS);
}

const rulesOf = (read: XlsxBook) => read.sheets[0].validations.map(validation => validation.rule);

describe('validations from an .xlsx', () => {
  it('reads a list typed into the rule, and its message and style', () => {
    const read = book(
      '<dataValidation type="list" allowBlank="1" showErrorMessage="1" error="Pick one" sqref="A2:A20 C2"><formula1>"North,South,East"</formula1></dataValidation>'
    );
    const [first, second] = read.sheets[0].validations;
    expect(first.rule).toEqual({ kind: 'list', values: ['North', 'South', 'East'] });
    expect(first.strict).toBe(true);
    expect(first.message).toBe('Pick one');
    // One rule over two ranges is two rules here, one each.
    expect(second.range.start).toMatchObject({ row: 1, column: 2 });
  });

  it('reads a list that names a range on another sheet, as its values', () => {
    const read = book(
      '<dataValidation type="list" showErrorMessage="1" sqref="B2:B9"><formula1>Lists!$A$1:$A$3</formula1></dataValidation>',
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Red</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Green</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Red</t></is></c></row>'
    );
    expect(rulesOf(read)).toEqual([{ kind: 'list', values: ['Red', 'Green'] }]);
  });

  it('reads Excel’s operators as bounds, a strict one on a whole number moved by one', () => {
    const read = book(
      '<dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation>' +
        '<dataValidation type="decimal" operator="lessThanOrEqual" sqref="A2"><formula1>2.5</formula1></dataValidation>' +
        '<dataValidation type="whole" sqref="A3"><formula1>1</formula1><formula2>10</formula2></dataValidation>' +
        '<dataValidation type="textLength" operator="lessThan" sqref="A4"><formula1>6</formula1></dataValidation>'
    );
    expect(rulesOf(read)).toEqual([
      { kind: 'number', min: 1, integer: true },
      { kind: 'number', max: 2.5 },
      { kind: 'number', min: 1, max: 10, integer: true },
      { kind: 'text', maxLength: 5 }
    ]);
  });

  it('lets a warning or a note in, as a rule that is not strict does', () => {
    const read = book(
      '<dataValidation type="whole" errorStyle="warning" showErrorMessage="1" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="greaterThan" sqref="A2"><formula1>0</formula1></dataValidation>'
    );
    expect(read.sheets[0].validations.map(validation => validation.strict)).toEqual([false, false]);
  });

  it('moves a date bound in a 1904 file to this sheet’s days', () => {
    const parts: Record<string, string> = {
      'xl/workbook.xml': '<workbook><workbookPr date1904="1"/><sheets><sheet name="One" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml':
        '<worksheet><sheetData/><dataValidations><dataValidation type="date" operator="greaterThanOrEqual" sqref="A1"><formula1>44000</formula1></dataValidation></dataValidations></worksheet>'
    };
    const read = readXlsx(path => parts[path] ?? null, LIMITS);
    expect(read.sheets[0].validations[0].rule).toEqual({ kind: 'date', from: 45462, to: undefined });
  });

  it('reads a list that names a defined name, through the name', () => {
    const parts: Record<string, string> = {
      'xl/workbook.xml':
        '<workbook><sheets><sheet name="One" r:id="rId1"/><sheet name="Lists" r:id="rId2"/></sheets><definedNames><definedName name="states">Lists!$A$1:$A$2</definedName></definedNames></workbook>',
      'xl/_rels/workbook.xml.rels':
        '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml':
        '<worksheet><sheetData/><dataValidations><dataValidation type="list" sqref="A1"><formula1>states</formula1></dataValidation></dataValidations></worksheet>',
      'xl/worksheets/sheet2.xml':
        '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Ohio</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Utah</t></is></c></row></sheetData></worksheet>'
    };
    const read = readXlsx(path => parts[path] ?? null, LIMITS);
    expect(rulesOf(read)).toEqual([{ kind: 'list', values: ['Ohio', 'Utah'] }]);
  });

  it('does not count a rule of any value as one left out', () => {
    const read = book('<dataValidation allowBlank="1" showInputMessage="1" prompt="Type here" sqref="A1"/>');
    expect(read.sheets[0].validations).toEqual([]);
    expect(read.leftOut.validations).toBe(0);
  });

  it('counts what it cannot keep, rather than keeping something else', () => {
    const read = book(
      '<dataValidation type="custom" sqref="A1"><formula1>ISNUMBER(A1)</formula1></dataValidation>' +
        '<dataValidation type="time" operator="greaterThan" sqref="A2"><formula1>0.5</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="notBetween" sqref="A3"><formula1>1</formula1><formula2>2</formula2></dataValidation>' +
        '<dataValidation type="list" sqref="A4"><formula1>Lists!$A$1:$A$3</formula1></dataValidation>'
    );
    expect(read.sheets[0].validations).toEqual([]);
    // The last is a list naming cells that hold nothing.
    expect(read.leftOut.validations).toBe(4);
  });
});
