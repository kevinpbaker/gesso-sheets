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

/** Conditional formats read from an `.xlsx` as Excel writes them — Phase 23. */
describe('conditional formats from an .xlsx', () => {
  const STYLES =
    '<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="2">' +
    '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>' +
    '<dxf><font><i/></font></dxf>' +
    '</dxfs></styleSheet>';

  function conditional(blocks: string) {
    const parts: Record<string, string> = {
      'xl/workbook.xml': '<workbook><sheets><sheet name="One" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/styles.xml': STYLES,
      'xl/worksheets/sheet1.xml': `<worksheet><sheetData/>${blocks}</worksheet>`
    };
    return readXlsx(path => parts[path] ?? null, LIMITS);
  }

  it('reads the shorthands as themselves, in priority order, with their paint', () => {
    const read = conditional(
      '<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1:B9"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="late"><formula>NOT(ISERROR(SEARCH("late",B1)))</formula></cfRule></conditionalFormatting>'
    );
    expect(read.sheets[0].conditional.map(rule => [rule.test, rule.paint])).toEqual([
      [{ kind: 'textContains', text: 'late' }, { fill: '#ffc7ce', color: '#9c0006', bold: true }],
      [{ kind: 'lessThan', value: 0 }, { italic: true }]
    ]);
  });

  it('keeps what has no shorthand as the formula it means', () => {
    const read = conditional(
      '<conditionalFormatting sqref="C2:C9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThanOrEqual"><formula>$Z$1</formula></cfRule>' +
        '<cfRule type="beginsWith" dxfId="0" priority="2" operator="beginsWith" text="N"><formula>LEFT(C2,1)="N"</formula></cfRule></conditionalFormatting>'
    );
    expect(read.sheets[0].conditional.map(rule => rule.test)).toEqual([
      { kind: 'formula', input: '=C2>=$Z$1' },
      { kind: 'formula', input: '=LEFT(C2,1)="N"' }
    ]);
  });

  it('gives a second range the formula moved to its own first cell', () => {
    const read = conditional(
      '<conditionalFormatting sqref="A2:A5 C2:C5"><cfRule type="expression" dxfId="0" priority="1"><formula>A2&gt;B2</formula></cfRule></conditionalFormatting>'
    );
    expect(read.sheets[0].conditional.map(rule => rule.test)).toEqual([
      { kind: 'formula', input: '=A2>B2' },
      { kind: 'formula', input: '=C2>D2' }
    ]);
  });

  it('reads a colour scale of two or three colours', () => {
    const read = conditional(
      '<conditionalFormatting sqref="D1:D9"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
        '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    );
    expect(read.sheets[0].conditional[0].scale).toEqual({ from: '#f8696b', middle: '#ffeb84', to: '#63be7b' });
  });

  it('counts what has no formula and no counterpart here', () => {
    const read = conditional(
      '<conditionalFormatting sqref="E1:E9"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>' +
        '<cfRule type="top10" dxfId="0" priority="2" rank="3"/><cfRule type="duplicateValues" dxfId="0" priority="3"/></conditionalFormatting>'
    );
    expect(read.sheets[0].conditional).toEqual([]);
    expect(read.leftOut['conditional formats']).toBe(3);
  });
});
