import { describe, expect, it } from 'vitest';

import { Notes } from './Notes';
import { readXlsx } from './Xlsx';

describe('the notes on a sheet', () => {
  it('keeps a note on its cell, and forgets it when it is emptied', () => {
    const notes = new Notes();
    notes.set(2, 1, 'Check with finance');
    expect(notes.at(2, 1)).toBe('Check with finance');
    expect(notes.at(1, 2)).toBe('');
    notes.set(2, 1, '');
    expect(notes.size).toBe(0);
  });

  it('moves a note down with its cell when a row is inserted above it', () => {
    const notes = new Notes();
    notes.set(4, 0, 'here');
    notes.shift({ axis: 'row', at: 2, by: 3 });
    expect(notes.at(7, 0)).toBe('here');
    expect(notes.at(4, 0)).toBe('');
  });

  it('lets a note go with the column it was on', () => {
    const notes = new Notes();
    notes.set(0, 3, 'gone');
    notes.set(0, 5, 'kept');
    notes.shift({ axis: 'column', at: 3, by: -1 });
    expect(notes.all()).toEqual([{ row: 0, column: 4, text: 'kept' }]);
  });

  it('lists the notes in reading order, for a file', () => {
    const notes = new Notes();
    notes.set(3, 0, 'c');
    notes.set(0, 2, 'b');
    notes.set(0, 1, 'a');
    expect(notes.all().map(note => note.text)).toEqual(['a', 'b', 'c']);
  });
});

describe('an .xlsx’s comments, as notes', () => {
  const WORKBOOK = {
    'xl/workbook.xml': '<workbook><sheets><sheet name="One" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels':
      '<Relationships><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>'
  };
  const LIMITS = { rows: 10_000, columns: 100 };

  function notesOf(comments: string, limits = LIMITS) {
    const parts: Record<string, string> = { ...WORKBOOK, 'xl/comments1.xml': comments };
    return readXlsx(path => parts[path] ?? null, limits).sheets[0].notes;
  }

  it('finds the comments from the worksheet’s own relationships', () => {
    const notes = notesOf(
      '<comments><authors><author>Kev</author></authors><commentList>' +
        '<comment ref="B3" authorId="0"><text><t>Plain text</t></text></comment>' +
        '</commentList></comments>'
    );
    expect(notes).toEqual([{ row: 2, column: 1, text: 'Plain text' }]);
  });

  it('takes off the bold name Excel writes in front, whoever it names', () => {
    const notes = notesOf(
      '<comments><authors><author>Somebody Else</author></authors><commentList>' +
        '<comment ref="A1" authorId="0"><text>' +
        '<r><rPr><b/></rPr><t>Balaji:</t></r><r><t xml:space="preserve">\nTestComment1</t></r>' +
        '</text></comment></commentList></comments>'
    );
    expect(notes[0]?.text).toBe('TestComment1');
  });

  it('keeps a colon that is part of what somebody wrote', () => {
    const notes = notesOf(
      '<comments><authors><author>Kev</author></authors><commentList>' +
        '<comment ref="A1" authorId="0"><text><r><t>Note: check this</t></r></text></comment>' +
        '</commentList></comments>'
    );
    expect(notes[0]?.text).toBe('Note: check this');
  });

  it('unwraps a threaded comment to the comment itself', () => {
    const notes = notesOf(
      '<comments><authors><author>tc={1}</author></authors><commentList>' +
        '<comment ref="C2" authorId="0"><text><t>[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel.\n\nComment:\n    testing</t></text></comment>' +
        '</commentList></comments>'
    );
    expect(notes[0]).toEqual({ row: 1, column: 2, text: 'testing' });
  });

  it('drops a comment past the edge of the sheet', () => {
    const notes = notesOf(
      '<comments><commentList><comment ref="A20000"><text><t>far</t></text></comment></commentList></comments>'
    );
    expect(notes).toEqual([]);
  });

  it('is nothing for a sheet with no comments', () => {
    const parts: Record<string, string> = { ...WORKBOOK };
    delete parts['xl/worksheets/_rels/sheet1.xml.rels'];
    expect(readXlsx(path => parts[path] ?? null, LIMITS).sheets[0].notes).toEqual([]);
  });
});
