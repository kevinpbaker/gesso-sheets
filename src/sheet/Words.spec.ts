import { describe, expect, it } from 'vitest';

import { Workbook } from './Workbook';
import { Words } from './Words';

describe('the words a column holds', () => {
  function words(...texts: string[]): Words {
    const kept = new Words();
    for (const text of texts) {
      kept.add(text);
    }
    return kept;
  }

  it('finishes the one word a prefix starts', () => {
    expect(words('North', 'South', 'East').complete('No')).toBe('North');
  });

  it('ignores case, and answers in the word’s own spelling', () => {
    expect(words('North').complete('nO')).toBe('North');
  });

  it('offers nothing when the prefix could be more than one word', () => {
    expect(words('North', 'Northwest').complete('Nor')).toBeNull();
    expect(words('North', 'Northwest').complete('Northw')).toBe('Northwest');
  });

  it('offers nothing for a prefix that is already the whole word', () => {
    expect(words('North').complete('north')).toBeNull();
  });

  it('offers nothing for nothing', () => {
    expect(words('North').complete('')).toBeNull();
    expect(words().complete('N')).toBeNull();
  });

  it('keeps a word until the last cell holding it lets it go', () => {
    const kept = words('North', 'North');
    kept.remove('North');
    expect(kept.complete('No')).toBe('North');
    kept.remove('North');
    expect(kept.complete('No')).toBeNull();
    expect(kept.size).toBe(0);
  });
});

describe('a workbook’s words', () => {
  function book(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.sheet(0);
    sheet.setCell(0, 0, 'Region');
    sheet.setCell(1, 0, 'North');
    sheet.setCell(2, 0, 'South');
    sheet.setCell(3, 0, '42');
    sheet.setCell(4, 0, '=A2&"ern"');
    sheet.setCell(0, 1, 'Nowhere');
    sheet.recalculate();
    return workbook;
  }

  it('are the column’s own, not the next column’s', () => {
    expect(book().completeIn(0, 0, 'No')).toBe('North');
    expect(book().completeIn(0, 1, 'No')).toBe('Nowhere');
  });

  it('are text somebody typed, not numbers or what a formula said', () => {
    const workbook = book();
    workbook.sheet(0).recalculate();
    expect(workbook.completeIn(0, 0, '4')).toBeNull();
    // "Northern" is A5's answer and not something typed there.
    expect(workbook.completeIn(0, 0, 'Northe')).toBeNull();
  });

  it('follow the column as it is edited after they were first asked for', () => {
    const workbook = book();
    expect(workbook.completeIn(0, 0, 'We')).toBeNull();
    workbook.setCell(0, 5, 0, 'West');
    expect(workbook.completeIn(0, 0, 'We')).toBe('West');
    workbook.setCell(0, 5, 0, 'Westward');
    expect(workbook.completeIn(0, 0, 'Wes')).toBe('Westward');
    workbook.clearCell(0, 5, 0);
    expect(workbook.completeIn(0, 0, 'We')).toBeNull();
    // A word overwritten by a formula is no longer a word there.
    workbook.setCell(0, 2, 0, '=1');
    expect(workbook.completeIn(0, 0, 'So')).toBeNull();
  });

  it('move with the cells when a column is inserted before them', () => {
    const workbook = book();
    expect(workbook.completeIn(0, 0, 'No')).toBe('North');
    workbook.shift(0, { axis: 'column', at: 0, by: 1 });
    expect(workbook.completeIn(0, 1, 'So')).toBe('South');
    expect(workbook.completeIn(0, 0, 'So')).toBeNull();
  });
});
