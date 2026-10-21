import { describe, expect, it } from 'vitest';

import { SheetDocument } from './SheetDocument';
import { TEMPLATE_BUILDS } from './SheetTemplates';
import { TEMPLATES } from './Templates';

/**
 * The workbooks File ▸ New starts from, each through the engine.
 *
 * The seed's spec asserts the seed; this one asserts that every card
 * in the dialog has a workbook behind it and that none of them opens
 * on a `#` anywhere, which is the mistake a template is likeliest to
 * make and the one a person would see first.
 */

function built(id: string): SheetDocument {
  const document = new SheetDocument();
  TEMPLATE_BUILDS[id].build(document);
  document.sheet.recalculate();
  return document;
}

describe('the templates', () => {
  it('has a workbook behind every card, named as the card is', () => {
    for (const template of TEMPLATES) {
      expect(TEMPLATE_BUILDS[template.id]?.name, template.id).toBe(template.name);
    }
    expect(Object.keys(TEMPLATE_BUILDS).sort()).toEqual(TEMPLATES.map(template => template.id).sort());
  });

  it('offers both kinds', () => {
    expect(TEMPLATES.some(template => template.kind === 'template')).toBe(true);
    expect(TEMPLATES.some(template => template.kind === 'example')).toBe(true);
  });

  for (const template of TEMPLATES) {
    it(`builds ${template.name} with nothing pending and no errors on any sheet`, () => {
      const document = built(template.id);
      expect(document.sheet.pending).toBe(0);
      for (let sheet = 0; sheet < document.sheets().length; sheet++) {
        document.activate(sheet);
        for (const { row, column } of document.sheet.entries()) {
          expect(document.display(row, column), `${template.id}, sheet ${sheet}, ${row},${column}`).not.toMatch(/^#/);
        }
      }
    });
  }

  it('totals a budget, planned and actual', () => {
    const document = built('budget');
    expect(document.sheet.value(7, 1)).toBe(4500);
    expect(document.sheet.value(18, 1)).toBe(3230);
    expect(document.sheet.value(20, 1)).toBe(1270);
  });

  it('adds tax to an invoice, and leaves its empty lines empty', () => {
    const document = built('invoice');
    // 12 × 85 + 240, and a tenth on top.
    expect(document.sheet.value(23, 3)).toBe(1260);
    expect(document.sheet.value(25, 3)).toBe(126);
    expect(document.sheet.value(26, 3)).toBe(1386);
    expect(document.display(14, 3)).toBe('');
  });

  it('pays a loan down to nothing at the end of its term', () => {
    const document = built('loan');
    const last = document.sheet.value(70, 1) as number;
    expect(last).toBeCloseTo(0, 6);
    expect(document.sheet.value(11, 1) as number).toBeLessThan(30000);
  });

  it('pays a timesheet with overtime at time and a half', () => {
    const document = built('timesheet');
    expect(document.sheet.value(14, 3)).toBe(37.5);
    expect(document.sheet.value(16, 3)).toBe(37.5 * 45);
  });
});
