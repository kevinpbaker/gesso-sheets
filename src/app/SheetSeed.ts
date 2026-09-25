import { GENERAL, NO_BORDERS, PLAIN, type CellFormat } from '../sheet/Format';
import type { SheetDocument } from './SheetDocument';

/**
 * Something to open onto.
 *
 * Phase 1 gave the engine formulas and Phase 2 gives it a wire; an
 * empty grid demonstrates neither. This is a small sheet with real
 * dependencies in it — a column that sums, a column that divides, a
 * row of totals reading the columns — so that a change to one cell
 * visibly moves several others, which is the thing the whole project
 * is evidence for.
 *
 * It is not a fixture: specs build their own sheets. Phase 6 replaces
 * it with whatever the repository loaded.
 */
export function seed(document: SheetDocument): void {
  const regions = ['North', 'South', 'East', 'West', 'Central'];

  document.setCell(0, 0, 'Region');
  document.setCell(0, 1, 'Units');
  document.setCell(0, 2, 'Price');
  document.setCell(0, 3, 'Revenue');
  document.setCell(0, 4, 'Share');

  regions.forEach((region, index) => {
    const row = index + 1;
    const line = row + 1;
    document.setCell(row, 0, region);
    document.setCell(row, 1, String(120 + index * 37));
    document.setCell(row, 2, (9.5 + index * 2.25).toFixed(2));
    document.setCell(row, 3, `=B${line}*C${line}`);
    // Absolute on the total, relative on the row: the pair Phase 5's
    // fill handle has to tell apart when it extends this downwards.
    document.setCell(row, 4, `=ROUND(D${line}/$D$7*100,1)`);
  });

  document.setCell(6, 0, 'Total');
  document.setCell(6, 1, '=SUM(B2:B6)');
  document.setCell(6, 2, '=AVERAGE(C2:C6)');
  document.setCell(6, 3, '=SUM(D2:D6)');

  document.setCell(8, 0, 'Largest');
  document.setCell(8, 1, '=MAX(D2:D6)');
  document.setCell(9, 0, 'Smallest');
  document.setCell(9, 1, '=MIN(D2:D6)');
  document.setCell(10, 0, 'Counted');
  document.setCell(10, 1, '=COUNT(D2:D6)');
  document.setCell(11, 0, 'Healthy');
  document.setCell(11, 1, '=IF(B7>700,"yes","no")');

  // Formatted from Phase 9 onwards, because a sheet that opens with
  // no formatting in it is a demonstration of the format axis that
  // demonstrates nothing — and because the proof needs a formatted
  // sheet to scroll over.
  const header: CellFormat = {
    number: GENERAL,
    paint: {
      ...PLAIN,
      bold: true,
      fill: '#eef2f7',
      align: 'center',
      borders: { ...NO_BORDERS, bottom: { width: 2, color: '' } }
    }
  };
  const money: CellFormat = {
    number: { kind: 'currency', places: 2, symbol: '$' },
    paint: PLAIN
  };
  const share: CellFormat = { number: { kind: 'number', places: 1, thousands: false }, paint: PLAIN };
  const ruledAbove = { ...NO_BORDERS, top: { width: 2, color: '' } };
  const total: CellFormat = { number: money.number, paint: { ...PLAIN, bold: true, borders: ruledAbove } };

  for (let column = 0; column <= 4; column++) {
    document.setFormat(0, column, header);
  }
  for (let row = 1; row <= 6; row++) {
    document.setFormat(row, 2, money);
    document.setFormat(row, 3, row === 6 ? total : money);
    document.setFormat(row, 4, share);
  }
  document.setFormat(6, 0, { number: GENERAL, paint: { ...PLAIN, bold: true, borders: ruledAbove } });
  document.setFormat(6, 1, { number: GENERAL, paint: { ...PLAIN, bold: true, borders: ruledAbove } });
  document.setFormat(6, 4, { number: GENERAL, paint: { ...PLAIN, borders: ruledAbove } });

  document.setSelection(1, 1, 1, 1);
}
