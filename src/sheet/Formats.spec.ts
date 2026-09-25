import { describe, expect, it } from 'vitest';

import { DEFAULT_FORMAT, PLAIN, type CellFormat } from './Format';
import { Formats } from './Formats';

const bold: CellFormat = { ...DEFAULT_FORMAT, paint: { ...PLAIN, bold: true } };
const money: CellFormat = { number: { kind: 'currency', places: 2, symbol: '$' }, paint: PLAIN };

describe('the palette', () => {
  it('starts with the default format as entry zero', () => {
    const formats = new Formats();
    expect(formats.entries).toHaveLength(1);
    expect(formats.idFor(DEFAULT_FORMAT)).toBe(0);
  });

  /**
   * The property the whole design rests on: a column of fifty
   * thousand cells formatted the same way is one palette entry, and
   * what goes on the wire per cell is a small integer.
   */
  it('gives equal formats the same id, however many cells use them', () => {
    const formats = new Formats();
    for (let row = 0; row < 1_000; row++) {
      formats.setFormat(row, 0, { ...money });
    }
    expect(formats.entries).toHaveLength(2);
    expect(formats.idAt(0, 0)).toBe(1);
    expect(formats.idAt(999, 0)).toBe(1);
  });

  it('gives different formats different ids', () => {
    const formats = new Formats();
    expect(formats.idFor(bold)).toBe(1);
    expect(formats.idFor(money)).toBe(2);
    expect(formats.idFor(bold)).toBe(1);
  });

  it('answers the default for a cell nobody has formatted', () => {
    const formats = new Formats();
    expect(formats.idAt(9, 9)).toBe(0);
    expect(formats.formatAt(9, 9)).toEqual(DEFAULT_FORMAT);
  });

  /**
   * `size` answers "how many cells has anybody formatted", not "how
   * many cells has anybody looked at".
   */
  it('holds nothing for a cell put back to the default', () => {
    const formats = new Formats();
    formats.setFormat(3, 3, bold);
    expect(formats.size).toBe(1);
    formats.setFormat(3, 3, DEFAULT_FORMAT);
    expect(formats.size).toBe(0);
    expect(formats.idAt(3, 3)).toBe(0);
  });
});

describe('writing the palette out', () => {
  it('drops entries nothing points at, and renumbers what is left', () => {
    const formats = new Formats();
    formats.setFormat(0, 0, bold);
    formats.setFormat(0, 0, money);
    // `bold` is now an orphan: it is in the palette and no cell uses it.
    expect(formats.entries).toHaveLength(3);

    const written = formats.compact();
    expect(written.palette).toHaveLength(2);
    expect(written.palette[1]).toEqual(money);
    expect(written.cells).toEqual([{ row: 0, column: 0, id: 1 }]);
  });

  /**
   * Compacting is for the way to disk and not for a live sheet:
   * renumbering changes the index of every cell on screen, so the
   * window and the palette would both go out in full to reclaim a few
   * bytes. This is the spec that says it did not happen behind
   * anyone's back.
   */
  it('leaves the live palette alone', () => {
    const formats = new Formats();
    formats.setFormat(0, 0, bold);
    formats.setFormat(0, 0, money);
    formats.compact();
    expect(formats.entries).toHaveLength(3);
    expect(formats.idAt(0, 0)).toBe(2);
  });
});

describe('reading the palette back', () => {
  it('round-trips through compact and restore', () => {
    const formats = new Formats();
    formats.setFormat(1, 1, bold);
    formats.setFormat(4, 2, money);
    const written = formats.compact();

    const loaded = new Formats();
    loaded.restore(written.palette, written.cells);
    expect(loaded.formatAt(1, 1)).toEqual(bold);
    expect(loaded.formatAt(4, 2)).toEqual(money);
    expect(loaded.formatAt(0, 0)).toEqual(DEFAULT_FORMAT);
  });

  /**
   * A file written by a build that numbered the palette differently
   * still has to mean the same thing by "unformatted".
   */
  it('keeps the default as entry zero whatever the file says', () => {
    const loaded = new Formats();
    loaded.restore([bold, money], [{ row: 0, column: 0, id: 0 }]);
    expect(loaded.formatAt(0, 0)).toEqual(bold);
    expect(loaded.byId(0)).toEqual(DEFAULT_FORMAT);
  });

  it('survives a file that lists the same format twice', () => {
    const loaded = new Formats();
    loaded.restore(
      [DEFAULT_FORMAT, money, { ...money }],
      [
        { row: 0, column: 0, id: 1 },
        { row: 1, column: 0, id: 2 }
      ]
    );
    expect(loaded.formatAt(0, 0)).toEqual(money);
    expect(loaded.formatAt(1, 0)).toEqual(money);
    expect(loaded.idAt(0, 0)).toBe(loaded.idAt(1, 0));
    expect(loaded.entries).toHaveLength(2);
  });

  it('forgets what was there before', () => {
    const formats = new Formats();
    formats.setFormat(7, 7, bold);
    formats.restore([DEFAULT_FORMAT], []);
    expect(formats.size).toBe(0);
    expect(formats.formatAt(7, 7)).toEqual(DEFAULT_FORMAT);
  });
});
