import { cellKey, columnOf, rowOf } from './A1';
import { DEFAULT_FORMAT, keyOf, type CellFormat } from './Format';
import { shiftIndex, type Shift } from './Shift';

/**
 * Which format each cell has, and the table of formats themselves.
 *
 * **The palette is the point.** A column formatted as currency is
 * fifty thousand cells sharing one format, and the wire carries one
 * palette entry and a small integer per *visible* cell. The
 * alternative — a format record per cell — puts a nested object on
 * the wire for every cell in the window and gives the differ one to
 * walk on every publish, to discover that fifty thousand of them are
 * identical. Phase 0 measured the same argument about values and
 * settled it the same way.
 *
 * Entry 0 is always the default format, so a cell with no entry and a
 * cell pointing at 0 mean the same thing and nothing has to special-
 * case the difference.
 *
 * **Entries are never removed.** Renumbering would move every cell's
 * index and put the whole window back on the wire to reclaim a few
 * bytes. A sheet reformatted a thousand times accumulates a thousand
 * entries, which is a few kilobytes and is bounded by how many
 * *distinct* formats have ever existed — see `compact`, which a save
 * could use and deliberately does not.
 *
 * **A region is stored as a region.** A whole column formatted as
 * currency is one number here, not ten thousand, and the whole sheet
 * made bold is one number rather than a million. This is the same
 * insight as the palette applied to the other axis, and it was not
 * in the first version of this file: pressing ctrl-A and then ctrl-B
 * wrote a million cell entries, which reached somebody's disk as a
 * thirty-megabyte file and would have been parsed back on every load
 * for the rest of the sheet's life. A browser found that; no spec
 * did.
 *
 * Resolution runs cell, then row, then column, then sheet — the
 * order every spreadsheet uses, and the one that makes "this cell is
 * special" beat "this column is currency".
 */
export class Formats {
  /** Palette entries by id. Index 0 is the default. */
  private readonly palette: CellFormat[] = [DEFAULT_FORMAT];
  /** Palette id by the format's own key, so equal formats intern. */
  private readonly ids = new Map<string, number>([[keyOf(DEFAULT_FORMAT), 0]]);
  /** Cell key to palette id. A cell with the default is absent. */
  private readonly cells = new Map<number, number>();
  /** Row index to palette id, for a whole row formatted at once. */
  private readonly rows = new Map<number, number>();
  private readonly columns = new Map<number, number>();
  /** The whole sheet's format, or 0. */
  private sheet = 0;

  /** The palette as the contract carries it. */
  get entries(): readonly CellFormat[] {
    return this.palette;
  }

  get size(): number {
    return this.cells.size;
  }

  /** The palette id for a format, adding it if it is new. */
  idFor(format: CellFormat): number {
    const key = keyOf(format);
    const existing = this.ids.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const id = this.palette.length;
    this.palette.push(format);
    this.ids.set(key, id);
    return id;
  }

  /**
   * The palette id a cell ends up with.
   *
   * Cell, then row, then column, then sheet: the order every
   * spreadsheet resolves in, and the one that lets "this cell is
   * special" beat "this column is currency" rather than the other way
   * round.
   */
  idAt(row: number, column: number): number {
    return (
      this.cells.get(cellKey(row, column)) ??
      this.rows.get(row) ??
      this.columns.get(column) ??
      this.sheet
    );
  }

  /** What the regions alone would give a cell, ignoring any override. */
  regionIdAt(row: number, column: number): number {
    return this.rows.get(row) ?? this.columns.get(column) ?? this.sheet;
  }

  /** The whole sheet, one row, or one column. */
  get sheetId(): number {
    return this.sheet;
  }

  rowId(row: number): number {
    return this.rows.get(row) ?? 0;
  }

  columnId(column: number): number {
    return this.columns.get(column) ?? 0;
  }

  /**
   * The cell overrides inside a region, as keys and palette ids.
   *
   * Bounded by how many cells anybody has actually formatted, which
   * is what makes a region change affordable: the whole sheet made
   * bold walks the overrides — a handful — and not the million cells
   * the region covers.
   */
  overridesIn(scope: 'sheet' | 'row' | 'column', index: number): [number, number][] {
    const found: [number, number][] = [];
    for (const [key, held] of this.cells) {
      if (covers(scope, index, key)) {
        found.push([key, held]);
      }
    }
    return found;
  }

  /**
   * Points a region at a palette id, leaving the overrides alone.
   *
   * It does **not** clear them, and the first version of this did.
   * Clearing looked right — formatting a column that holds
   * individually-formatted cells has to change all of it — and was
   * wrong for the reason a format change is a change and not a
   * format: making the whole sheet bold has to keep the currency in
   * column C, and clearing threw it away. Applying the change to the
   * region *and* to each override inside it is what does both, and
   * that is the caller's job because only the caller knows the
   * change.
   */
  setRegion(scope: 'sheet' | 'row' | 'column', index: number, id: number): void {
    this.writeRegion(scope, index, id);
  }

  private writeRegion(scope: 'sheet' | 'row' | 'column', index: number, id: number): void {
    if (scope === 'sheet') {
      this.sheet = id;
      return;
    }
    const map = scope === 'row' ? this.rows : this.columns;
    if (id === 0) {
      map.delete(index);
    } else {
      map.set(index, id);
    }
  }

  /** Every row and column format, for a repository to write out. */
  get regions(): { sheet: number; rows: [number, number][]; columns: [number, number][] } {
    return { sheet: this.sheet, rows: [...this.rows], columns: [...this.columns] };
  }

  formatAt(row: number, column: number): CellFormat {
    return this.palette[this.idAt(row, column)] ?? DEFAULT_FORMAT;
  }

  byId(id: number): CellFormat {
    return this.palette[id] ?? DEFAULT_FORMAT;
  }

  /**
   * Points a cell at a palette id.
   *
   * Id 0 removes the entry rather than storing a zero, so an
   * unformatted sheet holds nothing at all and `size` answers "how
   * many cells has anybody formatted" rather than "how many cells has
   * anybody looked at".
   */
  setId(row: number, column: number, id: number): void {
    const key = cellKey(row, column);
    // An override equal to what the regions already say is not an
    // override. Without this, formatting a bold sheet bold again
    // writes a million entries that all mean "the same as the sheet".
    if (id === this.regionIdAt(row, column)) {
      this.cells.delete(key);
      return;
    }
    this.cells.set(key, id);
  }

  /** Points a cell at a palette id, by its packed key. */
  setIdByKey(key: number, id: number): void {
    this.setId(rowOf(key), columnOf(key), id);
  }

  setFormat(row: number, column: number, format: CellFormat): void {
    this.setId(row, column, this.idFor(format));
  }

  /** Every formatted cell, for a repository to write out. */
  *entriesOf(): Generator<{ row: number; column: number; id: number }> {
    for (const [key, id] of this.cells) {
      yield { row: rowOf(key), column: columnOf(key), id };
    }
  }

  /**
   * The palette with nothing in it that no cell points at, and the
   * cells renumbered onto it.
   *
   * Used when writing a file and *not* while a sheet is open, which
   * is the whole reason it is a separate method. Renumbering a live
   * palette changes the index of every cell on screen, so the window
   * and the palette would both go out in full — a hundred patches to
   * reclaim a hundred bytes. On the way to disk none of that is true:
   * nothing is watching, and the file is smaller forever.
   */
  compact(): {
    palette: readonly CellFormat[];
    cells: readonly { row: number; column: number; id: number }[];
    regions: { sheet: number; rows: [number, number][]; columns: [number, number][] };
  } {
    const used = new Map<number, number>([[0, 0]]);
    const palette: CellFormat[] = [DEFAULT_FORMAT];
    const renumber = (id: number): number => {
      let next = used.get(id);
      if (next === undefined) {
        next = palette.length;
        used.set(id, next);
        palette.push(this.byId(id));
      }
      return next;
    };
    const regions = {
      sheet: renumber(this.sheet),
      rows: [...this.rows].map(([row, id]): [number, number] => [row, renumber(id)]),
      columns: [...this.columns].map(([column, id]): [number, number] => [column, renumber(id)])
    };
    const cells: { row: number; column: number; id: number }[] = [];
    for (const cell of this.entriesOf()) {
      cells.push({ row: cell.row, column: cell.column, id: renumber(cell.id) });
    }
    return { palette, cells, regions };
  }

  /**
   * Moves every format because the sheet changed shape.
   *
   * Three maps to carry rather than one, and the region maps are the
   * reason this is cheap: a column formatted as currency is one entry
   * that moves one place, not ten thousand cells that each move.
   *
   * A region whose index was deleted is dropped — there is no column
   * left to be currency — and so is a cell override in a deleted row.
   */
  shift(shift: Shift): void {
    const cells = new Map<number, number>();
    for (const [key, id] of this.cells) {
      const row = rowOf(key);
      const column = columnOf(key);
      const index = shift.axis === 'row' ? row : column;
      const moved = shiftIndex(index, shift);
      if (moved === -1) {
        continue;
      }
      cells.set(moved === index ? key : shift.axis === 'row' ? cellKey(moved, column) : cellKey(row, moved), id);
    }
    this.cells.clear();
    for (const [key, id] of cells) {
      this.cells.set(key, id);
    }

    const along = shift.axis === 'row' ? this.rows : this.columns;
    const moved = new Map<number, number>();
    for (const [index, id] of along) {
      const next = shiftIndex(index, shift);
      if (next !== -1) {
        moved.set(next, id);
      }
    }
    along.clear();
    for (const [index, id] of moved) {
      along.set(index, id);
    }
  }

  /** Replaces everything, for a load. */
  restore(
    palette: readonly CellFormat[],
    cells: readonly { row: number; column: number; id: number }[],
    regions: {
      sheet: number;
      rows: readonly (readonly [number, number])[];
      columns: readonly (readonly [number, number])[];
    } = {
      sheet: 0,
      rows: [],
      columns: []
    }
  ): void {
    this.palette.length = 0;
    this.ids.clear();
    this.cells.clear();
    this.rows.clear();
    this.columns.clear();
    this.sheet = 0;
    // The default is always entry 0, whatever the file says, so a
    // file written by a build that numbered them differently still
    // loads with the same meaning for "unformatted".
    this.palette.push(DEFAULT_FORMAT);
    this.ids.set(keyOf(DEFAULT_FORMAT), 0);
    for (let id = 1; id < palette.length; id++) {
      const format = palette[id];
      const key = keyOf(format);
      if (this.ids.has(key)) {
        continue;
      }
      this.ids.set(key, this.palette.length);
      this.palette.push(format);
    }
    // Regions before overrides, because `setId` compares an override
    // against what the regions give and drops it when they agree.
    const idOf = (id: number): number => this.idFor(palette[id] ?? DEFAULT_FORMAT);
    this.sheet = regions.sheet === 0 ? 0 : idOf(regions.sheet);
    for (const [row, id] of regions.rows) {
      this.writeRegion('row', row, idOf(id));
    }
    for (const [column, id] of regions.columns) {
      this.writeRegion('column', column, idOf(id));
    }
    for (const cell of cells) {
      // Through `idFor` rather than by the file's number, because the
      // loop above may have dropped a duplicate and shifted the rest.
      this.setFormat(cell.row, cell.column, palette[cell.id] ?? DEFAULT_FORMAT);
    }
  }
}

/** Whether a region covers a cell key. */
function covers(scope: 'sheet' | 'row' | 'column', index: number, key: number): boolean {
  if (scope === 'sheet') {
    return true;
  }
  return scope === 'row' ? rowOf(key) === index : columnOf(key) === index;
}
