import type { RangeRef } from './A1';
import type { Names } from './Names';
import type { Shift } from './Shift';
import type { CellValue } from './Values';
import type { RecalcResult, Workbook } from './Workbook';

export type { RecalcResult };

/**
 * One sheet of a workbook: a name, and the cells under it.
 *
 * A facade, and deliberately a thin one. The store, the dependency
 * graph, the dirty set and the recalculation plan all live on the
 * `Workbook`, because a formula on this sheet may read another one
 * and there is no evaluation order two separate graphs could agree
 * on. What is left here is the *address*: which sheet a row and
 * column mean, which is exactly what a caller holding a sheet wants
 * to stop saying.
 *
 * That split is also what kept twelve phases of code working. Every
 * call reading `sheet.value(row, column)` goes on meaning the same
 * thing; it is the key underneath that grew a third axis.
 *
 * Nothing here imports the framework, touches a worker or knows what
 * a viewport is — the constraint Phase 1 set and every phase since
 * has kept. These specs run in node in milliseconds.
 */
export class Sheet {
  constructor(
    private readonly book: Workbook,
    readonly index: number
  ) {}

  /** What this sheet is called, which is what a reference to it says. */
  get name(): string {
    return this.book.nameOf(this.index);
  }

  get workbook(): Workbook {
    return this.book;
  }

  get colour(): string | null {
    return this.book.colourOf(this.index);
  }

  set colour(value: string | null) {
    this.book.setColour(this.index, value);
  }

  /** The names the *workbook* knows; see `Workbook.names`. */
  get names(): Names {
    return this.book.names;
  }

  get stats(): { evaluated: number; planned: number; plans: number } {
    return this.book.stats;
  }

  get clock(): () => number {
    return this.book.clock;
  }

  set clock(value: () => number) {
    this.book.clock = value;
  }

  get dice(): () => number {
    return this.book.dice;
  }

  set dice(value: () => number) {
    this.book.dice = value;
  }

  // ---------------------------------------------------------------------
  // Editing
  // ---------------------------------------------------------------------

  setCell(row: number, column: number, input: string, asText = false): void {
    this.book.setCell(this.index, row, column, input, asText);
  }

  clearCell(row: number, column: number): void {
    this.book.clearCell(this.index, row, column);
  }

  /** Inserts or deletes rows or columns of *this* sheet; see `Workbook.shift`. */
  shift(shift: Shift): number {
    return this.book.shift(this.index, shift);
  }

  // ---------------------------------------------------------------------
  // Recalculation
  // ---------------------------------------------------------------------

  /**
   * Cells still waiting, anywhere in the workbook.
   *
   * Not "on this sheet", and that is not a shortcut: a recalculation
   * is the workbook's because the order is the workbook's, so how
   * much of it is left is a workbook-wide number. A per-sheet count
   * would be a number nobody could act on.
   */
  get pending(): number {
    return this.book.pending;
  }

  recalculate(budget?: number): RecalcResult {
    return this.book.recalculate(budget);
  }

  namesChanged(): void {
    this.book.namesChanged();
  }

  // ---------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------

  value(row: number, column: number): CellValue {
    return this.book.value(this.index, row, column);
  }

  /** What was typed, which is what a cell editor opens with. */
  input(row: number, column: number): string {
    return this.book.input(this.index, row, column);
  }

  /**
   * The value as general text.
   *
   * **Not what the screen shows.** A cell's display string depends on
   * its number format as well as its value, and a format is not the
   * sheet's; ask `SheetDocument.display` for what a person sees.
   */
  display(row: number, column: number): string {
    return this.book.display(this.index, row, column);
  }

  /** The `EvaluationContext`'s name half. */
  rangeForName(name: string): RangeRef | null {
    return this.book.names.rangeOf(name);
  }

  /** The value behind a key, which already carries its sheet. */
  valueAt(key: number): CellValue {
    return this.book.valueAt(key);
  }

  /** How far a whole-column reference reads on this sheet. */
  get usedRows(): number {
    return this.book.usedRowsOf(this.index);
  }

  /** Cells that hold something on this sheet. */
  get size(): number {
    return this.book.sizeOf(this.index);
  }

  /** Every non-empty cell, for a repository to write out. */
  entries(): Generator<{ row: number; column: number; input: string }> {
    return this.book.entriesOf(this.index);
  }

  /** What a cell reads, for `engine.explain` in Phase 7. */
  precedentsOf(row: number, column: number): number[] {
    return this.book.precedentsOfCell(this.index, row, column);
  }

  /** What reads a cell. */
  dependentsOf(row: number, column: number): number[] {
    return this.book.dependentsOfCell(this.index, row, column);
  }
}
