import { cellKey, columnOf, rowOf } from './A1';
import { shiftIndex, type Shift } from './Shift';

/** One note, where it is. */
export interface Note {
  readonly row: number;
  readonly column: number;
  readonly text: string;
}

/**
 * The notes on one sheet: a few words somebody left on a cell.
 *
 * Sparse, keyed like the formats are, and moved by a structural edit
 * the way the formats are — a note belongs to its cell, and inserting a
 * row above the cell carries the note down with it. A note on a row
 * that is deleted goes with the row.
 */
export class Notes {
  private readonly cells = new Map<number, string>();

  get size(): number {
    return this.cells.size;
  }

  at(row: number, column: number): string {
    return this.cells.get(cellKey(row, column)) ?? '';
  }

  /** Writes a note; an empty one takes the note away. */
  set(row: number, column: number, text: string): void {
    const key = cellKey(row, column);
    if (text === '') {
      this.cells.delete(key);
    } else {
      this.cells.set(key, text);
    }
  }

  /** Every note, by row and then column, which is the order a file writes them in. */
  all(): Note[] {
    return [...this.cells]
      .map(([key, text]) => ({ row: rowOf(key), column: columnOf(key), text }))
      .sort((a, b) => a.row - b.row || a.column - b.column);
  }

  shift(shift: Shift): void {
    const moved = new Map<number, string>();
    for (const [key, text] of this.cells) {
      const row = rowOf(key);
      const column = columnOf(key);
      const index = shift.axis === 'row' ? row : column;
      const next = shiftIndex(index, shift);
      if (next === -1) {
        continue;
      }
      moved.set(shift.axis === 'row' ? cellKey(next, column) : cellKey(row, next), text);
    }
    this.cells.clear();
    for (const [key, text] of moved) {
      this.cells.set(key, text);
    }
  }

  /** Replaces every note, for a load and for an undo. */
  restore(notes: readonly Note[]): void {
    this.cells.clear();
    for (const note of notes) {
      if (note.text !== '') {
        this.cells.set(cellKey(note.row, note.column), note.text);
      }
    }
  }

  copy(): Notes {
    const copied = new Notes();
    copied.restore(this.all());
    return copied;
  }
}
