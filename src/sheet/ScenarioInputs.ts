import { cellKey, columnOf, rowOf } from './A1';
import { shiftFormula, shiftIndex, type Shift } from './Shift';

/** A named way the workbook might have gone: Optimistic, Pessimistic. */
export interface Scenario {
  /** Stable across a rename, which is what the inputs are kept under. */
  readonly id: string;
  readonly name: string;
}

/** One cell a scenario types differently, where it is on its sheet. */
export interface ScenarioInput {
  readonly scenario: string;
  readonly row: number;
  readonly column: number;
  readonly input: string;
}

/**
 * What each scenario types into one sheet's cells.
 *
 * Per sheet and keyed by cell, as the notes are, and for the notes'
 * reason: an input belongs to its cell, so inserting a row above it
 * carries it down, deleting its row takes it away, and deleting the
 * sheet takes all of it. The scenarios themselves are the workbook's
 * (see `Scenario`); this holds their cells, under each one's id.
 *
 * An input is whatever was typed, formula or not. A formula is moved
 * the way the sheet's own formulas are by an insert on its own sheet,
 * so `=B2*2` still means the cell it meant.
 */
export class ScenarioInputs {
  private readonly byScenario = new Map<string, Map<number, string>>();

  get size(): number {
    let size = 0;
    for (const cells of this.byScenario.values()) {
      size += cells.size;
    }
    return size;
  }

  /** What `scenario` types into a cell, or null where it types nothing. */
  at(scenario: string, row: number, column: number): string | null {
    return this.byScenario.get(scenario)?.get(cellKey(row, column)) ?? null;
  }

  /** Types an input into a cell for `scenario`; null takes it away and leaves the cell as the base has it. */
  set(scenario: string, row: number, column: number, input: string | null): void {
    const key = cellKey(row, column);
    if (input === null) {
      const cells = this.byScenario.get(scenario);
      cells?.delete(key);
      if (cells?.size === 0) {
        this.byScenario.delete(scenario);
      }
      return;
    }
    let cells = this.byScenario.get(scenario);
    if (cells === undefined) {
      cells = new Map();
      this.byScenario.set(scenario, cells);
    }
    cells.set(key, input);
  }

  /** One scenario's inputs on this sheet, by row and then column. */
  of(scenario: string): ScenarioInput[] {
    const cells = this.byScenario.get(scenario);
    if (cells === undefined) {
      return [];
    }
    return [...cells]
      .map(([key, input]) => ({ scenario, row: rowOf(key), column: columnOf(key), input }))
      .sort((a, b) => a.row - b.row || a.column - b.column);
  }

  /** Every scenario's inputs on this sheet, which is what a file keeps. */
  all(): ScenarioInput[] {
    return [...this.byScenario.keys()].sort().flatMap(scenario => this.of(scenario));
  }

  /** Forgets a scenario that has been deleted. */
  drop(scenario: string): void {
    this.byScenario.delete(scenario);
  }

  /** Copies one scenario's inputs to another, for a scenario made from one. */
  duplicate(from: string, to: string): void {
    const cells = this.byScenario.get(from);
    if (cells !== undefined) {
      this.byScenario.set(to, new Map(cells));
    }
  }

  /**
   * Follows an insert or delete anywhere in the workbook, for inputs on
   * the sheet called `onSheet`.
   *
   * Called for every sheet, as the workbook rewrites every sheet's
   * formulas: an input moves only when its own sheet is the one that
   * changed shape, but a formula among them is rewritten wherever it is,
   * because `=Plan!B2` on another sheet points at a row that moved.
   */
  shift(shift: Shift, onSheet: string): void {
    const moves = shift.sheet === undefined || shift.sheet.toUpperCase() === onSheet.toUpperCase();
    for (const [scenario, cells] of this.byScenario) {
      const moved = new Map<number, string>();
      for (const [key, input] of cells) {
        const row = rowOf(key);
        const column = columnOf(key);
        const next = moves ? shiftIndex(shift.axis === 'row' ? row : column, shift) : shift.axis === 'row' ? row : column;
        if (next === -1) {
          continue;
        }
        moved.set(shift.axis === 'row' ? cellKey(next, column) : cellKey(row, next), shiftFormula(input, shift, onSheet));
      }
      this.byScenario.set(scenario, moved);
    }
  }

  /** Replaces everything, for a load and for an undo. */
  restore(inputs: readonly ScenarioInput[]): void {
    this.byScenario.clear();
    for (const entry of inputs) {
      this.set(entry.scenario, entry.row, entry.column, entry.input);
    }
  }

  copy(): ScenarioInputs {
    const copied = new ScenarioInputs();
    copied.restore(this.all());
    return copied;
  }
}
