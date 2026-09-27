import type { ScriptFormat, ScriptValue } from './protocol';

/**
 * The API a script is handed, as types.
 *
 * Everything a script can reach is in this file, and `SCRIPTS.md`
 * lists the same things. A new capability is a new line in both and a
 * spec, and one that reaches past the workbook is a change to the
 * model first. `scriptWorker.ts` implements these interfaces.
 *
 * It is synchronous on purpose. The run starts from a copy of the
 * workbook's values, so a read needs no round trip and a script needs
 * no `await`. Writes are collected and applied when the run ends, as
 * one step of undo, so a value read back after a write in the same run
 * is what was written: a formula reads back as its text, and its value
 * arrives when the sheet recalculates.
 */

/** The globals a script has: these three, and nothing else. */
export interface ScriptGlobals {
  /** The sheet that was showing when the run started. */
  readonly sheet: ScriptSheet;
  readonly workbook: ScriptWorkbook;
  /** Lines for the run's log, shown under the editor. */
  readonly console: { log(...parts: unknown[]): void };
}

export interface ScriptWorkbook {
  /** Every sheet, in tab order, and the ones this run added after them. */
  readonly sheets: readonly ScriptSheet[];
  /** A sheet by name, ignoring case. Throws when there is none. */
  sheet(name: string): ScriptSheet;
  /**
   * Adds a sheet at the end.
   *
   * Adding a sheet by hand cannot be undone, and neither can this.
   * Undo takes back what the run wrote, and the sheet stays.
   */
  addSheet(name: string): ScriptSheet;
}

export interface ScriptSheet {
  readonly name: string;
  /** `A1`, or `A1:C9`. */
  range(address: string): ScriptRange;
  /** One cell's value: `sheet.read('B4')`. */
  read(address: string): ScriptValue;
  /** One cell: `sheet.write('B4', 12)`. */
  write(address: string, value: ScriptValue): void;
}

export interface ScriptRange {
  /** As written back: `A1:C9`. */
  readonly address: string;
  readonly rows: number;
  readonly columns: number;
  /** Row by row. An empty cell is null, and an error is its code, like `#DIV/0!`. */
  readonly values: ScriptValue[][];
  /**
   * Row by row from the top left, or one value for every cell.
   *
   * A string is written as if typed, so `'=A1*2'` is a formula and
   * `'12'` is a number. Rows may be shorter than the range, and a
   * shorter row leaves the cells past its end alone. More rows or
   * columns than the range has is an error, not a silent crop.
   */
  write(rows: readonly (readonly ScriptValue[])[] | ScriptValue): void;
  /** Formats every cell; a field left out is left alone. */
  format(change: ScriptFormat): void;
  /** Empties every cell, and leaves its format alone. */
  clear(): void;
}
