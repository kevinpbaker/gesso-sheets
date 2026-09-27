import type { ArrayValue } from './FunctionKit';
import type { CellValue } from './Values';

/**
 * Functions the workbook's script defines, as the evaluator sees them —
 * Part Six.
 *
 * An interface and nothing else, because the interpreter that runs them
 * is not the sheet's: `src/sheet` imports nothing outside itself, and
 * QuickJS is WebAssembly on the application worker. The app layer
 * implements this (`SheetFunctions`) and hands it to the `Workbook`.
 *
 * A formula reaches one only after every built-in and every defined
 * name has been asked, so a script can never shadow the sheet's own
 * library, and a name somebody defined keeps meaning what it meant.
 */
export interface ScriptFunctions {
  /** Whether the script defines a function by this name, as a formula writes it (upper-cased). */
  has(name: string): boolean;
  /**
   * Calls it, synchronously. Arguments are values or arrays, as the
   * evaluator has them; the answer is one or the other, or an error.
   * `at` is the cell calling, so a failure can be explained there.
   */
  call(
    name: string,
    args: readonly (CellValue | ArrayValue)[],
    at: { readonly sheet: number; readonly row: number; readonly column: number } | undefined
  ): CellValue | ArrayValue;
  /** A recalculation slice has begun; the time spent in scripts is counted from here. */
  startSlice(): void;
  /**
   * Whether this slice has spent its share of time in scripts, and should
   * hand the thread back even though it has cells left in its budget.
   */
  overBudget(): boolean;
}
