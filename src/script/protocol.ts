/**
 * What crosses between a script's host and the worker a script runs in.
 *
 * Plain data both ways, and the host trusts none of what comes back:
 * see `ScriptHost`, which checks every change a worker reports before
 * it lets one near the workbook. The worker is where untrusted code
 * runs, so the worker is the other side of the trust boundary, and its
 * messages are treated the way a server treats a request.
 */

/** A value a script may write into a cell: what a cell can hold, and nothing with behaviour. */
export type ScriptValue = string | number | boolean | null;

/**
 * A formatting change a script may ask for.
 *
 * A short list of plain fields, each checked by the host, rather than
 * the sheet's own `SheetFormatChange`: every field here is something a
 * script can reach, and the model says each one is written down.
 */
export interface ScriptFormat {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly wrap?: boolean;
  /** `#rrggbb`. */
  readonly color?: string;
  /** `#rrggbb`. */
  readonly fill?: string;
  readonly align?: 'auto' | 'start' | 'center' | 'end';
  /** In points, 6 to 72. */
  readonly fontSize?: number;
  readonly number?: ScriptNumberFormat;
}

export const SCRIPT_NUMBER_FORMATS = ['general', 'number', 'currency', 'percent', 'date', 'text'] as const;
export type ScriptNumberFormat = (typeof SCRIPT_NUMBER_FORMATS)[number];

/** The sheet's extent, and how much one run may change. */
export interface ScriptRunLimits {
  /** Cells a run may write or format, counted together. */
  readonly writes: number;
  readonly rows: number;
  readonly columns: number;
  /** Sheets a run may add. */
  readonly sheets: number;
}

/** One sheet as a script sees it: its name, and its cells' values by `row:column`. */
export interface ScriptBookSheet {
  readonly name: string;
  readonly cells: Readonly<Record<string, ScriptValue>>;
}

/** The workbook as a script sees it, as values: what the run starts from. */
export interface ScriptBook {
  readonly sheets: readonly ScriptBookSheet[];
  /** Which sheet `sheet` means. */
  readonly active: number;
}

/** Host to worker: run this, against this workbook, within these limits. */
export interface ScriptRunMessage {
  readonly type: 'run';
  readonly source: string;
  readonly book: ScriptBook;
  readonly limits: ScriptRunLimits;
}

/**
 * One change a script asked for, as the worker reports it.
 *
 * `sheet` is an index: the workbook's sheets in order, then the sheets
 * the run added, in the order it added them.
 */
export type ScriptOp =
  | { readonly kind: 'write'; readonly sheet: number; readonly row: number; readonly column: number; readonly value: ScriptValue }
  | {
      readonly kind: 'format';
      readonly sheet: number;
      readonly firstRow: number;
      readonly firstColumn: number;
      readonly lastRow: number;
      readonly lastColumn: number;
      readonly change: ScriptFormat;
    }
  | { readonly kind: 'addSheet'; readonly name: string };

/** Worker to host: what the run did. */
export type ScriptWorkerMessage =
  | { readonly type: 'ready' }
  | { readonly type: 'done'; readonly ops: readonly ScriptOp[]; readonly log: readonly string[] }
  | { readonly type: 'failed'; readonly message: string; readonly log: readonly string[] };
