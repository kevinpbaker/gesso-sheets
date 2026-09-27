/**
 * What crosses between a script's host and the worker a script runs in.
 *
 * Plain data both ways, and the host trusts none of what comes back:
 * see `ScriptHost`, which checks every write a worker reports before it
 * lets one near the workbook. The worker is where untrusted code runs,
 * so the worker is the other side of the trust boundary, and its
 * messages are treated the way a server treats a request.
 */

/** A value a script may write into a cell: what a cell can hold, and nothing with behaviour. */
export type ScriptValue = string | number | boolean | null;

/** One cell a script asked to change, as the worker reports it. */
export interface ScriptWrite {
  readonly row: number;
  readonly column: number;
  readonly value: ScriptValue;
}

/** Host to worker: run this, against these cells, within these limits. */
export interface ScriptRunMessage {
  readonly type: 'run';
  readonly source: string;
  /** The cells the script may read, by `row:column`, as values. */
  readonly cells: Readonly<Record<string, ScriptValue>>;
  readonly limits: { readonly writes: number; readonly rows: number; readonly columns: number };
}

/** Worker to host: what the run did. */
export type ScriptWorkerMessage =
  | { readonly type: 'ready' }
  | { readonly type: 'done'; readonly writes: readonly ScriptWrite[]; readonly log: readonly string[] }
  | { readonly type: 'failed'; readonly message: string; readonly log: readonly string[] };
