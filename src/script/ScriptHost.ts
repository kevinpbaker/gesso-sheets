import { SCRIPT_NUMBER_FORMATS, type ScriptBook, type ScriptOp, type ScriptRunMessage, type ScriptWorkerMessage } from './protocol';

/**
 * A script, and where it came from.
 *
 * Where it came from is the part that matters. A script somebody typed
 * is theirs; a script that came in a file is a stranger's code, and the
 * first rule of Part Five is that nothing in a file runs by itself.
 */
export interface Script {
  readonly name: string;
  readonly source: string;
  /** `typed` in this sheet, or brought in with a file, named. */
  readonly origin: { readonly kind: 'typed' } | { readonly kind: 'file'; readonly file: string };
}

/** The worker, as the host sees it: something it can talk to and kill. */
export interface ScriptWorker {
  post(message: ScriptRunMessage): void;
  onMessage(listener: (message: unknown) => void): void;
  terminate(): void;
}

export interface ScriptLimits {
  /** How long a run may take, in milliseconds, before its worker is killed. */
  readonly milliseconds: number;
  /** How many cells a run may write or format, counted together. */
  readonly writes: number;
  /** How many sheets a run may add. */
  readonly sheets: number;
  /** The sheet's extent, which no write may fall outside. */
  readonly rows: number;
  readonly columns: number;
}

export const DEFAULT_LIMITS: ScriptLimits = { milliseconds: 5_000, writes: 100_000, rows: 10_000, columns: 100, sheets: 20 };

export type ScriptRunResult =
  | { readonly outcome: 'done'; readonly ops: readonly ScriptOp[]; readonly log: readonly string[] }
  | { readonly outcome: 'failed'; readonly message: string; readonly log: readonly string[] }
  | { readonly outcome: 'timeout' }
  /** Ended by `stop`, which is the person pressing Stop. */
  | { readonly outcome: 'stopped' }
  /**
   * Not run at all, or not believed: a file's script nobody confirmed,
   * a run while another is running, or a worker that spoke out of turn.
   */
  | { readonly outcome: 'refused'; readonly reason: string };

/**
 * Runs scripts, one worker each, and believes nothing a worker says.
 *
 * The host is the trusted side. It decides whether a script may run —
 * a file's only with the person's say-so, each time — starts a worker
 * for the run and ends it after, kills it when the time is up, and
 * checks every write it reports against the limits before handing the
 * list back to be applied as one transaction. A worker is where the
 * script's code runs, so a worker is assumed to be hostile: one that
 * reports a write outside the sheet, too many writes, or a value that
 * is not a cell's, has its whole run refused, not trimmed.
 */
export class ScriptHost {
  /** Ends the run in progress, or null when there is none. */
  private ending: (() => void) | null = null;

  constructor(
    private readonly spawn: () => ScriptWorker,
    private readonly limits: ScriptLimits = DEFAULT_LIMITS
  ) {}

  /** Whether a run is in progress. */
  get running(): boolean {
    return this.ending !== null;
  }

  /** Ends the run in progress, keeping none of it. */
  stop(): void {
    this.ending?.();
  }

  run(script: Script, book: ScriptBook, options: { readonly confirmed?: boolean } = {}): Promise<ScriptRunResult> {
    if (script.origin.kind === 'file' && options.confirmed !== true) {
      return Promise.resolve({
        outcome: 'refused',
        reason: `This script came with ${script.origin.file}. It runs only when you choose to run it.`
      });
    }
    if (this.ending !== null) {
      return Promise.resolve({ outcome: 'refused', reason: 'Another script is running. Stop it, or wait for it to finish.' });
    }
    const worker = this.spawn();
    const limits = this.limits;
    return new Promise<ScriptRunResult>(resolve => {
      let settled = false;
      const finish = (result: ScriptRunResult): void => {
        if (settled) {
          return;
        }
        settled = true;
        this.ending = null;
        clearTimeout(timer);
        // One run, one worker: whatever the script left behind — a
        // promise, a pending job — goes with it.
        worker.terminate();
        resolve(result);
      };
      const timer = setTimeout(() => finish({ outcome: 'timeout' }), limits.milliseconds);
      this.ending = () => finish({ outcome: 'stopped' });
      worker.onMessage(raw => {
        const message = raw as ScriptWorkerMessage;
        if (message?.type === 'ready') {
          return;
        }
        if (message?.type === 'failed') {
          finish({ outcome: 'failed', message: String(message.message).slice(0, 1000), log: logOf(message.log) });
          return;
        }
        if (message?.type !== 'done') {
          finish({ outcome: 'refused', reason: 'The script worker sent something it should not have.' });
          return;
        }
        const problem = checkOps(message.ops, book.sheets.length, limits);
        finish(problem === null ? { outcome: 'done', ops: message.ops, log: logOf(message.log) } : { outcome: 'refused', reason: problem });
      });
      worker.post({
        type: 'run',
        source: script.source,
        book,
        limits: { writes: limits.writes, rows: limits.rows, columns: limits.columns, sheets: limits.sheets }
      });
    });
  }
}

const COLOUR = /^#[0-9a-f]{6}$/i;
const ALIGNS: readonly unknown[] = ['auto', 'start', 'center', 'end'];
const NUMBER_FORMATS: readonly unknown[] = SCRIPT_NUMBER_FORMATS;

/**
 * What is wrong with what a worker says its run did, or null when
 * nothing is.
 *
 * Every field of every change, checked as if a stranger wrote it,
 * because one did: a change to a sheet that does not exist, a cell
 * outside the sheet, a value no cell holds, a format field the model
 * does not list, more cells than a run may touch or more sheets than it
 * may add. `sheets` is how many the workbook had when the run began.
 */
export function checkOps(ops: unknown, sheets: number, limits: ScriptLimits): string | null {
  if (!Array.isArray(ops)) {
    return 'The script worker reported its changes in a shape that is not a list.';
  }
  let spent = 0;
  let known = sheets;
  let added = 0;
  const inside = (row: unknown, column: unknown): boolean =>
    Number.isInteger(row) &&
    Number.isInteger(column) &&
    (row as number) >= 0 &&
    (column as number) >= 0 &&
    (row as number) < limits.rows &&
    (column as number) < limits.columns;
  for (const raw of ops) {
    if (typeof raw !== 'object' || raw === null) {
      return 'The script worker reported a change that is not one.';
    }
    const op = raw as Record<string, unknown>;
    if (op.kind === 'addSheet') {
      if (typeof op.name !== 'string' || op.name.trim() === '' || op.name.length > 31) {
        return 'The script reported a sheet with a name no sheet can have.';
      }
      if (++added > limits.sheets) {
        return `The script added more than ${limits.sheets} sheets.`;
      }
      known++;
      continue;
    }
    if (!Number.isInteger(op.sheet) || (op.sheet as number) < 0 || (op.sheet as number) >= known) {
      return 'The script reported a change to a sheet that does not exist.';
    }
    if (op.kind === 'write') {
      if (!inside(op.row, op.column)) {
        return 'The script reported a write outside the sheet.';
      }
      const value = op.value;
      if (!(value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value)))) {
        return 'The script reported a value no cell can hold.';
      }
      spent += 1;
    } else if (op.kind === 'format') {
      if (
        !inside(op.firstRow, op.firstColumn) ||
        !inside(op.lastRow, op.lastColumn) ||
        (op.firstRow as number) > (op.lastRow as number) ||
        (op.firstColumn as number) > (op.lastColumn as number)
      ) {
        return 'The script reported a format outside the sheet.';
      }
      const problem = checkFormat(op.change);
      if (problem !== null) {
        return problem;
      }
      spent += ((op.lastRow as number) - (op.firstRow as number) + 1) * ((op.lastColumn as number) - (op.firstColumn as number) + 1);
    } else {
      return 'The script worker reported a kind of change there is no such thing as.';
    }
    if (spent > limits.writes) {
      return `The script changed more than ${limits.writes} cells, which is as many as a run may.`;
    }
  }
  return null;
}

function checkFormat(change: unknown): string | null {
  if (typeof change !== 'object' || change === null || Array.isArray(change)) {
    return 'The script reported a format that is not one.';
  }
  for (const [key, value] of Object.entries(change)) {
    const fine =
      key === 'bold' || key === 'italic' || key === 'underline' || key === 'wrap'
        ? typeof value === 'boolean'
        : key === 'color' || key === 'fill'
          ? typeof value === 'string' && COLOUR.test(value)
          : key === 'align'
            ? ALIGNS.includes(value)
            : key === 'fontSize'
              ? typeof value === 'number' && value >= 6 && value <= 72
              : key === 'number'
                ? NUMBER_FORMATS.includes(value)
                : false;
    if (!fine) {
      return `The script reported a format the model does not allow: ${key.slice(0, 40)}.`;
    }
  }
  return null;
}

function logOf(log: unknown): readonly string[] {
  return Array.isArray(log) ? log.slice(0, 200).map(line => String(line).slice(0, 1000)) : [];
}
