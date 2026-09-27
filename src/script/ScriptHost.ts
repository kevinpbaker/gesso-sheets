import type { ScriptRunMessage, ScriptValue, ScriptWorkerMessage, ScriptWrite } from './protocol';

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
  /** How many cells a run may write. */
  readonly writes: number;
  /** The sheet's extent, which no write may fall outside. */
  readonly rows: number;
  readonly columns: number;
}

export const DEFAULT_LIMITS: ScriptLimits = { milliseconds: 5_000, writes: 100_000, rows: 10_000, columns: 100 };

export type ScriptRunResult =
  | { readonly outcome: 'done'; readonly writes: readonly ScriptWrite[]; readonly log: readonly string[] }
  | { readonly outcome: 'failed'; readonly message: string; readonly log: readonly string[] }
  | { readonly outcome: 'timeout' }
  /** Not run at all: a file's script nobody confirmed, or a worker that spoke out of turn. */
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
  constructor(
    private readonly spawn: () => ScriptWorker,
    private readonly limits: ScriptLimits = DEFAULT_LIMITS
  ) {}

  run(
    script: Script,
    cells: Readonly<Record<string, ScriptValue>>,
    options: { readonly confirmed?: boolean } = {}
  ): Promise<ScriptRunResult> {
    if (script.origin.kind === 'file' && options.confirmed !== true) {
      return Promise.resolve({
        outcome: 'refused',
        reason: `This script came with ${script.origin.file}. It runs only when you choose to run it.`
      });
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
        clearTimeout(timer);
        // One run, one worker: whatever the script left behind — a
        // promise, a pending job — goes with it.
        worker.terminate();
        resolve(result);
      };
      const timer = setTimeout(() => finish({ outcome: 'timeout' }), limits.milliseconds);
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
        const problem = checkWrites(message.writes, limits);
        finish(problem === null ? { outcome: 'done', writes: message.writes, log: logOf(message.log) } : { outcome: 'refused', reason: problem });
      });
      worker.post({
        type: 'run',
        source: script.source,
        cells,
        limits: { writes: limits.writes, rows: limits.rows, columns: limits.columns }
      });
    });
  }
}

/** What is wrong with a worker's writes, or null when nothing is. */
export function checkWrites(writes: unknown, limits: ScriptLimits): string | null {
  if (!Array.isArray(writes)) {
    return 'The script worker reported its writes in a shape that is not a list.';
  }
  if (writes.length > limits.writes) {
    return `The script wrote ${writes.length} cells, and a run may write at most ${limits.writes}.`;
  }
  for (const write of writes) {
    const candidate = write as Partial<ScriptWrite> | null;
    const row = candidate?.row;
    const column = candidate?.column;
    const value = candidate?.value;
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      (row as number) < 0 ||
      (column as number) < 0 ||
      (row as number) >= limits.rows ||
      (column as number) >= limits.columns
    ) {
      return 'The script reported a write outside the sheet.';
    }
    if (!(value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')) {
      return 'The script reported a value no cell can hold.';
    }
  }
  return null;
}

function logOf(log: unknown): readonly string[] {
  return Array.isArray(log) ? log.slice(0, 200).map(line => String(line).slice(0, 1000)) : [];
}
