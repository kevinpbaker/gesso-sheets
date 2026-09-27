/**
 * The worker a script runs in: one per run, and thrown away after it.
 *
 * The security model is in `SCRIPTS.md`. What this file does about it:
 *
 * 1. **Locks the realm down before anything else runs.** `lockdown()`
 *    freezes every shared intrinsic — `Object.prototype`, `Array`,
 *    `Function` — so nothing a script does can change how the rest of
 *    this worker behaves, and removes the ones that reach outside.
 * 2. **Evaluates the script in a `Compartment` whose only globals are
 *    the ones handed to it**: `sheet` and `console`. No `fetch`, no
 *    `postMessage`, no `self`, no `importScripts`, no timers, no clock,
 *    no randomness. SES rejects `import(…)` and direct `eval` in the
 *    source outright, which is what a list of deleted globals cannot
 *    do: dynamic import is syntax, not a global, and a worker that only
 *    deleted `fetch` could still load and run code from anywhere.
 * 3. **Counts writes as they are made**, and stops the script at the
 *    limit — but the host counts again, because this side is where the
 *    untrusted code is and nothing it says is believed.
 *
 * What it cannot do is stop a script that never returns, or one that
 * allocates until the tab runs out of memory. The host's time limit
 * terminates this worker for the first; the second is the one limit a
 * browser gives no way to set, and the model says so.
 *
 * Self-contained on purpose: it imports `ses` and types and nothing
 * else, so the same file is the browser's worker and, in the specs,
 * a Node worker thread given the two globals it needs.
 */
import 'ses';

import type { ScriptRunMessage, ScriptValue, ScriptWorkerMessage, ScriptWrite } from './protocol';

interface WorkerScope {
  postMessage(message: ScriptWorkerMessage): void;
  onmessage: ((event: { data: ScriptRunMessage }) => void) | null;
}

// Taken before lockdown, and never handed to a script.
const scope = globalThis as unknown as WorkerScope;
const post = scope.postMessage.bind(scope);

lockdown({ errorTaming: 'unsafe', overrideTaming: 'severe' });

scope.onmessage = event => {
  const message = event.data;
  if (message?.type !== 'run') {
    return;
  }
  // One run per worker. A second message is ignored, because the worker
  // is terminated after its run and a message to it is somebody else's.
  scope.onmessage = null;
  post(run(message));
};

post({ type: 'ready' });

function run(message: ScriptRunMessage): ScriptWorkerMessage {
  const writes: ScriptWrite[] = [];
  const log: string[] = [];
  const { limits } = message;

  const cellAt = (address: unknown): { row: number; column: number } => {
    if (typeof address !== 'string') {
      throw new TypeError('A cell is written as an address, like "B4".');
    }
    const match = /^\s*([A-Za-z]{1,3})(\d{1,7})\s*$/.exec(address);
    if (match === null) {
      throw new TypeError(`"${address}" is not a cell address.`);
    }
    let column = 0;
    for (const letter of match[1].toUpperCase()) {
      column = column * 26 + (letter.charCodeAt(0) - 64);
    }
    const at = { row: Number(match[2]) - 1, column: column - 1 };
    if (at.row < 0 || at.row >= limits.rows || at.column < 0 || at.column >= limits.columns) {
      throw new RangeError(`${address} is outside the sheet.`);
    }
    return at;
  };

  const sheet = harden({
    read(address: unknown): ScriptValue {
      const at = cellAt(address);
      return message.cells[`${at.row}:${at.column}`] ?? null;
    },
    write(address: unknown, value: unknown): void {
      const at = cellAt(address);
      if (!(value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')) {
        throw new TypeError('A cell holds text, a number, true or false, or nothing.');
      }
      if (writes.length >= limits.writes) {
        throw new RangeError(`A script may write at most ${limits.writes} cells in one run.`);
      }
      writes.push({ ...at, value });
    }
  });
  const console = harden({
    log: (...parts: unknown[]) => {
      if (log.length < 200) {
        log.push(parts.map(part => String(part)).join(' ').slice(0, 1000));
      }
    }
  });

  try {
    const compartment = new Compartment({ globals: { sheet, console }, __options__: true });
    compartment.evaluate(message.source);
    return { type: 'done', writes, log };
  } catch (error) {
    return { type: 'failed', message: error instanceof Error ? error.message : String(error), log };
  }
}
