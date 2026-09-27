import { Worker } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';

import type { ScriptRunMessage } from './protocol';
import { checkWrites, DEFAULT_LIMITS, ScriptHost, type Script, type ScriptLimits, type ScriptWorker } from './ScriptHost';

/**
 * The script security model, as specs — Phase 29.
 *
 * One `describe` per question `SCRIPTS.md` answers, each failing if the
 * answer stops being true. They run the real worker file on a Node
 * worker thread, which has `fetch`, `process`, `require` through
 * `module`, timers and the rest as real globals — so "a script cannot
 * reach them" is a claim about the sandbox and not about Node.
 */

const LIMITS: ScriptLimits = { ...DEFAULT_LIMITS, milliseconds: 3_000, writes: 50, rows: 100, columns: 10 };

/** A real script worker, on a Node thread, counting what was started and ended. */
function nodeWorkers() {
  const started: Worker[] = [];
  let ended = 0;
  const spawn = (): ScriptWorker => {
    const worker = new Worker(new URL('./nodeWorker.mjs', import.meta.url));
    started.push(worker);
    return {
      post: message => worker.postMessage(message),
      onMessage: listener => worker.on('message', listener),
      terminate: () => {
        ended++;
        void worker.terminate();
      }
    };
  };
  return { spawn, started, ended: () => ended };
}

const typed = (source: string): Script => ({ name: 'spec', source, origin: { kind: 'typed' } });

async function run(source: string, limits = LIMITS) {
  const workers = nodeWorkers();
  const host = new ScriptHost(workers.spawn, limits);
  const result = await host.run(typed(source), { '0:0': 'Region', '1:0': 42 });
  return { result, workers };
}

describe('where a script runs: a worker of its own, one per run', () => {
  it('runs, and its worker is ended when it has', async () => {
    const { result, workers } = await run('sheet.write("B1", sheet.read("A2") * 2)');
    expect(result).toEqual({ outcome: 'done', writes: [{ row: 0, column: 1, value: 84 }], log: [] });
    expect(workers.started).toHaveLength(1);
    expect(workers.ended()).toBe(1);
  });

  it('is stopped by its time limit, and its writes are not kept', async () => {
    const { result, workers } = await run('sheet.write("A1", "partial"); for (;;) {}', { ...LIMITS, milliseconds: 1_500 });
    expect(result).toEqual({ outcome: 'timeout' });
    expect(workers.ended()).toBe(1);
  }, 10_000);
});

describe('what a script can reach: the sheet and nothing else', () => {
  const unreachable: Record<string, string> = {
    fetch: 'typeof fetch',
    'the worker’s postMessage': 'typeof postMessage',
    'the worker itself': 'typeof self',
    'globalThis’s own fetch': 'typeof globalThis.fetch',
    process: 'typeof process',
    require: 'typeof require',
    timers: 'typeof setTimeout',
    'a WebSocket': 'typeof WebSocket',
    'a clock': 'Date.now()',
    'randomness': 'Math.random()'
  };
  for (const [what, probe] of Object.entries(unreachable)) {
    it(`cannot reach ${what}`, async () => {
      const { result } = await run(`sheet.write("A1", String(${probe}))`);
      // Either the probe throws, or what it finds is nothing.
      if (result.outcome === 'done') {
        expect(result.writes[0]?.value).toBe('undefined');
      } else {
        expect(result.outcome).toBe('failed');
      }
    });
  }

  it('cannot load code, by import or by eval, even through the Function constructor', async () => {
    for (const source of ['import("node:fs")', 'eval("1")', '(() => {}).constructor("return import(\\"node:fs\\")")()']) {
      const { result } = await run(source);
      expect(result.outcome, source).toBe('failed');
    }
  });

  it('cannot change what the rest of the worker shares', async () => {
    const { result } = await run('Object.prototype.stolen = 1');
    expect(result.outcome).toBe('failed');
    const tamper = await run('sheet.write = () => {}');
    expect(tamper.result.outcome).toBe('failed');
  });
});

describe('how much a script can do: limits the host counts', () => {
  it('stops at the write limit, in the worker, and keeps nothing', async () => {
    const { result } = await run('for (let row = 1; row <= 60; row++) sheet.write("A" + row, row)');
    expect(result.outcome).toBe('failed');
    expect(result.outcome === 'failed' && result.message).toContain('at most 50');
  });

  it('refuses a write outside the sheet', async () => {
    const { result } = await run('sheet.write("Z1", 1)');
    expect(result.outcome).toBe('failed');
  });

  it('refuses a value no cell can hold', async () => {
    const { result } = await run('sheet.write("A1", {})');
    expect(result.outcome).toBe('failed');
  });

  /**
   * The worker is where the hostile code is, so the host checks again.
   * These are workers that lie — the shape a compromised sandbox would
   * take — and each whole run is refused, not trimmed.
   */
  it('believes nothing a worker reports', async () => {
    const lying = (report: unknown): (() => ScriptWorker) => () => {
      let listener: ((message: unknown) => void) | null = null;
      return {
        post: (_message: ScriptRunMessage) => queueMicrotask(() => listener?.(report)),
        onMessage: next => (listener = next),
        terminate: () => {}
      };
    };
    const script = typed('');
    const outOfSheet = { type: 'done', writes: [{ row: 5_000, column: 0, value: 1 }], log: [] };
    const tooMany = { type: 'done', writes: Array.from({ length: 51 }, (_, row) => ({ row, column: 0, value: 1 })), log: [] };
    const notAValue = { type: 'done', writes: [{ row: 0, column: 0, value: { toString: () => 'x' } }], log: [] };
    const nonsense = { type: 'shell', command: 'rm -rf /' };
    for (const report of [outOfSheet, tooMany, notAValue, nonsense]) {
      const result = await new ScriptHost(lying(report), LIMITS).run(script, {});
      expect(result.outcome, JSON.stringify(report).slice(0, 40)).toBe('refused');
    }
    expect(checkWrites([{ row: 0, column: 0, value: 'ok' }], LIMITS)).toBeNull();
  });
});

describe('where scripts come from: a file’s runs only when someone chooses to run it', () => {
  it('refuses a file’s script nobody confirmed, without starting a worker', async () => {
    const workers = nodeWorkers();
    const host = new ScriptHost(workers.spawn, LIMITS);
    const fromFile: Script = { name: 'Tidy', source: 'sheet.write("A1", 1)', origin: { kind: 'file', file: 'budget.gsheet' } };

    const refused = await host.run(fromFile, {});
    expect(refused).toMatchObject({ outcome: 'refused' });
    expect(refused.outcome === 'refused' && refused.reason).toContain('budget.gsheet');
    expect(workers.started).toHaveLength(0);

    const confirmed = await host.run(fromFile, {}, { confirmed: true });
    expect(confirmed.outcome).toBe('done');
  });
});
