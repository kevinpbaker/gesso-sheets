import { Worker } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';

import type { ScriptBook, ScriptRunMessage } from './protocol';
import { checkOps, DEFAULT_LIMITS, ScriptHost, type Script, type ScriptLimits, type ScriptWorker } from './ScriptHost';

/**
 * The script security model, as specs — Phase 29.
 *
 * One `describe` per question `SCRIPTS.md` answers, each failing if the
 * answer stops being true. They run the real worker file on a Node
 * worker thread, which has `fetch`, `process`, `require` through
 * `module`, timers and the rest as real globals — so "a script cannot
 * reach them" is a claim about the sandbox and not about Node.
 */

const LIMITS: ScriptLimits = { ...DEFAULT_LIMITS, milliseconds: 3_000, writes: 50, rows: 100, columns: 10, sheets: 2 };

const BOOK: ScriptBook = {
  sheets: [
    { name: 'Sales', cells: { '0:0': 'Region', '1:0': 42, '2:0': 8, '0:1': 'Total' } },
    { name: 'Notes', cells: { '0:0': 'kept' } }
  ],
  active: 0
};

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
  const result = await host.run(typed(source), BOOK);
  return { result, workers };
}

/** The first write's value, or the reason there was none. */
function firstValue(result: Awaited<ReturnType<typeof run>>['result']): unknown {
  if (result.outcome !== 'done') {
    return result.outcome === 'failed' ? `failed: ${result.message}` : result.outcome;
  }
  const op = result.ops[0];
  return op?.kind === 'write' ? op.value : op;
}

describe('where a script runs: a worker of its own, one per run', () => {
  it('runs, and its worker is ended when it has', async () => {
    const { result, workers } = await run('sheet.write("B1", sheet.read("A2") * 2)');
    expect(result).toEqual({ outcome: 'done', ops: [{ kind: 'write', sheet: 0, row: 0, column: 1, value: 84 }], log: [] });
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
        expect(firstValue(result)).toBe('undefined');
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
    const formatted = await run('sheet.range("A1:J6").format({ bold: true })');
    expect(formatted.result.outcome === 'failed' && formatted.result.message).toContain('at most 50');
    const sheets = await run('workbook.addSheet("One"); workbook.addSheet("Two"); workbook.addSheet("Three")');
    expect(sheets.result.outcome === 'failed' && sheets.result.message).toContain('at most 2 sheets');
  });

  it('refuses a write outside the sheet', async () => {
    const { result } = await run('sheet.write("Z1", 1)');
    expect(result.outcome).toBe('failed');
  });

  it('refuses a value no cell can hold, and a format the model does not list', async () => {
    for (const source of [
      'sheet.write("A1", {})',
      'sheet.write("A1", NaN)',
      'sheet.range("A1").format({ fill: "url(https://example.com)" })',
      'sheet.range("A1").format({ onClick: "x" })'
    ]) {
      const { result } = await run(source);
      expect(result.outcome, source).toBe('failed');
    }
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
    const done = (...ops: unknown[]) => ({ type: 'done', ops, log: [] });
    const write = (row: number, column = 0, sheet = 0) => ({ kind: 'write', sheet, row, column, value: 1 });
    const format = (change: unknown, lastRow = 0) => ({ kind: 'format', sheet: 0, firstRow: 0, firstColumn: 0, lastRow, lastColumn: 0, change });
    const lies = {
      'a write outside the sheet': done(write(5_000)),
      'a write to a sheet that does not exist': done(write(0, 0, 2)),
      'too many writes': done(...Array.from({ length: 51 }, (_, row) => write(row % 100))),
      'too many cells formatted': done(format({ bold: true }, 99)),
      'a value that is not a value': done({ kind: 'write', sheet: 0, row: 0, column: 0, value: { toString: () => 'x' } }),
      'a format field nobody listed': done(format({ bold: true, script: 'alert(1)' })),
      'a colour that is not one': done(format({ color: 'red; background: url(x)' })),
      'too many sheets': done({ kind: 'addSheet', name: 'A' }, { kind: 'addSheet', name: 'B' }, { kind: 'addSheet', name: 'C' }),
      'a kind of change nobody made': done({ kind: 'deleteEverything' }),
      'something that is not a report at all': { type: 'shell', command: 'rm -rf /' }
    };
    for (const [lie, report] of Object.entries(lies)) {
      const result = await new ScriptHost(lying(report), LIMITS).run(script, BOOK);
      expect(result.outcome, lie).toBe('refused');
    }
    // And the truth is believed: a write to a sheet the run added is to a sheet that exists.
    expect(checkOps([{ kind: 'addSheet', name: 'New' }, write(0, 0, 2), format({ bold: true, number: 'currency' })], 2, LIMITS)).toBeNull();
  });
});

describe('where scripts come from: a file’s runs only when someone chooses to run it', () => {
  it('refuses a file’s script nobody confirmed, without starting a worker', async () => {
    const workers = nodeWorkers();
    const host = new ScriptHost(workers.spawn, LIMITS);
    const fromFile: Script = { name: 'Tidy', source: 'sheet.write("A1", 1)', origin: { kind: 'file', file: 'budget.gsheet' } };

    const refused = await host.run(fromFile, BOOK);
    expect(refused).toMatchObject({ outcome: 'refused' });
    expect(refused.outcome === 'refused' && refused.reason).toContain('budget.gsheet');
    expect(workers.started).toHaveLength(0);

    const confirmed = await host.run(fromFile, BOOK, { confirmed: true });
    expect(confirmed.outcome).toBe('done');
  });
});

describe('stopping a run', () => {
  it('ends the worker when asked, keeps nothing, and runs one script at a time', async () => {
    const workers = nodeWorkers();
    const host = new ScriptHost(workers.spawn, LIMITS);
    const running = host.run(typed('sheet.write("A1", 1); for (;;) {}'), BOOK);
    expect(host.running).toBe(true);
    expect(await host.run(typed('sheet.write("A1", 2)'), BOOK)).toMatchObject({ outcome: 'refused' });
    host.stop();
    expect(await running).toEqual({ outcome: 'stopped' });
    expect(host.running).toBe(false);
    expect(workers.ended()).toBe(1);
  });
});

/**
 * The API itself — Phase 30. What `api.ts` promises, through the real
 * worker: ranges, values as rows, formats, other sheets and new ones.
 */
describe('the API a script is handed', () => {
  const ops = async (source: string) => {
    const { result } = await run(source);
    if (result.outcome !== 'done') {
      throw new Error(result.outcome === 'failed' ? result.message : result.outcome);
    }
    return { ops: result.ops, log: result.log };
  };

  it('reads a range as rows, with nothing as null', async () => {
    const { log } = await ops('console.log(JSON.stringify(sheet.range("A1:B3").values))');
    expect(log).toEqual(['[["Region","Total"],[42,null],[8,null]]']);
  });

  it('fills a column from another and formats it', async () => {
    const { ops: done } = await ops(`
      const source = sheet.range("A2:A3").values;
      sheet.range("B2:B3").write(source.map(([n]) => [n * 2]));
      sheet.range("B2:B3").format({ bold: true, number: "currency" });
    `);
    expect(done).toEqual([
      { kind: 'write', sheet: 0, row: 1, column: 1, value: 84 },
      { kind: 'write', sheet: 0, row: 2, column: 1, value: 16 },
      { kind: 'format', sheet: 0, firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 1, change: { bold: true, number: 'currency' } }
    ]);
  });

  it('reads back what it wrote, in the same run', async () => {
    const { log } = await ops('sheet.write("C1", "=A2*2"); sheet.range("D1:D2").write(7); console.log(sheet.read("C1"), sheet.read("D2"))');
    expect(log).toEqual(['=A2*2 7']);
  });

  it('refuses rows that do not fit, rather than cropping them', async () => {
    const { result } = await run('sheet.range("A1:B1").write([[1, 2, 3]])');
    expect(result.outcome === 'failed' && result.message).toContain('does not fit in A1:B1');
  });

  it('reaches another sheet by name, and adds one', async () => {
    const { ops: done, log } = await ops(`
      console.log(workbook.sheets.map(s => s.name).join(","), workbook.sheet("notes").read("A1"));
      const added = workbook.addSheet("Summary");
      added.write("A1", workbook.sheet("Sales").read("A2"));
    `);
    expect(log).toEqual(['Sales,Notes kept']);
    expect(done).toEqual([
      { kind: 'addSheet', name: 'Summary' },
      { kind: 'write', sheet: 2, row: 0, column: 0, value: 42 }
    ]);
  });

  it('says what went wrong in words', async () => {
    for (const [source, words] of [
      ['workbook.sheet("Nowhere")', 'no sheet called "Nowhere"'],
      ['sheet.range("A1:")', 'is not a cell address'],
      ['workbook.addSheet("sales")', 'already a sheet called'],
      ['sheet.range("A1").format({ bold: "yes" })', 'bold is true or false']
    ]) {
      const { result } = await run(source);
      expect(result.outcome === 'failed' && result.message, source).toContain(words);
    }
  });
});
