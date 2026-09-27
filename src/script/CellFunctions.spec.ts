import { beforeAll, describe, expect, it } from 'vitest';

import { CellFunctions, DEFAULT_FUNCTION_LIMITS, loadInterpreter } from './CellFunctions';

/**
 * Functions in script, the interpreter — Phase 31.
 *
 * The real QuickJS, compiled to WebAssembly, as the application worker
 * will run it. What it can reach, what stops it, and whether a stopped
 * call leaves the next one working; then what a call costs.
 */

let module: Awaited<ReturnType<typeof loadInterpreter>>;
beforeAll(async () => {
  module = await loadInterpreter();
});

function functions(source: string, limits = DEFAULT_FUNCTION_LIMITS) {
  const host = new CellFunctions(module, limits);
  const defined = host.define(source);
  if ('error' in defined) {
    throw new Error(defined.error);
  }
  return host;
}

describe('a function in the interpreter', () => {
  it('is called by name, with values in and a value out', () => {
    const host = functions('function TAX(amount, rate) { return amount * rate; }\nfunction HELLO(name) { return "Hello, " + name; }');
    expect(host.names).toEqual(['TAX', 'HELLO']);
    expect(host.call('TAX', [200, 0.2])).toEqual({ ok: true, value: 40 });
    expect(host.call('HELLO', ['Priya'])).toEqual({ ok: true, value: 'Hello, Priya' });
    expect(host.call('NOPE', [])).toMatchObject({ ok: false, kind: 'missing' });
  });

  it('takes a range as rows, and returns rows to spill', () => {
    const host = functions('function DOUBLED(rows) { return rows.map(row => row.map(v => typeof v === "number" ? v * 2 : v)); }');
    expect(host.call('DOUBLED', [[[1, 'a'], [3, null]]])).toEqual({ ok: true, value: [[2, 'a'], [6, null]] });
  });

  it('turns a result no cell can hold into an error, not a value', () => {
    const host = functions('function OBJ() { return { a: 1 }; }\nfunction FN() { return () => 1; }\nfunction NOTHING() {}');
    expect(host.call('OBJ', [])).toMatchObject({ ok: false, kind: 'result' });
    expect(host.call('FN', [])).toMatchObject({ ok: false, kind: 'result' });
    expect(host.call('NOTHING', [])).toEqual({ ok: true, value: null });
  });

  it('says what a function threw', () => {
    const host = functions('function CHECK(n) { if (n < 0) throw new Error("negative"); return n; }');
    expect(host.call('CHECK', [-1])).toEqual({ ok: false, kind: 'thrown', message: 'negative' });
    expect(host.call('CHECK', [1])).toEqual({ ok: true, value: 1 });
  });
});

describe('what a function can reach: its arguments, and nothing of the host', () => {
  it('finds none of the host’s globals', () => {
    const probes = ['fetch', 'XMLHttpRequest', 'WebSocket', 'postMessage', 'self', 'importScripts', 'process', 'require', 'setTimeout', 'setInterval', 'queueMicrotask', 'performance', 'console', 'std', 'os', 'Worker', 'indexedDB'];
    const host = functions(`function REACH(name) { return typeof globalThis[name]; }`);
    for (const probe of probes) {
      expect(host.call('REACH', [probe]), probe).toEqual({ ok: true, value: 'undefined' });
    }
  });

  it('cannot load code, since there is nothing to load it from', () => {
    const host = functions('function LOAD() { return import("node:fs"); }\nfunction EVAL() { return eval("1 + 1"); }');
    // `import()` has no module loader to answer it; eval is the interpreter's own and reaches nothing new.
    expect(host.call('LOAD', [])).toMatchObject({ ok: false });
    expect(host.call('EVAL', [])).toEqual({ ok: true, value: 2 });
  });

  it('has no randomness and no clock, so its value cannot move on its own', () => {
    const host = functions(`
      function RANDOM() { return typeof Math.random; }
      function NOW() { return Date.now(); }
      function TODAY() { return new Date().getFullYear(); }
      function YEAR(y, m, d) { return new Date(y, m - 1, d).getFullYear(); }
    `);
    expect(host.call('RANDOM', [])).toEqual({ ok: true, value: 'undefined' });
    expect(host.call('NOW', [])).toMatchObject({ ok: false, kind: 'thrown' });
    expect(host.call('TODAY', [])).toMatchObject({ ok: false, kind: 'thrown' });
    expect(host.call('YEAR', [2026, 9, 26])).toEqual({ ok: true, value: 2026 });
  });

  it('starts from nothing each time the module is defined again', () => {
    const host = functions('var left = "behind"; function A() { return typeof left; }');
    expect(host.call('A', [])).toEqual({ ok: true, value: 'string' });
    host.define('function B() { return typeof left; }');
    expect(host.names).toEqual(['B']);
    expect(host.call('B', [])).toEqual({ ok: true, value: 'undefined' });
    expect(host.call('A', [])).toMatchObject({ ok: false, kind: 'missing' });
  });
});

describe('how much a function can do: limits the interpreter enforces', () => {
  it('ends a call that loops at its deadline, and the next call works', () => {
    const host = functions('function LOOP() { for (;;) {} }\nfunction ONE() { return 1; }', { ...DEFAULT_FUNCTION_LIMITS, milliseconds: 30 });
    const started = performance.now();
    expect(host.call('LOOP', [])).toMatchObject({ ok: false, kind: 'deadline' });
    expect(performance.now() - started).toBeLessThan(500);
    expect(host.call('ONE', [])).toEqual({ ok: true, value: 1 });
  });

  it('ends a call that allocates at its memory limit, and the next call works', () => {
    const host = functions(
      'function HOARD() { const all = []; for (;;) all.push(new Array(10000).fill(1)); }\nfunction ONE() { return 1; }',
      { ...DEFAULT_FUNCTION_LIMITS, milliseconds: 5_000, memoryBytes: 8 * 1024 * 1024 }
    );
    expect(host.call('HOARD', [])).toMatchObject({ ok: false, kind: 'memory' });
    expect(host.call('ONE', [])).toEqual({ ok: true, value: 1 });
  });

  it('ends a call that recurses without end at its stack, and the next call works', () => {
    const host = functions('function DEEP(n) { return DEEP(n + 1) + 1; }\nfunction ONE() { return 1; }');
    // The interpreter's own error, in the cell: not a RangeError on the host's thread.
    expect(host.call('DEEP', [0])).toEqual({ ok: false, kind: 'thrown', message: 'stack overflow' });
    expect(host.call('ONE', [])).toEqual({ ok: true, value: 1 });
  });

  it('refuses a module that loops while it is being defined', () => {
    const host = new CellFunctions(module, { ...DEFAULT_FUNCTION_LIMITS, milliseconds: 30 });
    expect(host.define('for (;;) {}')).toMatchObject({ error: expect.stringMatching(/interrupted/i) });
    expect(host.names).toEqual([]);
  });
});

/**
 * What a call costs. Printed as well as bounded, because the roadmap's
 * exit is the numbers: a few microseconds a call, or Part Six stops.
 */
describe('what a call costs', () => {
  const time = (label: string, calls: number, run: () => void): number => {
    for (let at = 0; at < Math.min(calls, 2_000); at++) run();
    const started = performance.now();
    for (let at = 0; at < calls; at++) run();
    const each = ((performance.now() - started) * 1000) / calls;
    process.stderr.write(`  ${label}: ${each.toFixed(2)}µs a call\n`);
    return each;
  };

  it('is a few microseconds with a number in and out, and more with a range', async () => {
    const readyStarted = performance.now();
    const fresh = new CellFunctions(await loadInterpreter());
    fresh.define('function TAX(amount, rate) { return amount * rate; }');
    process.stderr.write(`  a runtime and a module defined: ${(performance.now() - readyStarted).toFixed(2)}ms\n`);

    const host = functions('function TAX(amount, rate) { return amount * rate; }\nfunction SUM(rows) { let t = 0; for (const r of rows) for (const v of r) t += v; return t; }');
    const scalar = time('two numbers in, one out', 50_000, () => host.call('TAX', [200, 0.2]));
    const rows = Array.from({ length: 100 }, (_, row) => [row]);
    const range = time('a 100-row range in, one out', 5_000, () => host.call('SUM', [rows]));
    expect(host.call('SUM', [rows])).toEqual({ ok: true, value: 4950 });
    // Generous against a loaded machine; the printed numbers are the finding.
    expect(scalar).toBeLessThan(50);
    expect(range).toBeLessThan(500);
  });
});
