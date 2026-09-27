import { newQuickJSWASMModuleFromVariant, type QuickJSContext, type QuickJSRuntime, type QuickJSWASMModule } from 'quickjs-emscripten-core';
import variant from '@jitl/quickjs-wasmfile-release-sync';

/**
 * Functions written in JavaScript that a formula calls — Part Six.
 *
 * Run by QuickJS, compiled to WebAssembly, on the thread that holds the
 * workbook. A call is synchronous, which is the whole reason for an
 * interpreter here rather than the script worker: the recalculation
 * needs the answer before it moves to the next cell.
 *
 * The interpreter is its own world. It has its own heap and its own
 * global object, and nothing of the host is in it: no `fetch`, no
 * worker scope, no `postMessage`, no timers. What it is handed is its
 * arguments, as plain data, and what comes back is a value, as plain
 * data. It also has what a browser worker cannot: a memory limit, and
 * an interrupt handler that ends a call at its deadline.
 *
 * Phase 31 is this class and its measurements. Phase 32 wires it into
 * the evaluator.
 */

/** What a cell holds, as a function sees it. */
export type FunctionScalar = number | string | boolean | null;
/** An argument or a result: one value, or rows of them. */
export type FunctionValue = FunctionScalar | FunctionScalar[][];

export type FunctionResult =
  | { readonly ok: true; readonly value: FunctionValue }
  | {
      readonly ok: false;
      /** `thrown` by the function, past its `deadline`, out of `memory`, or `missing`, or a result no cell can hold. */
      readonly kind: 'thrown' | 'deadline' | 'memory' | 'missing' | 'result';
      readonly message: string;
    };

export interface CellFunctionLimits {
  /** How long one call may run. */
  readonly milliseconds: number;
  /** How much the interpreter's heap may hold, for all the workbook's functions together. */
  readonly memoryBytes: number;
  /** How deep a call may recurse, in bytes of the interpreter's stack. */
  readonly stackBytes: number;
}

export const DEFAULT_FUNCTION_LIMITS: CellFunctionLimits = {
  milliseconds: 50,
  memoryBytes: 32 * 1024 * 1024,
  // 128 kB is about seven hundred and forty calls deep, and it is half
  // of what Chrome's worker stack survives. At 256 kB the host thread's
  // own stack overflowed before the interpreter's check fired, in
  // Chrome and in Node: a `RangeError` on the workbook's thread rather
  // than an error in a cell. 224 kB was the largest that held in
  // Chrome, measured with nothing else on the stack, and a call from
  // the evaluator arrives with the evaluator's frames under it.
  stackBytes: 128 * 1024
};

/**
 * Run first in every context, before a line of the workbook's code.
 *
 * A function whose value moves on its own is not one a dependency graph
 * can trust: it would show one answer, and recalculating nothing would
 * change it. So there is no randomness, and `Date` makes dates from
 * what it is told and refuses to say what time it is now.
 *
 * `__call` is how the host calls in: a name and the arguments as JSON,
 * the result back as JSON. One string each way is cheaper than building
 * a handle per argument once a range is involved, and it means a result
 * can only be data: JSON has no functions, getters or cycles.
 */
const PRELUDE = `
"use strict";
delete Math.random;
{
  const RealDate = Date;
  const refuse = () => { throw new Error("A function cannot ask what time it is: its answer would change on its own."); };
  const FixedDate = function (...args) {
    if (args.length === 0) refuse();
    return new.target === undefined ? refuse() : new RealDate(...args);
  };
  FixedDate.prototype = RealDate.prototype;
  FixedDate.UTC = RealDate.UTC;
  FixedDate.parse = RealDate.parse;
  FixedDate.now = refuse;
  globalThis.Date = FixedDate;
}
Object.defineProperty(globalThis, "__call", {
  value: (name, json) => {
    const found = globalThis[name];
    if (typeof found !== "function") return JSON.stringify({ missing: true });
    const result = found(...JSON.parse(json));
    return JSON.stringify({ value: result === undefined ? null : result });
  }
});
`;

let loading: Promise<QuickJSWASMModule> | null = null;

/** The interpreter's WebAssembly module, loaded once per thread. */
export function loadInterpreter(): Promise<QuickJSWASMModule> {
  loading ??= newQuickJSWASMModuleFromVariant(variant);
  return loading;
}

export class CellFunctions {
  private readonly runtime: QuickJSRuntime;
  private context: QuickJSContext | null = null;
  private deadline = Number.POSITIVE_INFINITY;
  private known: readonly string[] = [];
  /** Each function's handle, taken once per module rather than looked up per call. */
  private readonly handles = new Map<string, ReturnType<QuickJSContext['getProp']>>();

  constructor(
    module: QuickJSWASMModule,
    private readonly limits: CellFunctionLimits = DEFAULT_FUNCTION_LIMITS,
    private readonly now: () => number = () => performance.now()
  ) {
    this.runtime = module.newRuntime();
    this.runtime.setMemoryLimit(limits.memoryBytes);
    this.runtime.setMaxStackSize(limits.stackBytes);
    this.runtime.setInterruptHandler(() => this.now() > this.deadline);
  }

  /** The functions the module defines, by the names a formula calls them. */
  get names(): readonly string[] {
    return this.known;
  }

  /**
   * Replaces the workbook's functions with what `source` defines.
   *
   * A fresh context each time, so nothing the last module left in the
   * global object survives into the next. Every top-level function the
   * source declares becomes a name. The source runs once, under the
   * same deadline a call has, so a module that loops is refused rather
   * than hanging the thread.
   */
  define(source: string): { readonly names: readonly string[] } | { readonly error: string } {
    this.releaseHandles();
    this.context?.dispose();
    this.context = null;
    this.known = [];
    const context = this.runtime.newContext();
    const prelude = context.evalCode(PRELUDE);
    if (prelude.error !== undefined) {
      const message = String(context.dump(prelude.error)?.message ?? 'the prelude failed');
      prelude.error.dispose();
      context.dispose();
      throw new Error(message);
    }
    prelude.value.dispose();
    const before = this.globals(context);

    this.deadline = this.now() + this.limits.milliseconds;
    const ran = context.evalCode(source, 'functions.js');
    this.deadline = Number.POSITIVE_INFINITY;
    if (ran.error !== undefined) {
      const message = describe(context, ran.error);
      ran.error.dispose();
      context.dispose();
      return { error: message };
    }
    ran.value.dispose();
    const after = this.globals(context);
    this.context = context;
    this.known = after.filter(name => !before.includes(name) && /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name));
    return { names: this.known };
  }

  /** The global names that hold functions, in a context. */
  private globals(context: QuickJSContext): string[] {
    const found = context.evalCode(`Object.getOwnPropertyNames(globalThis).filter(k => typeof globalThis[k] === "function")`);
    if (found.error !== undefined) {
      found.error.dispose();
      return [];
    }
    const names = context.dump(found.value) as string[];
    found.value.dispose();
    return names;
  }

  private releaseHandles(): void {
    for (const handle of this.handles.values()) {
      handle.dispose();
    }
    this.handles.clear();
  }

  /**
   * Calls one of the module's functions. Synchronous, and bounded by the limits.
   *
   * Values with no rows in them cross as handles, one per argument, and
   * a number or a string comes back the same way: the common call, a
   * few numbers in and one out, pays for no JSON. A range crosses as
   * JSON, which is cheaper than a handle per cell.
   */
  call(name: string, args: readonly FunctionValue[]): FunctionResult {
    const context = this.context;
    if (context === null || !this.known.includes(name)) {
      return { ok: false, kind: 'missing', message: `There is no function called ${name}.` };
    }
    if (args.every(arg => !Array.isArray(arg))) {
      return this.callPlain(context, name, args as readonly FunctionScalar[]);
    }
    const call = context.getProp(context.global, '__call');
    const nameHandle = context.newString(name);
    const argsHandle = context.newString(JSON.stringify(args));
    this.deadline = this.now() + this.limits.milliseconds;
    const result = context.callFunction(call, context.undefined, nameHandle, argsHandle);
    const late = this.now() > this.deadline;
    this.deadline = Number.POSITIVE_INFINITY;
    call.dispose();
    nameHandle.dispose();
    argsHandle.dispose();
    if (result.error !== undefined) {
      const message = describe(context, result.error);
      result.error.dispose();
      return this.failure(name, message, late);
    }
    const text = context.getString(result.value);
    result.value.dispose();
    return resultOf(name, text);
  }

  private callPlain(context: QuickJSContext, name: string, args: readonly FunctionScalar[]): FunctionResult {
    let fn = this.handles.get(name);
    if (fn === undefined) {
      fn = context.getProp(context.global, name);
      this.handles.set(name, fn);
    }
    const made = args.map(arg =>
      arg === null
        ? context.null
        : typeof arg === 'number'
          ? context.newNumber(arg)
          : typeof arg === 'string'
            ? context.newString(arg)
            : arg
              ? context.true
              : context.false
    );
    this.deadline = this.now() + this.limits.milliseconds;
    const result = context.callFunction(fn, context.undefined, ...made);
    const late = this.now() > this.deadline;
    this.deadline = Number.POSITIVE_INFINITY;
    for (const handle of made) {
      handle.dispose();
    }
    if (result.error !== undefined) {
      const message = describe(context, result.error);
      result.error.dispose();
      return this.failure(name, message, late);
    }
    const kind = context.typeof(result.value);
    let value: FunctionResult;
    if (kind === 'number') {
      value = { ok: true, value: context.getNumber(result.value) };
    } else if (kind === 'string') {
      value = { ok: true, value: context.getString(result.value) };
    } else {
      // Anything else is rare enough to go through JSON, which is also
      // what checks it into something a cell can hold.
      const json = context.getProp(context.global, 'JSON');
      const stringify = context.getProp(json, 'stringify');
      const text = context.callFunction(stringify, json, result.value);
      json.dispose();
      stringify.dispose();
      if (text.error !== undefined) {
        text.error.dispose();
        value = { ok: false, kind: 'result', message: `${name} returned something a cell cannot hold.` };
      } else {
        const written = context.typeof(text.value) === 'string' ? context.getString(text.value) : 'null';
        text.value.dispose();
        value = resultOf(name, `{"value":${kind === 'undefined' ? 'null' : written}}`, kind === 'function');
      }
    }
    result.value.dispose();
    return value;
  }

  private failure(name: string, message: string, late: boolean): FunctionResult {
    if (late || /interrupted/i.test(message)) {
      return { ok: false, kind: 'deadline', message: `${name} ran past its ${this.limits.milliseconds}ms.` };
    }
    if (/out of memory/i.test(message)) {
      return { ok: false, kind: 'memory', message: `${name} ran out of memory.` };
    }
    return { ok: false, kind: 'thrown', message };
  }

  dispose(): void {
    this.releaseHandles();
    this.context?.dispose();
    this.context = null;
    this.runtime.dispose();
  }
}

function describe(context: QuickJSContext, error: Parameters<QuickJSContext['dump']>[0]): string {
  const dumped = context.dump(error) as { message?: unknown; name?: unknown } | unknown;
  if (typeof dumped === 'object' && dumped !== null && 'message' in dumped) {
    return String((dumped as { message: unknown }).message);
  }
  return String(dumped);
}

/** A result as JSON, checked into something a cell can hold. */
function resultOf(name: string, text: string, notData = false): FunctionResult {
  const parsed = JSON.parse(text) as { missing?: true; value?: unknown };
  if (notData) {
    return { ok: false, kind: 'result', message: `${name} returned something a cell cannot hold.` };
  }
  if (parsed.missing === true) {
    return { ok: false, kind: 'missing', message: `There is no function called ${name}.` };
  }
  const value = parsed.value;
  const scalar = (item: unknown): item is FunctionScalar =>
    item === null || typeof item === 'number' || typeof item === 'string' || typeof item === 'boolean';
  if (scalar(value)) {
    return { ok: true, value };
  }
  if (Array.isArray(value) && value.every(row => Array.isArray(row) && row.every(scalar))) {
    return { ok: true, value: value as FunctionScalar[][] };
  }
  return { ok: false, kind: 'result', message: `${name} returned something a cell cannot hold.` };
}
