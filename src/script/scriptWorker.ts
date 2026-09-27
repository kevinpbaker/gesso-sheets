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
 *    the ones handed to it**: `sheet`, `workbook` and `console`. No `fetch`, no
 *    `postMessage`, no `self`, no `importScripts`, no timers, no clock,
 *    no randomness. SES rejects `import(…)` and direct `eval` in the
 *    source outright, which is what a list of deleted globals cannot
 *    do: dynamic import is syntax, not a global, and a worker that only
 *    deleted `fetch` could still load and run code from anywhere.
 * 3. **Counts what it changes as it goes**, and stops the script at the
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
 * a Node worker thread given the two globals it needs. That is why the
 * number formats are listed here as well as in `protocol.ts`.
 */
import 'ses';

import type { ScriptGlobals, ScriptRange, ScriptSheet, ScriptWorkbook } from './api';
import type { ScriptFormat, ScriptOp, ScriptRunMessage, ScriptValue, ScriptWorkerMessage } from './protocol';

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

const NUMBER_FORMATS = ['general', 'number', 'currency', 'percent', 'date', 'text'];
const ALIGNS = ['auto', 'start', 'center', 'end'];
const COLOUR = /^#[0-9a-f]{6}$/i;

interface Rect {
  readonly firstRow: number;
  readonly firstColumn: number;
  readonly lastRow: number;
  readonly lastColumn: number;
}

function columnName(column: number): string {
  let name = '';
  for (let at = column + 1; at > 0; at = Math.floor((at - 1) / 26)) {
    name = String.fromCharCode(65 + ((at - 1) % 26)) + name;
  }
  return name;
}

function isValue(value: unknown): value is ScriptValue {
  return (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean' ||
    (typeof value === 'number' && Number.isFinite(value))
  );
}

function run(message: ScriptRunMessage): ScriptWorkerMessage {
  const ops: ScriptOp[] = [];
  const log: string[] = [];
  const { limits } = message;
  // The run's own copy of every sheet's values, so a read after a write
  // in the same run sees what was written.
  const books = message.book.sheets.map(sheet => ({ name: sheet.name, cells: new Map(Object.entries(sheet.cells)) }));
  let spent = 0;
  let added = 0;

  const spend = (cells: number): void => {
    if (spent + cells > limits.writes) {
      throw new RangeError(`A script may write or format at most ${limits.writes} cells in one run.`);
    }
    spent += cells;
  };

  const cellAt = (text: string): { row: number; column: number } => {
    const match = /^([A-Za-z]{1,3})(\d{1,7})$/.exec(text);
    if (match === null) {
      throw new TypeError(`"${text}" is not a cell address.`);
    }
    let column = 0;
    for (const letter of match[1].toUpperCase()) {
      column = column * 26 + (letter.charCodeAt(0) - 64);
    }
    const at = { row: Number(match[2]) - 1, column: column - 1 };
    if (at.row < 0 || at.row >= limits.rows || at.column < 0 || at.column >= limits.columns) {
      throw new RangeError(`${text} is outside the sheet.`);
    }
    return at;
  };

  const rectOf = (address: unknown): Rect => {
    if (typeof address !== 'string') {
      throw new TypeError('A range is written as an address, like "B4" or "A1:C9".');
    }
    const parts = address.trim().split(':');
    if (parts.length > 2) {
      throw new TypeError(`"${address}" is not a range.`);
    }
    const from = cellAt(parts[0].trim());
    const to = parts.length === 2 ? cellAt(parts[1].trim()) : from;
    return {
      firstRow: Math.min(from.row, to.row),
      firstColumn: Math.min(from.column, to.column),
      lastRow: Math.max(from.row, to.row),
      lastColumn: Math.max(from.column, to.column)
    };
  };

  const put = (sheet: number, row: number, column: number, value: unknown): void => {
    if (!isValue(value)) {
      throw new TypeError('A cell holds text, a number, true or false, or nothing.');
    }
    spend(1);
    ops.push({ kind: 'write', sheet, row, column, value });
    const key = `${row}:${column}`;
    if (value === null || value === '') {
      books[sheet].cells.delete(key);
    } else {
      books[sheet].cells.set(key, value);
    }
  };

  const formatOf = (change: unknown): ScriptFormat => {
    if (typeof change !== 'object' || change === null || Array.isArray(change)) {
      throw new TypeError('A format is an object, like { bold: true }.');
    }
    const source = change as Record<string, unknown>;
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(source)) {
      const value = source[key];
      const wrong = (what: string): never => {
        throw new TypeError(`${key} is ${what}.`);
      };
      if (key === 'bold' || key === 'italic' || key === 'underline' || key === 'wrap') {
        if (typeof value !== 'boolean') wrong('true or false');
      } else if (key === 'color' || key === 'fill') {
        if (typeof value !== 'string' || !COLOUR.test(value)) wrong('a colour written as "#rrggbb"');
      } else if (key === 'align') {
        if (!ALIGNS.includes(value as string)) wrong(`one of ${ALIGNS.join(', ')}`);
      } else if (key === 'fontSize') {
        if (typeof value !== 'number' || !(value >= 6 && value <= 72)) wrong('a size from 6 to 72');
      } else if (key === 'number') {
        if (!NUMBER_FORMATS.includes(value as string)) wrong(`one of ${NUMBER_FORMATS.join(', ')}`);
      } else {
        throw new TypeError(`"${key}" is not something a script can format.`);
      }
      out[key] = value;
    }
    return out as ScriptFormat;
  };

  const rangeOf = (sheet: number, rect: Rect): ScriptRange => {
    const rows = rect.lastRow - rect.firstRow + 1;
    const columns = rect.lastColumn - rect.firstColumn + 1;
    const address =
      rows === 1 && columns === 1
        ? `${columnName(rect.firstColumn)}${rect.firstRow + 1}`
        : `${columnName(rect.firstColumn)}${rect.firstRow + 1}:${columnName(rect.lastColumn)}${rect.lastRow + 1}`;
    return harden({
      address,
      rows,
      columns,
      get values(): ScriptValue[][] {
        const cells = books[sheet].cells;
        const out: ScriptValue[][] = [];
        for (let row = rect.firstRow; row <= rect.lastRow; row++) {
          const line: ScriptValue[] = [];
          for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
            line.push(cells.get(`${row}:${column}`) ?? null);
          }
          out.push(line);
        }
        return out;
      },
      write(given: unknown): void {
        if (!Array.isArray(given)) {
          for (let row = rect.firstRow; row <= rect.lastRow; row++) {
            for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
              put(sheet, row, column, given);
            }
          }
          return;
        }
        if (given.length > rows) {
          throw new RangeError(`${given.length} rows do not fit in ${address}, which has ${rows}.`);
        }
        given.forEach((line: unknown, index: number) => {
          if (!Array.isArray(line)) {
            throw new TypeError('Rows are written as a list of lists: [[1, 2], [3, 4]].');
          }
          if (line.length > columns) {
            throw new RangeError(`A row of ${line.length} does not fit in ${address}, which is ${columns} wide.`);
          }
          line.forEach((value: unknown, offset: number) => put(sheet, rect.firstRow + index, rect.firstColumn + offset, value));
        });
      },
      format(change: unknown): void {
        const checked = formatOf(change);
        spend(rows * columns);
        ops.push({ kind: 'format', sheet, ...rect, change: checked });
      },
      clear(): void {
        for (let row = rect.firstRow; row <= rect.lastRow; row++) {
          for (let column = rect.firstColumn; column <= rect.lastColumn; column++) {
            put(sheet, row, column, null);
          }
        }
      }
    });
  };

  const sheets: ScriptSheet[] = [];
  const sheetOf = (index: number): ScriptSheet => {
    sheets[index] ??= harden({
      name: books[index].name,
      range: (address: string) => rangeOf(index, rectOf(address)),
      read(address: unknown): ScriptValue {
        const rect = rectOf(address);
        return books[index].cells.get(`${rect.firstRow}:${rect.firstColumn}`) ?? null;
      },
      write(address: unknown, value: unknown): void {
        const rect = rectOf(address);
        put(index, rect.firstRow, rect.firstColumn, value);
      }
    });
    return sheets[index];
  };

  const workbook: ScriptWorkbook = harden({
    get sheets(): readonly ScriptSheet[] {
      return harden(books.map((_, index) => sheetOf(index)));
    },
    sheet(name: unknown): ScriptSheet {
      const wanted = typeof name === 'string' ? name.trim().toUpperCase() : null;
      const index = books.findIndex(book => book.name.toUpperCase() === wanted);
      if (index < 0) {
        throw new RangeError(`There is no sheet called "${String(name)}".`);
      }
      return sheetOf(index);
    },
    addSheet(name: unknown): ScriptSheet {
      const trimmed = typeof name === 'string' ? name.trim() : '';
      if (trimmed === '' || trimmed.length > 31 || /[[\]:*?/\\]/.test(trimmed)) {
        throw new TypeError('A sheet name is 1 to 31 characters, without [ ] : * ? / or \\.');
      }
      if (books.some(book => book.name.toUpperCase() === trimmed.toUpperCase())) {
        throw new RangeError(`There is already a sheet called "${trimmed}".`);
      }
      if (added >= limits.sheets) {
        throw new RangeError(`A script may add at most ${limits.sheets} sheets in one run.`);
      }
      added++;
      ops.push({ kind: 'addSheet', name: trimmed });
      books.push({ name: trimmed, cells: new Map() });
      return sheetOf(books.length - 1);
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
    const globals: ScriptGlobals = { sheet: sheetOf(message.book.active), workbook, console };
    const compartment = new Compartment({ globals, __options__: true });
    compartment.evaluate(message.source);
    return { type: 'done', ops, log };
  } catch (error) {
    return { type: 'failed', message: describe(error), log };
  }
}

/** What went wrong, as text. A script can throw anything, including a thing that throws when read. */
function describe(error: unknown): string {
  try {
    return error instanceof Error ? String(error.message) : String(error);
  } catch {
    return 'The script failed, with an error that could not be read.';
  }
}
