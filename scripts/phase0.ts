/**
 * The Phase 0 measurement, unattended.
 *
 * Builds the spike, serves it, opens it in headless Chrome with
 * `?bench`, and prints the table of runs the render worker logs.
 *
 * Two things are worth knowing about what it reports.
 *
 *   - **Cost, not frame rate.** Headless Chrome schedules frames
 *     however it likes, so the gap between them says more about the
 *     compositor than about the application. `cost` is what the render
 *     worker spent building, laying out and painting the frame, and
 *     that is the number 16.7 ms has to be compared against.
 *   - **The console comes from a worker.** The render worker's console
 *     is its own DevTools target, so this attaches to every target the
 *     page spawns rather than reading the page's log.
 *
 * It needs Chrome on the path, or `CHROME_BIN`.
 */
import { execFileSync, spawn, type ChildProcess } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';

const PORT = 4173;
const DEVTOOLS_PORT = 9333;
const CHROME_CANDIDATES = ['google-chrome-stable', 'google-chrome', 'chromium', 'chromium-browser', 'chrome'];
const BENCH_PREFIX = 'PHASE0 ';
const BENCH_DONE = 'PHASE0-DONE';

function findChrome(): string {
  for (const candidate of process.env.CHROME_BIN ? [process.env.CHROME_BIN] : CHROME_CANDIDATES) {
    try {
      execFileSync(candidate, ['--version'], { stdio: 'ignore' });
      return candidate;
    } catch {
      /* next */
    }
  }
  throw new Error(`No Chrome found. Tried ${CHROME_CANDIDATES.join(', ')}. Set CHROME_BIN.`);
}

/**
 * A DevTools client flat enough to reach a worker.
 *
 * `flatten: true` puts every attached target's traffic on this one
 * socket tagged with a session id, which is the only reason this is
 * sixty lines instead of a socket per target.
 */
class Client {
  private nextId = 1;
  private readonly pending = new Map<number, (value: unknown) => void>();
  readonly lines: string[] = [];
  private doneResolve: (() => void) | null = null;

  private readonly socket: WebSocket;

  constructor(socket: WebSocket) {
    this.socket = socket;
    socket.addEventListener('message', event => {
      const message = JSON.parse(String(event.data)) as {
        id?: number;
        result?: unknown;
        error?: { message: string };
        method?: string;
        sessionId?: string;
        params?: Record<string, unknown>;
      };
      if (message.id !== undefined) {
        const resolve = this.pending.get(message.id);
        this.pending.delete(message.id);
        resolve?.(message.error !== undefined ? new Error(message.error.message) : message.result);
        return;
      }
      if (message.method === 'Target.attachedToTarget') {
        const session = (message.params as { sessionId: string }).sessionId;
        void this.send('Runtime.enable', {}, session);
        void this.send('Target.setAutoAttach', AUTO_ATTACH, session);
        return;
      }
      if (message.method === 'Runtime.consoleAPICalled') {
        const args = (message.params as { args: { value?: unknown }[] }).args;
        const text = args.map(arg => String(arg.value ?? '')).join(' ');
        if (text.startsWith(BENCH_PREFIX)) {
          this.lines.push(text.slice(BENCH_PREFIX.length));
          process.stderr.write('.');
        } else if (text.startsWith(BENCH_DONE)) {
          this.doneResolve?.();
        }
      }
    });
  }

  static async connect(url: string): Promise<Client> {
    const socket = new WebSocket(url);
    await new Promise<void>((resolve, reject) => {
      socket.addEventListener('open', () => resolve(), { once: true });
      socket.addEventListener('error', () => reject(new Error(`Could not connect to ${url}`)), { once: true });
    });
    return new Client(socket);
  }

  send(method: string, params: Record<string, unknown> = {}, sessionId?: string): Promise<unknown> {
    const id = this.nextId++;
    return new Promise(resolve => {
      this.pending.set(id, resolve);
      this.socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });
  }

  finished(timeoutMs: number): Promise<void> {
    return new Promise((resolve, reject) => {
      this.doneResolve = resolve;
      setTimeout(() => reject(new Error(`The bench did not finish within ${timeoutMs / 1000}s.`)), timeoutMs);
    });
  }

  close(): void {
    this.socket.close();
  }
}

const AUTO_ATTACH = { autoAttach: true, waitForDebuggerOnStart: false, flatten: true };

async function waitForTarget(): Promise<string> {
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline) {
    try {
      const targets = (await (await fetch(`http://localhost:${DEVTOOLS_PORT}/json`)).json()) as {
        type: string;
        url: string;
        webSocketDebuggerUrl: string;
      }[];
      const page = targets.find(t => t.type === 'page' && t.url.includes(String(PORT)));
      if (page !== undefined) {
        return page.webSocketDebuggerUrl;
      }
    } catch {
      /* not up yet */
    }
    await sleep(250);
  }
  throw new Error('Chrome never opened the page.');
}

async function main(): Promise<void> {
  const chrome = findChrome();
  console.error(`building…`);
  execFileSync('npx', ['vite', 'build'], { stdio: 'inherit' });

  const preview = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore' });
  const profile = mkdtempSync(join(tmpdir(), 'gessosheet-phase0-'));
  let browser: ChildProcess | undefined;
  let client: Client | undefined;
  try {
    await sleep(1500);
    console.error('running the matrix (one dot per run)…');
    browser = spawn(
      chrome,
      [
        '--headless=new',
        '--no-first-run',
        '--no-default-browser-check',
        '--hide-scrollbars',
        '--force-device-scale-factor=1',
        '--window-size=1400,900',
        `--user-data-dir=${profile}`,
        `--remote-debugging-port=${DEVTOOLS_PORT}`,
        `http://localhost:${PORT}/?bench`
      ],
      { stdio: 'ignore' }
    );
    client = await Client.connect(await waitForTarget());
    await client.send('Runtime.enable');
    await client.send('Target.setAutoAttach', AUTO_ATTACH);
    const finished = client.finished(360_000);
    await finished;
    process.stderr.write('\n');
    report(client.lines.map(line => JSON.parse(line) as Record<string, unknown>));
  } finally {
    client?.close();
    browser?.kill();
    preview.kill();
    rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
  }
}

/** The table, in the order the runs were declared. */
function report(runs: Record<string, unknown>[]): void {
  const columns = [
    'run',
    'axis',
    'speed',
    'mountBand',
    'fetchBand',
    'shape',
    'busy',
    'costMs',
    'costP95Ms',
    'gapMs',
    'missFramesPct',
    'worstMissPct',
    'peakNodes',
    'peakMeasured',
    'cells',
    'patchesPerPublish',
    'bytesPerPublish'
  ];
  const phases = ['layout', 'render', 'patches', 'virtualize', 'semantics'];
  const all = [...columns, ...phases];
  const rows = runs.map(run =>
    all.map(column =>
      phases.includes(column) ? format((run.phases as Record<string, unknown>)?.[column]) : format(run[column])
    )
  );
  const widths = all.map((column, index) => Math.max(column.length, ...rows.map(row => row[index]?.length ?? 0)));
  const line = (cells: string[]) => cells.map((cell, index) => cell.padStart(widths[index])).join('  ');
  console.log(line(all));
  console.log(widths.map(width => '─'.repeat(width)).join('  '));
  for (const row of rows) {
    console.log(line(row));
  }
}

function format(value: unknown): string {
  if (Array.isArray(value)) {
    return value.join('/');
  }
  return String(value ?? '');
}

await main();
