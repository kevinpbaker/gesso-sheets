/**
 * Phase 7's frame budget, in CI.
 *
 * The claim: two hundred thousand dependent cells can be recalculated
 * while somebody scrolls, and the scroll does not notice. This script
 * is that sentence with numbers attached, and it fails the build when
 * one of them regresses.
 *
 * It drives the built application in headless Chrome and scrolls it by
 * *scrolling it* — real wheel events through the shell's own listener,
 * across the barrier, into the render worker — and presses the
 * recalculate button by clicking where the button is. Phase 3 had a
 * bench that wrote scroll offsets into the tree instead, and it is
 * retired: both times hand-scrolling found something the bench could
 * not, it was because the bench went in through a side door. The
 * numbers it reads back are the same numbers the proof strip is
 * displaying to whoever is watching.
 *
 *   pnpm proof
 *   SKIP_BUILD=1 pnpm proof     # against an existing dist/
 *   PROOF_KEEP=1 pnpm proof     # leave the browser up
 */
import { spawn, type ChildProcess } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';

import { DevTools, findChrome, openPage, waitFor } from './lib/devtools.ts';

/** The number on the button, and the number in the claim. */
const CELLS = 200_000;

/** Wheel notches per scroll run. Enough that a stall cannot hide in the noise. */
const NOTCHES = 150;

/**
 * The budgets.
 *
 * `frameMs` is the render worker's own work per frame, not the gap
 * between frames: a headless browser's frame pacing is at the mercy of
 * whatever else the machine is doing, and a gap budget would fail for
 * reasons that have nothing to do with this application. What the work
 * costs is ours.
 *
 * `measured` is the whole thesis. A scroll in a sheet with known
 * geometry measures the cells that came into view and nothing else, so
 * the number is bounded by the window — about a screenful — and is
 * flatly independent of how many rows the sheet has or how many cells
 * are being recalculated behind it. A regression that made the sheet
 * ask its cells how big they are would put a five-digit number here.
 */
const BUDGET = {
  frameMs: 12,
  worstFrameMs: 60,
  measured: 1_200,
  /**
   * How much slower a frame is allowed to get when two hundred
   * thousand cells are recalculating behind it.
   *
   * The only budget here that is independent of the machine, and so
   * the only one that means the same thing on somebody else's. The
   * other three catch a regression by its size; this one catches it by
   * its shape — recalculation leaking into the frame at all.
   */
  costOfRecalculating: 4,
  /**
   * How much slower a frame is allowed to get with a menu open over
   * the sheet.
   *
   * Phase 8's addition, and the reason it is a budget rather than a
   * screenshot. The chrome is drawn by the same renderer as the grid,
   * so an open menu is an overlay of twenty nodes being laid out and
   * painted on every frame of a scroll happening underneath it. Twenty
   * nodes should cost nothing measurable; a menu whose items each
   * subscribed to the selection would cost a great deal, and it would
   * cost it silently. Stated as a *difference*, like the one above, so
   * it means the same thing on a slower machine.
   */
  costOfAnOpenMenu: 2
};

const PORT = Number(process.env.PROOF_PORT ?? '4319');
const DEVTOOLS_PORT = Number(process.env.PROOF_DEVTOOLS_PORT ?? '9319');
const SIZE: readonly [number, number] = [1280, 900];

interface ProofFrame {
  readonly at: number;
  readonly durationMs: number;
  readonly measured: number;
  readonly nodes: number;
  readonly inputLatencyMs: number | null;
}

async function main(): Promise<void> {
  const failures: string[] = [];
  let preview: ChildProcess | undefined;
  let browser: ChildProcess | undefined;
  let devtools: DevTools | undefined;
  const profile = mkdtempSync(join(tmpdir(), 'gessosheet-proof-'));

  try {
    if (process.env.SKIP_BUILD === undefined) {
      await run('npx', ['vite', 'build']);
    }
    // `detached` so that the whole group can be signalled at the end.
    // `npx` is a wrapper around the process that actually holds the
    // port, and killing the wrapper leaves the server listening — which
    // is how this repository collected a drawer of orphaned preview
    // servers and headless browsers on earlier phases.
    preview = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], {
      stdio: 'ignore',
      detached: true
    });
    const url = `http://localhost:${PORT}/`;
    await waitFor('the preview server', async () => ((await fetch(url)).ok ? true : undefined), 30_000);

    ({ browser, devtools } = await openPage(findChrome(), {
      url,
      devtoolsPort: DEVTOOLS_PORT,
      windowSize: SIZE,
      profileDir: profile
    }));

    // The recording only exists once the panel is wired and the render
    // worker has drawn, so waiting for a frame is waiting for the whole
    // application to be up.
    await waitFor(
      'the first frame',
      async () => ((await devtools.evaluate<number>('globalThis.gessosheetProof?.frames().length ?? 0')) > 0 ? true : undefined),
      30_000
    );

    // --------------------------------------------------------------
    // A scroll with nothing else going on
    // --------------------------------------------------------------
    const idle = report('scrolling an idle sheet', await scrollRun(devtools, 'scrolling an idle sheet'), failures);

    // --------------------------------------------------------------
    // The same scroll, over the top of 200,000 recalculating cells
    // --------------------------------------------------------------
    const button = await devtools.evaluate<{ x: number; y: number } | null>(
      `(() => {
         const el = document.querySelector('[aria-label^="Recalculate"]');
         if (el === null) { return null; }
         const box = el.getBoundingClientRect();
         return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
       })()`
    );
    if (button === null) {
      throw new Error('The recalculate button is not in the accessibility tree, so it cannot be clicked.');
    }
    await devtools.click(button.x, button.y);

    const what = 'scrolling while 200,000 cells recalculate';
    const busy = report(what, await scrollRun(devtools, what), failures);

    // The comparison, which is the claim itself.
    check(
      failures,
      `${what}: median frame ${busy.median.toFixed(2)}ms against ${idle.median.toFixed(2)}ms idle`,
      busy.median - idle.median <= BUDGET.costOfRecalculating,
      BUDGET.costOfRecalculating
    );

    // And the recalculation really happened, or the run above measured
    // a sheet quietly doing nothing.
    const evaluated = await waitFor(
      'the recalculation to finish',
      async () => {
        const text = await devtools.evaluate<string>(
          `document.querySelector('[aria-live]')?.textContent ?? ''`
        );
        const count = evaluatedIn(text);
        return count >= CELLS ? count : undefined;
      },
      60_000
    );
    console.log(`\n  recalculated ${evaluated.toLocaleString()} cells\n`);

    // --------------------------------------------------------------
    // The same scroll again, with a menu open on top of it
    // --------------------------------------------------------------
    //
    // Phase 8 put a menu bar, a toolbar, two fields and a status bar
    // on the same renderer as the grid, and the question that raises
    // is whether the chrome has made the sheet slower. An open menu is
    // the worst case of it: an overlay laid out and painted on every
    // frame, over a sheet that is scrolling underneath.
    //
    // The trap it is watching for is not the drawing. It is a toolbar
    // or a menu whose items bind to the selection — arrow keys move
    // the selection, so every item would rebuild its bindings on every
    // keystroke, which is Phase 0's cell-binding bug wearing a hat.
    await openTheEditMenu(devtools);
    const withMenu = report(
      'scrolling with a menu open',
      await scrollRun(devtools, 'scrolling with a menu open'),
      failures
    );
    check(
      failures,
      `scrolling with a menu open: median frame ${withMenu.median.toFixed(2)}ms against ${idle.median.toFixed(2)}ms with none`,
      withMenu.median - idle.median <= BUDGET.costOfAnOpenMenu,
      BUDGET.costOfAnOpenMenu
    );
    // Closed with a press outside it, which is how a person closes
    // one — and it has to be closed before the section below, or its
    // first click would be spent dismissing the menu instead of
    // pressing the button it was aimed at.
    await devtools.click(SIZE[0] / 2, SIZE[1] - 120);
    await sleep(120);
    if ((await devtools.evaluate<number>(`document.querySelectorAll('[role="menu"]').length`)) !== 0) {
      throw new Error('The menu would not close.');
    }

    // --------------------------------------------------------------
    // Five seconds with no main thread at all
    // --------------------------------------------------------------
    //
    // The other two runs show that the application's own work stays
    // off this thread. This one shows what that buys, by taking the
    // thread away: the page is frozen solid — it cannot answer a
    // DevTools evaluation, repaint its own strip, or forward an event
    // — and the two threads behind it carry on. The application worker
    // finishes a second two-hundred-thousand-cell recalculation during
    // the freeze, and the render worker goes on laying out and drawing.
    //
    // What it loses is the display's cadence, and only that: vsync
    // reaches a worker by way of the shell's `requestAnimationFrame`,
    // so with the shell gone the render worker falls back to its own
    // clock and its frames spread out. It is drawing the whole time.
    // That distinction is worth stating precisely, because the
    // overclaim — "the sheet keeps scrolling at sixty" — is both
    // untrue and unnecessary.
    await devtools.evaluate('globalThis.gessosheetProof.reset()');
    const blockButton = await devtools.evaluate<{ x: number; y: number }>(
      `(() => {
         const box = document.getElementById('block').getBoundingClientRect();
         return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
       })()`
    );
    await devtools.click(button.x, button.y);
    await devtools.click(blockButton.x, blockButton.y);

    // The round trip is the measurement: a reply cannot be composed
    // until the main thread is free again. Asked repeatedly rather than
    // once, because the button yields a frame before it starts — so the
    // page has a paint to show itself pressed — and a single question
    // asked into that gap is answered instantly and proves nothing.
    let frozenMs = 0;
    const deadline = Date.now() + 12_000;
    while (Date.now() < deadline && frozenMs < 4_500) {
      const askedAt = Date.now();
      await devtools.evaluate('1');
      frozenMs = Math.max(frozenMs, Date.now() - askedAt);
    }
    await sleep(500);

    const during = await devtools.evaluate<ProofFrame[]>('globalThis.gessosheetProof.frames()');
    const drawn = during.filter(frame => frame.at >= (during[0]?.at ?? 0));
    let worstGap = 0;
    for (let index = 1; index < during.length; index++) {
      worstGap = Math.max(worstGap, during[index].at - during[index - 1].at);
    }
    const settled = await devtools.evaluate<string>(`document.querySelector('[aria-live]')?.textContent ?? ''`);

    console.log(
      `  blocking the main thread for five seconds…\n` +
        `    frozen for ${frozenMs}ms · ${drawn.length} frames drawn during it · ` +
        `worst gap ${worstGap.toFixed(0)}ms · status "${settled}"`
    );
    check(failures, `the block only froze the page for ${frozenMs}ms`, frozenMs >= 4_500, 4_500);
    check(
      failures,
      `only ${drawn.length} frames were drawn while the main thread was blocked`,
      drawn.length >= 10,
      10
    );
    check(
      failures,
      `the application thread did not finish its recalculation during the freeze (status "${settled}")`,
      evaluatedIn(settled) >= 2 * CELLS,
      2 * CELLS
    );

  } finally {
    if (process.env.PROOF_KEEP === undefined) {
      devtools?.close();
      // Chrome is given a moment to go before its profile is taken
      // away, or the removal races the last of its own writes and the
      // script fails for a reason that has nothing to do with frames.
      if (browser !== undefined) {
        const ended = new Promise<void>(resolve => browser?.once('exit', () => resolve()));
        endGroup(browser);
        await Promise.race([ended, sleep(3_000)]);
      }
      endGroup(preview);
      try {
        rmSync(profile, { recursive: true, force: true });
      } catch {
        // A leftover profile in the temp directory is not a failure.
      }
    }
  }

  if (failures.length > 0) {
    console.error(`\nFAIL\n${failures.map(line => `  - ${line}`).join('\n')}\n`);
    process.exitCode = 1;
    return;
  }
  console.log('OK — the scroll held its budget with 200,000 cells recalculating behind it.\n');
}

/**
 * Scrolls by `NOTCHES` wheel notches and returns the frames drawn.
 *
 * Paced at roughly a frame apart, because the thing being measured is
 * a frame's worth of work in response to a person's worth of input; a
 * burst of a hundred wheel events in one tick measures coalescing.
 */
/**
 * Opens the Edit menu, by clicking where the word "Edit" is.
 *
 * Through the accessibility tree and a real click, like the
 * recalculate button above, and for the same reason: a menu opened by
 * poking the application's own state would prove that the state can
 * be poked.
 */
async function openTheEditMenu(devtools: DevTools): Promise<void> {
  const title = await devtools.evaluate<{ x: number; y: number } | null>(
    `(() => {
       const bar = document.querySelector('[role="menubar"]');
       if (bar === null) { return null; }
       const box = bar.getBoundingClientRect();
       // The first menu title, a few pixels in from the bar's edge.
       return { x: box.x + 24, y: box.y + box.height / 2 };
     })()`
  );
  if (title === null) {
    throw new Error('The menu bar is not in the accessibility tree, so it cannot be opened.');
  }
  await devtools.click(title.x, title.y);
  await sleep(120);
  const open = await devtools.evaluate<number>(`document.querySelectorAll('[role="menu"]').length`);
  if (open === 0) {
    throw new Error('Clicking the menu bar did not open a menu.');
  }
}

async function scrollRun(devtools: DevTools, what: string): Promise<ProofFrame[]> {
  console.log(`  ${what}…`);
  await devtools.evaluate('globalThis.gessosheetProof.reset()');
  for (let notch = 0; notch < NOTCHES; notch++) {
    await devtools.wheel(SIZE[0] / 2, SIZE[1] / 2, notch % 20 === 19 ? 120 : 0, 100);
    await sleep(16);
  }
  await sleep(200);
  return devtools.evaluate<ProofFrame[]>('globalThis.gessosheetProof.frames()');
}

/** What the status line says the application thread has ever evaluated. */
function evaluatedIn(status: string): number {
  const match = /([\d,]+) evaluated/.exec(status);
  return match === null ? 0 : Number(match[1].replace(/,/g, ''));
}

interface Stats {
  readonly median: number;
  readonly worst: number;
  readonly measured: number;
}

function report(what: string, frames: readonly ProofFrame[], failures: string[]): Stats {
  if (frames.length < NOTCHES / 4) {
    failures.push(`${what}: only ${frames.length} frames for ${NOTCHES} wheel notches — the scroll did not happen`);
    return { median: 0, worst: 0, measured: 0 };
  }
  const durations = frames.map(frame => frame.durationMs).sort((a, b) => a - b);
  const median = durations[Math.floor(durations.length / 2)];
  const worst = durations.at(-1) ?? 0;
  const measured = Math.max(...frames.map(frame => frame.measured));
  const latencies = frames.map(frame => frame.inputLatencyMs).filter((ms): ms is number => ms !== null);
  const latency = latencies.length === 0 ? null : latencies.sort((a, b) => a - b).at(Math.floor(latencies.length * 0.95));

  console.log(
    `    ${frames.length} frames · median ${median.toFixed(2)}ms · worst ${worst.toFixed(2)}ms · ` +
      `most re-measured ${measured} · p95 input latency ${latency === null ? 'n/a' : `${latency.toFixed(1)}ms`}`
  );

  check(failures, `${what}: median frame ${median.toFixed(2)}ms`, median <= BUDGET.frameMs, BUDGET.frameMs);
  check(failures, `${what}: worst frame ${worst.toFixed(2)}ms`, worst <= BUDGET.worstFrameMs, BUDGET.worstFrameMs);
  check(failures, `${what}: re-measured ${measured} nodes`, measured <= BUDGET.measured, BUDGET.measured);
  return { median, worst, measured };
}

function check(failures: string[], described: string, ok: boolean, budget: number): void {
  if (!ok) {
    failures.push(`${described}, over the budget of ${budget}`);
  }
}

/**
 * Ends a child and everything it started.
 *
 * A bare `kill` reaches the process this script spawned and nothing
 * below it, and both of the things spawned here — `npx` and a browser
 * — are parents of the process that actually holds the resource.
 */
function endGroup(child: ChildProcess | undefined): void {
  if (child?.pid === undefined) {
    return;
  }
  try {
    process.kill(-child.pid, 'SIGTERM');
  } catch {
    // Not a group leader, or already gone.
    child.kill();
  }
}

function run(command: string, args: readonly string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, [...args], { stdio: 'inherit' });
    child.on('exit', code => (code === 0 ? resolve() : reject(new Error(`${command} exited with ${String(code)}`))));
  });
}

await main();
