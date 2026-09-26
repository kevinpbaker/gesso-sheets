/**
 * A finger on this spreadsheet, in a real browser.
 *
 *   pnpm touch
 *   SKIP_BUILD=1 pnpm touch
 *
 * Gesso's `check:touch` drives a probe page — one scroll view — with
 * `Input.dispatchTouchEvent`, because the specs reach only the engine's
 * own events and not what lies between a contact and them: whether
 * Chrome's `pointerType: 'touch'` arrives at all, whether `touch-action`
 * lets the page keep the gesture, whether the platform's slop and the
 * recognizer's agree. Phase 26 asked for the same run against this
 * application rather than the probe, so this is that: the built sheet,
 * served, and a finger that taps, drags, holds and pulls a handle,
 * with every answer read back off the accessibility tree the engine
 * mirrors into the page.
 *
 * Its ports are its own, as every script's here are.
 */
import { spawn, type ChildProcess } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';

import { DevTools, findChrome, openPage, waitFor } from './lib/devtools.ts';

const PORT = 4176;
const DEVTOOLS_PORT = 9363;
const SIZE: readonly [number, number] = [1000, 700];

interface Point {
  readonly x: number;
  readonly y: number;
}

async function main(): Promise<void> {
  const failures: string[] = [];
  let preview: ChildProcess | undefined;
  let browser: ChildProcess | undefined;
  let devtools: DevTools | undefined;
  const profile = mkdtempSync(join(tmpdir(), 'gessosheet-touch-'));
  try {
    if (process.env.SKIP_BUILD === undefined) {
      await run('npx', ['vite', 'build']);
    }
    preview = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore', detached: true });
    const url = `http://localhost:${PORT}/`;
    await waitFor('the preview server', async () => ((await fetch(url)).ok ? true : undefined), 30_000);
    ({ browser, devtools } = await openPage(findChrome(), { url, devtoolsPort: DEVTOOLS_PORT, windowSize: SIZE, profileDir: profile }));
    const page = devtools;
    await waitFor(
      'the sheet to draw',
      async () => ((await page.evaluate<number>(`document.querySelectorAll('[role="rowheader"]').length`)) > 5 ? true : undefined),
      30_000
    );

    const touch = async (type: 'touchStart' | 'touchMove' | 'touchEnd', at: Point): Promise<void> => {
      await page.send('Input.dispatchTouchEvent', { type, touchPoints: type === 'touchEnd' ? [] : [{ x: at.x, y: at.y, id: 1 }] });
    };
    const tap = async (at: Point): Promise<void> => {
      await touch('touchStart', at);
      await touch('touchEnd', at);
      await sleep(250);
    };
    const drag = async (from: Point, to: Point, steps = 10): Promise<void> => {
      await touch('touchStart', from);
      for (let step = 1; step <= steps; step++) {
        const t = step / steps;
        await touch('touchMove', { x: from.x + (to.x - from.x) * t, y: from.y + (to.y - from.y) * t });
        await sleep(16);
      }
      await touch('touchEnd', to);
      await sleep(300);
    };
    /** The middle of a cell, from the letter above it and the number beside it. */
    const cell = async (column: string, row: string): Promise<Point> => {
      const found = await page.evaluate<Point | null>(`(() => {
        const letter = [...document.querySelectorAll('[role="columnheader"]')].find(el => el.getAttribute('aria-label') === ${JSON.stringify(column)});
        const number = [...document.querySelectorAll('[role="rowheader"]')].find(el => el.getAttribute('aria-label') === ${JSON.stringify(row)});
        if (!letter || !number) { return null; }
        const a = letter.getBoundingClientRect();
        const b = number.getBoundingClientRect();
        return { x: a.x + a.width / 2, y: b.y + b.height / 2 };
      })()`);
      if (found === null) {
        throw new Error(`${column}${row} is not on screen`);
      }
      return found;
    };
    const nameBox = (): Promise<string> =>
      page.evaluate<string>(`(() => {
        const box = [...document.querySelectorAll('[role="textbox"]')].find(el => el.getAttribute('aria-label') === 'Name box');
        return box ? (box.value ?? box.textContent ?? '').trim() : '';
      })()`);
    /**
     * The rows on screen, in order. All of them rather than the first,
     * because the seeded sheet freezes its heading rows, and the first
     * row shown is row 1 however far it has scrolled.
     */
    const rowsShown = (): Promise<string> =>
      page.evaluate<string>(
        `[...document.querySelectorAll('[role="rowheader"]')].filter(el => el.getBoundingClientRect().y >= 0).map(el => el.getAttribute('aria-label')).join(',')`
      );
    const menuOpen = (): Promise<boolean> => page.evaluate<boolean>(`document.querySelectorAll('[role="menu"]').length > 0`);
    const button = (label: string): Promise<Point | null> =>
      page.evaluate<Point | null>(`(() => {
        const el = [...document.querySelectorAll('[role="button"]')].find(el => el.getAttribute('aria-label') === ${JSON.stringify(label)});
        if (!el) { return null; }
        const r = el.getBoundingClientRect();
        return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
      })()`);
    const check = (what: string, ok: boolean, detail: string): void => {
      console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}${ok ? '' : ` — ${detail}`}`);
      if (!ok) {
        failures.push(`${what}: ${detail}`);
      }
    };

    // A tap selects.
    await tap(await cell('C', '5'));
    check('a tap selects the cell under it', (await nameBox()) === 'C5', `the name box says ${await nameBox()}`);

    // A drag scrolls, and the selection stays.
    const before = await rowsShown();
    await drag(await cell('D', '14'), await cell('D', '4'));
    const after = await rowsShown();
    check('a drag scrolls the sheet', after !== before, `the rows shown are still ${after}`);
    check('a drag does not sweep a selection', (await nameBox()) === 'C5', `the name box says ${await nameBox()}`);

    // Back to the top for what follows, with the wheel, which a finger
    // would not use but a script can.
    await page.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: 500, y: 400, deltaX: 0, deltaY: -5000 });
    await sleep(300);

    // A long press opens the menu at the finger, over the cell under it.
    const e8 = await cell('E', '8');
    await touch('touchStart', e8);
    await sleep(800);
    check('a long press opens the menu', await menuOpen(), 'no menu is open');
    check('a long press selects what it opened over', (await nameBox()) === 'E8', `the name box says ${await nameBox()}`);
    await touch('touchEnd', e8);
    await sleep(150);
    await page.press('Escape', 27);
    await sleep(150);

    // A handle drag extends the selection. From B5, below the seeded
    // sheet's merged title rows, where a tap would select the merge.
    await tap(await cell('B', '5'));
    const end = await button('Selection end');
    check('a finger is given the selection’s handles', end !== null, 'there is no handle');
    if (end !== null) {
      await drag(end, await cell('D', '8'));
      check('dragging a handle extends the selection', (await nameBox()) === 'B5:D8', `the name box says ${await nameBox()}`);
    }
  } finally {
    await devtools?.close();
    browser?.kill();
    if (preview?.pid !== undefined) {
      try {
        process.kill(-preview.pid);
      } catch {
        // Already gone.
      }
    }
    rmSync(profile, { recursive: true, force: true });
  }
  if (failures.length > 0) {
    console.log(`\nFAIL\n${failures.map(failure => `  - ${failure}`).join('\n')}`);
    process.exit(1);
  }
  console.log('\nOK — a finger taps, scrolls, holds for a menu and pulls a handle.');
}

function run(command: string, args: readonly string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: 'inherit' });
    child.on('exit', code => (code === 0 ? resolve() : reject(new Error(`${command} ${args.join(' ')} exited ${code}`))));
  });
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
