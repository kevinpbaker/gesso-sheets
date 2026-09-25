import { afterEach, describe, expect, it, vi } from 'vitest';

import { createComponent } from 'gesso-framework';
import { nodesUnder, renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import type { UiNode } from 'gesso-core';

import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * The toolbar says what it means, to a pointer as well as to a reader.
 *
 * A row of pictures is only legible to somebody who already knows the
 * pictures, and the icons arrived in this phase: a glyph that means
 * *align centre* to one person is three grey lines to the next. The
 * tooltip is the way out, and it is also the only place in the
 * application where somebody with a mouse finds out that the thing
 * they are clicking has a key.
 *
 * So both halves are checked here: that every button has one, and that
 * what it says is the command's name and the command's shortcut rather
 * than a second, drifting copy of them.
 *
 * The pointer is the whole point of this file, which is why it is not
 * in `TopBar.spec.tsx` — that one is keyboard-only on purpose.
 */
interface Harness {
  ui: Rendered;
  served: ServedForTest;
}

/** Long enough that the tooltip's own delay has passed. */
const AFTER_THE_PAUSE = 600;

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 420 });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served };
}

describe('the toolbar tells you what its icons are', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
    vi.useRealTimers();
  });

  /**
   * Only the pause, and only `setTimeout`.
   *
   * The tooltip opens on a timer, and thirteen buttons waited out in
   * real time is eight seconds of suite for a pause nobody is
   * watching. So the pause is skipped rather than served.
   *
   * `shouldAdvanceTime` is not optional here, and the reason is worth
   * knowing before the next spec reaches for fake timers: `settle()`
   * yields to the macrotask queue with `setTimeout(…, 0)`. A frozen
   * clock never fires it, so every `await settle()` hangs until the
   * test times out — which is exactly what this file did first time.
   * Advancing with real time keeps those alive while
   * `advanceTimersByTime` jumps the four hundred milliseconds the
   * tooltip is waiting for.
   */
  function fakeThePause(): void {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'], shouldAdvanceTime: true });
  }

  /** Rests the pointer on a node, as somebody wondering what it is would. */
  async function rest(node: UiNode): Promise<void> {
    const box = h.ui.getVisibleBox(node);
    h.ui.fireEvent.pointerMove(box.x + box.width / 2, box.y + box.height / 2);
    await h.ui.settle();
    vi.advanceTimersByTime(AFTER_THE_PAUSE);
    await h.ui.settle();
  }

  /** Takes the pointer somewhere else, so the next hover is a fresh one. */
  async function leave(): Promise<void> {
    h.ui.fireEvent.pointerMove(5, 400);
    await h.ui.settle();
  }

  function toolbarButtons(): UiNode[] {
    return nodesUnder(h.ui.getByRole('toolbar')).filter(
      node => h.ui.querySemantics(node)?.role === 'button'
    );
  }

  it('names the command and the key that runs it', async () => {
    h = await mount();
    fakeThePause();

    await rest(h.ui.getByRole('button', { name: 'Bold' }));

    expect(h.ui.queryByText('Bold (Ctrl+B)')).not.toBeNull();
  });

  it('says nothing until the pointer has rested', async () => {
    h = await mount();
    fakeThePause();

    const box = h.ui.getVisibleBox(h.ui.getByRole('button', { name: 'Bold' }));
    h.ui.fireEvent.pointerMove(box.x + box.width / 2, box.y + box.height / 2);
    await h.ui.settle();

    // A tooltip that appeared on contact would follow a pointer
    // crossing the row and flash thirteen times on the way past.
    expect(h.ui.queryByText('Bold (Ctrl+B)')).toBeNull();
  });

  it('takes it back when the pointer leaves', async () => {
    h = await mount();
    fakeThePause();

    await rest(h.ui.getByRole('button', { name: 'Bold' }));
    await leave();

    expect(h.ui.queryByText('Bold (Ctrl+B)')).toBeNull();
  });

  /**
   * The one that would have caught an icon added without a word to go
   * with it — which is the failure this whole file exists for.
   */
  it('gives every button one, starting with the name the button already has', async () => {
    h = await mount();
    fakeThePause();

    const silent: string[] = [];
    for (const button of toolbarButtons()) {
      const name = h.ui.getSemantics(button).label ?? '';
      await rest(button);
      const tip = h.ui.queryByText(new RegExp(`^${escaped(name)}( \\(.+\\))?$`));
      if (tip === null) {
        silent.push(name);
      }
      await leave();
    }

    expect(silent).toEqual([]);
  });

  it('has the whole toolbar to check', async () => {
    h = await mount();
    // A guard against the check above passing because it found nothing.
    expect(toolbarButtons().length).toBeGreaterThan(10);
  });
});

function escaped(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
