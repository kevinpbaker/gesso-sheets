/**
 * The strip above the canvas, and the only DOM in the application.
 *
 * Everything here runs on the main thread on purpose. The claim the
 * whole project exists to make is that the main thread is not doing
 * the work, and a claim like that cannot be made from inside the
 * thread that is: numbers the render worker prints on its own canvas
 * are numbers a stranger has no reason to believe. So the readout
 * lives where the doubt lives. Block this thread and the strip stops
 * dead — the pulse freezes, the button stays pressed, the page cannot
 * be selected or scrolled — while the two threads behind it carry on,
 * because the application is not here.
 *
 * Three things are wired: the block, the heatmap, and the readout.
 */
import type { WorkerApp } from 'gesso-framework';
import type { FrameMetrics } from 'gesso-framework';
import type { UiNodeReport } from 'gesso-framework';

/** How long the block button holds the thread, in milliseconds. */
const BLOCK_MS = 5_000;

/** How many frames the rolling readout averages over. */
const WINDOW = 90;

/** How many frames the recording keeps for a machine to read back. */
const RECORDING = 2_000;

/** One frame, as the budget check reads it. */
export interface ProofFrame {
  readonly at: number;
  readonly durationMs: number;
  readonly measured: number;
  readonly nodes: number;
  readonly inputLatencyMs: number | null;
}

/**
 * What `scripts/frame-budget.ts` drives the page through.
 *
 * Deliberately the readout's own numbers rather than a measurement
 * the check installs for itself. Every time a budget in this project
 * went in through a side door it agreed with itself and disagreed with
 * the screen — the Phase 3 bench wrote scroll offsets instead of
 * scrolling and missed both scrolling bugs a hand found. So the check
 * sends real wheel events and real clicks, and reads the same frames
 * the strip is displaying while it does.
 */
export interface ProofHandle {
  frames(): readonly ProofFrame[];
  reset(): void;
}

declare global {
  // eslint-disable-next-line no-var
  var gessosheetProof: ProofHandle | undefined;
}

/**
 * Wires the strip's controls to an application.
 *
 * Returns the callbacks the application has to be *created* with —
 * `onFrame` and `onInspect` are constructor options, not something a
 * running app can be handed later — so the shell builds the panel
 * first and passes them in.
 */
export function proofPanel(): {
  readonly options: { onFrame: (metrics: FrameMetrics) => void; onInspect: (report: UiNodeReport | null) => void };
  readonly attach: (app: WorkerApp) => void;
} {
  const pulse = element('pulse');
  const block = element<HTMLButtonElement>('block');
  const heatmap = element<HTMLInputElement>('heatmap');
  const explain = element('explain');
  const fpsOut = element('fps');
  const gapOut = element('gap');
  const measuredOut = element('measured');
  const peakOut = element('peak');
  const mainOut = element('mainfps');

  // ---------------------------------------------------------------------
  // What the render worker did
  // ---------------------------------------------------------------------

  /**
   * Frame finish times on the render worker's clock.
   *
   * Deliberately `metrics.at` and not the time this callback ran. A
   * blocked main thread cannot receive messages, so every frame drawn
   * during the block arrives in one burst the moment it unblocks; if
   * the readout timed its own arrivals it would report a five second
   * stall that never happened, and it would report it about the wrong
   * thread. The render worker stamps each frame when it finishes, and
   * those stamps survive a queue.
   */
  const finishes: number[] = [];
  const recording: ProofFrame[] = [];
  let worstGap = 0;
  let peakMeasured = 0;

  globalThis.gessosheetProof = {
    frames: () => recording,
    reset: () => {
      recording.length = 0;
      worstGap = 0;
      peakMeasured = 0;
    }
  };

  const onFrame = (metrics: FrameMetrics): void => {
    recording.push({
      at: metrics.at,
      durationMs: metrics.durationMs,
      measured: metrics.measured,
      nodes: metrics.nodes,
      inputLatencyMs: metrics.inputLatencyMs
    });
    if (recording.length > RECORDING) {
      recording.shift();
    }
    finishes.push(metrics.at);
    if (finishes.length > WINDOW) {
      finishes.shift();
    }
    const previous = finishes.at(-2);
    if (previous !== undefined) {
      worstGap = Math.max(worstGap, metrics.at - previous);
    }
    measuredOut.textContent = String(metrics.measured);
    /**
     * The peak, which is the number that makes the heatmap readable.
     *
     * Sweep the sheet and the wash goes red everywhere, and it is
     * telling the truth: every row on screen came into view during the
     * sweep, and a row is measured once when it does. The claim was
     * never that a scroll measures nothing — it is that a scroll
     * measures a screenful and stops, whether the sheet is ten
     * thousand rows or ten million, and whether or not two hundred
     * thousand cells are being recalculated behind it. Without a peak
     * to look at, the red says the first thing and hides the second.
     */
    peakMeasured = Math.max(peakMeasured, metrics.measured);
    peakOut.textContent = String(peakMeasured);
  };

  // ---------------------------------------------------------------------
  // What this thread did
  // ---------------------------------------------------------------------

  let mainFrames = 0;
  let sampledAt = performance.now();

  /**
   * The pulse, the main thread's frames, and the readout, all on this
   * thread's animation frame — so all three stop together when it is
   * blocked, and stop in front of a sheet that has not.
   */
  const tick = (now: number): void => {
    mainFrames++;
    pulse.style.opacity = String(0.35 + 0.65 * Math.abs(Math.sin(now / 350)));

    const since = now - sampledAt;
    if (since >= 500) {
      mainOut.textContent = String(Math.round((mainFrames / since) * 1000));
      mainFrames = 0;
      sampledAt = now;

      const first = finishes[0];
      const last = finishes.at(-1);
      if (first !== undefined && last !== undefined && last > first) {
        fpsOut.textContent = String(Math.round(((finishes.length - 1) / (last - first)) * 1000));
      }
      gapOut.textContent = worstGap === 0 ? '—' : `${worstGap.toFixed(1)}ms`;
    }
    requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);

  // ---------------------------------------------------------------------
  // The controls
  // ---------------------------------------------------------------------

  const attach = (app: WorkerApp): void => {
    /**
     * Five seconds of the thing every web application is told not to
     * do, done on purpose.
     *
     * A busy loop and not a `sleep`, because the point is not that
     * time passes — it is that this thread has no turn to give
     * anybody. What survives it, and what does not, is worth being
     * exact about, because the loose version of the claim is both
     * wrong and weaker than the true one:
     *
     * The application worker does not notice. Start a recalculation
     * and press this, and the two hundred thousand cells are already
     * done when the page comes back — it worked through the freeze.
     *
     * The render worker does not stop either. It keeps laying out and
     * drawing on its own clock, which is why the frame recording has
     * entries stamped all the way through the five seconds. What it
     * loses is the display's cadence and only that: `requestAnimationFrame`
     * exists on this thread alone, so a worker gets vsync by way of a
     * shell that is currently not answering, and its frames spread out
     * to a timer's interval until the thread comes back.
     *
     * The page, meanwhile, is gone: no events are forwarded, so nothing
     * a person does during the five seconds reaches the sheet at all.
     * That is the honest shape of it — the work is elsewhere, the input
     * is not.
     *
     * The label is repainted and the frame yielded before the loop
     * starts, or the only evidence would be a button that never looked
     * pressed.
     */
    block.addEventListener('click', () => {
      block.disabled = true;
      block.textContent = `Blocking for ${BLOCK_MS / 1000}s…`;
      pulse.classList.add('blocked');
      worstGap = 0;
      requestAnimationFrame(() =>
        requestAnimationFrame(() => {
          const until = performance.now() + BLOCK_MS;
          while (performance.now() < until) {
            /* Holding the thread. That is the whole experiment. */
          }
          block.disabled = false;
          block.textContent = `Block the main thread for ${BLOCK_MS / 1000}s`;
          pulse.classList.remove('blocked');
        })
      );
    });

    /**
     * The engine's own layout inspector: every node the layout phase
     * measured this frame, washed over the scene, plus `engine.explain`
     * for whatever is under the pointer.
     *
     * This is the measurement claim made checkable. Scroll with it on
     * and the grid stays cold, because a scroll in a sheet with known
     * geometry measures nothing — the cells that come into view were
     * never measured by asking them how big they are.
     */
    heatmap.addEventListener('change', () => {
      app.setInspector(heatmap.checked);
      explain.hidden = !heatmap.checked;
      if (!heatmap.checked) {
        explain.textContent = '';
      }
    });
  };

  /**
   * `engine.explain` for the hovered node, which is the cell inspector.
   *
   * The engine already formats the answer; the panel adds only what it
   * takes to know which node the answer is about.
   */
  const onInspect = (report: UiNodeReport | null): void => {
    if (!heatmap.checked) {
      return;
    }
    if (report === null) {
      explain.textContent = 'Point at a cell.';
      return;
    }
    const { x, y, width, height } = report.box;
    const label = report.semantics?.label;
    explain.textContent =
      `${report.type}${label === undefined ? '' : ` "${label}"`} ` +
      `— ${width.toFixed(1)}x${height.toFixed(1)} at ${x.toFixed(1)},${y.toFixed(1)}\n\n` +
      report.explanation;
  };

  return { options: { onFrame, onInspect }, attach };
}

function element<T extends HTMLElement = HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (found === null) {
    throw new Error(`index.html has no #${id} element.`);
  }
  return found as T;
}
