import { applyPatches, provide, type Patch } from 'gesso-framework';

import { Sheet, type SheetView, type SheetWindow } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService, type Schedule } from './SheetService';

/**
 * What a spec needs to drive the application thread for real.
 *
 * Shared by the specs that make claims about the barrier, because a
 * claim about the barrier is only worth making against the real
 * `provide`, the real differ and the real service. Anything a spec
 * built for itself here would be a claim about the spec.
 */

/** A `ChannelPort` that keeps what was posted. */
export class RecordingPort {
  onmessage: ((event: { data: unknown }) => void) | null = null;
  /** Since the last `clear`, which is what a budget is measured over. */
  readonly posted: unknown[] = [];
  /** Everything, ever, which is what replaying the view needs. */
  private readonly history: unknown[] = [];

  postMessage(message: unknown): void {
    this.posted.push(message);
    this.history.push(message);
  }

  /** The render thread's end, sending a message in. */
  send(message: unknown): void {
    this.onmessage?.({ data: message });
  }

  patches(): Patch[] {
    return this.collect(this.posted);
  }

  /** Every patch for a key since the channel opened. */
  allPatchesFor(projection: keyof SheetView): Patch[] {
    return this.collect(this.history).filter(patch => patch.projection === projection);
  }

  private collect(messages: readonly unknown[]): Patch[] {
    const all: Patch[] = [];
    for (const message of messages) {
      const typed = message as { type?: string; patches?: Patch[] };
      if (typed.type === 'channel:patch') {
        all.push(...(typed.patches ?? []));
      }
      if (typed.type === 'channel:error') {
        throw new Error(`the channel reported: ${String((message as { message: string }).message)}`);
      }
    }
    return all;
  }

  /** Patches for one view key, which is how a budget is stated. */
  patchesFor(projection: keyof SheetView): Patch[] {
    return this.patches().filter(patch => patch.projection === projection);
  }

  clear(): void {
    this.posted.length = 0;
  }
}

/** A scheduler the spec advances by hand, so slices are deterministic. */
export class ManualSchedule {
  private queue: (() => void)[] = [];

  readonly schedule: Schedule = run => {
    this.queue.push(run);
  };

  get pending(): number {
    return this.queue.length;
  }

  /** Runs one slice. */
  tick(): boolean {
    const next = this.queue.shift();
    next?.();
    return next !== undefined;
  }

  /** Runs slices until the pump stops. */
  drain(): number {
    let ran = 0;
    while (this.tick()) {
      ran++;
      if (ran > 10_000) {
        throw new Error('the pump never finished');
      }
    }
    return ran;
  }
}

export interface Harness {
  readonly port: RecordingPort;
  readonly service: SheetService;
  readonly document: SheetDocument;
  readonly clock: ManualSchedule;
  window(): SheetWindow;
}

export function attach(options: { budget?: number; rowCount?: number; columnCount?: number } = {}): Harness {
  const document = new SheetDocument();
  const clock = new ManualSchedule();
  const service = new SheetService(document, {
    schedule: clock.schedule,
    budget: options.budget ?? 2_000,
    rowCount: options.rowCount,
    columnCount: options.columnCount
  });
  const port = new RecordingPort();
  provide(Sheet, sheetChannel(service).source as never, port);
  // The replica asks rather than waiting to be pushed to; this is that.
  port.send({ type: 'channel:sync' });
  return {
    port,
    service,
    document,
    clock,
    window: () => rebuild(port)
  };
}

/**
 * The window as the render thread would hold it.
 *
 * Replayed with the framework's own `applyPatches` rather than
 * something written for the spec: the claim being tested is that these
 * patches reconstruct the window, and a hand-rolled applier would only
 * test that they reconstruct it *the way the spec imagined*. It also
 * has to start from the token's initial value, which is where a
 * replica starts.
 */
function rebuild(port: RecordingPort): SheetWindow {
  return applyPatches(Sheet.initial.window, port.allPatchesFor('window')) as SheetWindow;
}
