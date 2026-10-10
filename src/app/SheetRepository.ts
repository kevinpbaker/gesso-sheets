import type { SheetSnapshot } from './SheetFile';

/**
 * Where a sheet is kept.
 *
 * The seam the notes example calls a repository, and for the same
 * reason: the layers above must not know whether there is a disk
 * behind them, so their specs run in node and the one implementation
 * that needs a browser is the only thing that needs one.
 */
export interface SheetRepository {
  /**
   * What was stored, or null when nothing has been written yet.
   *
   * Rejects when something is there and cannot be read. Null is taken
   * as leave to start a fresh workbook and save it, so an answer of
   * null for a read that failed would save an empty sheet over the
   * real one; see `OpfsSheetRepository.load`.
   */
  load(): Promise<SheetSnapshot | null>;
  /**
   * Persists a snapshot shortly after being asked.
   *
   * Returns nothing and never throws: a keystroke must not wait for a
   * disk, and a disk that is full is not a reason to stop editing.
   */
  save(snapshot: SheetSnapshot): void;
  /** Writes anything outstanding now, for a worker shutting down. */
  flush(): Promise<void>;
  /**
   * Tells `listener` where the saves have got to, now and whenever it
   * changes, until the function returned is called.
   *
   * The one thing `save` cannot say for itself, since it returns
   * nothing: whether the copy in this browser has caught up with the
   * screen. The title bar says *Saving…* and *Saved* from this, and
   * says so when a write failed, which before this phase only the
   * console was told.
   */
  watch(listener: (state: SaveState) => void): () => void;
}

/**
 * Where the browser's copy of a workbook has got to.
 *
 * `saving` from the moment an edit is asked to be kept until it is on
 * the disk, `saved` after, `failed` when the disk refused it, and
 * `memory` when there is no disk at all — a browser without the Origin
 * Private File System — so nothing will be there after a reload.
 */
export type SaveState = 'saving' | 'saved' | 'failed' | 'memory';

/** The listeners of one repository, and the last state they were told. */
export class SaveStates {
  private readonly listeners = new Set<(state: SaveState) => void>();

  constructor(private current: SaveState = 'saved') {}

  get state(): SaveState {
    return this.current;
  }

  set(state: SaveState): void {
    if (state === this.current) {
      return;
    }
    this.current = state;
    for (const listener of [...this.listeners]) {
      listener(state);
    }
  }

  watch(listener: (state: SaveState) => void): () => void {
    this.listeners.add(listener);
    listener(this.current);
    return () => this.listeners.delete(listener);
  }
}

/** A repository that forgets, for specs and for a browser without OPFS. */
export class InMemorySheetRepository implements SheetRepository {
  private stored: SheetSnapshot | null;
  /** Saves asked for, which is what a spec about debouncing counts. */
  saves = 0;
  /** Set by a spec to have every save refused, as a full disk would. */
  failing = false;
  private readonly states = new SaveStates();

  constructor(initial: SheetSnapshot | null = null) {
    this.stored = initial;
  }

  load(): Promise<SheetSnapshot | null> {
    return Promise.resolve(this.stored);
  }

  save(snapshot: SheetSnapshot): void {
    this.saves++;
    if (this.failing) {
      this.states.set('failed');
      return;
    }
    this.stored = snapshot;
    // Through saving, as a disk goes, so a listener hears each save land.
    this.states.set('saving');
    this.states.set('saved');
  }

  flush(): Promise<void> {
    return Promise.resolve();
  }

  watch(listener: (state: SaveState) => void): () => void {
    return this.states.watch(listener);
  }

  /** What is stored now, for a spec that wants to look. */
  peek(): SheetSnapshot | null {
    return this.stored;
  }
}
