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
  /** What was stored, or null when nothing has been written yet. */
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
}

/** A repository that forgets, for specs and for a browser without OPFS. */
export class InMemorySheetRepository implements SheetRepository {
  private stored: SheetSnapshot | null;
  /** Saves asked for, which is what a spec about debouncing counts. */
  saves = 0;

  constructor(initial: SheetSnapshot | null = null) {
    this.stored = initial;
  }

  load(): Promise<SheetSnapshot | null> {
    return Promise.resolve(this.stored);
  }

  save(snapshot: SheetSnapshot): void {
    this.saves++;
    this.stored = snapshot;
  }

  flush(): Promise<void> {
    return Promise.resolve();
  }

  /** What is stored now, for a spec that wants to look. */
  peek(): SheetSnapshot | null {
    return this.stored;
  }
}
