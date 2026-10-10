import { parseSnapshot, type SheetSnapshot } from './SheetFile';
import { SaveStates, type SaveState, type SheetRepository } from './SheetRepository';

/**
 * A sheet kept in the Origin Private File System.
 *
 * The same three members `InMemorySheetRepository` has, so nothing
 * above it changes and every layer's specs still run in node. That is
 * what the seam was for.
 *
 * It runs in the application worker, which is not incidental. The fast
 * OPFS path — `createSyncAccessHandle` — exists only in a dedicated
 * worker, and the write below is a *synchronous* file write that would
 * block whatever thread it happened on. On the shell it would stall a
 * keystroke; here it competes with nothing anybody can see, and the
 * render worker goes on drawing at sixty frames a second while it
 * happens. The thread split predicted this in Phase 0; this is the
 * first line that depends on it.
 */

/**
 * The slice of the OPFS API used here.
 *
 * Declared rather than imported: TypeScript's DOM library does not
 * describe `createSyncAccessHandle`, and the worker-only half of the
 * File System API is not in the lib this project compiles against.
 */
interface SyncAccessHandle {
  getSize(): number;
  read(buffer: ArrayBufferView, options?: { at?: number }): number;
  write(buffer: ArrayBufferView, options?: { at?: number }): number;
  truncate(size: number): void;
  flush(): void;
  close(): void;
}

interface SyncFileHandle {
  createSyncAccessHandle(): Promise<SyncAccessHandle>;
}

interface OpfsDirectory {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<SyncFileHandle>;
}

/**
 * How long writes are gathered before one reaches the disk.
 *
 * Long enough that typing a number is one write rather than one per
 * character, short enough that a person who types and immediately
 * closes the tab keeps what they typed. The notes example settled on
 * the same number for the same reasons.
 */
const SAVE_DEBOUNCE_MS = 300;

/**
 * How many times a read waits for a file another tab has open, at 40,
 * 80, 120… milliseconds: about two seconds in all, which is far longer
 * than any write holds it and short enough that a genuinely stuck file
 * is reported rather than waited on for ever.
 */
const READ_ATTEMPTS = 9;

class NoOpfs extends Error {}

/** A file that is there and could not be read — which must not be taken as a file that is not. */
export class SheetReadError extends Error {}

/** The error a sync access handle is refused with while another context holds one. */
function isHeldElsewhere(error: unknown): boolean {
  return typeof error === 'object' && error !== null && 'name' in error && error.name === 'NoModificationAllowedError';
}

function pause(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export class OpfsSheetRepository implements SheetRepository {
  private readonly encoder = new TextEncoder();
  private readonly decoder = new TextDecoder();
  private pending: SheetSnapshot | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private writing = false;
  private writeAgain = false;
  private readonly states = new SaveStates();

  constructor(
    private readonly fileName = 'gessosheet.json',
    private readonly columnCount = 100
  ) {}

  /**
   * What is stored: a snapshot, or null when **nothing** is.
   *
   * Null means empty and only empty, because the caller takes it as
   * leave to start fresh — and a fresh workbook is saved, over whatever
   * was there. This used to answer null for a read that *failed* as
   * well, and a read fails whenever another tab has the file open for
   * a moment: a sync access handle is exclusive across the whole
   * origin. Two tabs opened together could each read nothing, and the
   * one that owned the document saved an empty sheet over it.
   *
   * So a file somebody else is holding is waited for, briefly, and a
   * file that is there and cannot be read — held too long, or written
   * by a build this one cannot parse — is an error the caller has to
   * deal with, rather than an empty document it will save. Only a
   * browser with no OPFS at all still answers null, since there is
   * nothing there to lose.
   */
  async load(): Promise<SheetSnapshot | null> {
    for (let attempt = 0; ; attempt++) {
      let handle: SyncAccessHandle;
      try {
        handle = await this.open();
      } catch (error) {
        if (error instanceof NoOpfs) {
          console.warn('[sheet] this browser has no OPFS; this session will not persist.');
          this.states.set('memory');
          return null;
        }
        if (isHeldElsewhere(error) && attempt < READ_ATTEMPTS) {
          await pause(40 * (attempt + 1));
          continue;
        }
        throw new SheetReadError(`${this.fileName} could not be read`, { cause: error });
      }
      let text: string;
      try {
        const size = handle.getSize();
        if (size === 0) {
          return null;
        }
        const buffer = new Uint8Array(size);
        handle.read(buffer, { at: 0 });
        text = this.decoder.decode(buffer);
      } finally {
        handle.close();
      }
      const snapshot = parseSnapshot(text, this.columnCount);
      if (snapshot === null) {
        throw new SheetReadError(`${this.fileName} holds a workbook this build cannot read`);
      }
      return snapshot;
    }
  }

  watch(listener: (state: SaveState) => void): () => void {
    return this.states.watch(listener);
  }

  save(snapshot: SheetSnapshot): void {
    this.pending = snapshot;
    if (this.states.state !== 'memory') {
      this.states.set('saving');
    }
    if (this.timer !== null) {
      clearTimeout(this.timer);
    }
    this.timer = setTimeout(() => {
      this.timer = null;
      void this.write();
    }, SAVE_DEBOUNCE_MS);
  }

  async flush(): Promise<void> {
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    await this.write();
  }

  private async write(): Promise<void> {
    const snapshot = this.pending;
    if (snapshot === null) {
      return;
    }
    if (this.writing) {
      // A sync access handle is exclusive, so two overlapping writes
      // would fight over it. The second is remembered and runs after.
      this.writeAgain = true;
      return;
    }
    this.writing = true;
    this.pending = null;
    try {
      const handle = await this.open();
      const encoded = this.encoder.encode(JSON.stringify(snapshot));
      handle.truncate(0);
      handle.write(encoded, { at: 0 });
      handle.flush();
      handle.close();
      // Saved only if nothing newer is waiting behind this write.
      if (this.pending === null && this.timer === null && !this.writeAgain) {
        this.states.set('saved');
      }
    } catch (error) {
      console.warn('[sheet] could not write to OPFS; this session will not persist.', error);
      this.states.set(error instanceof NoOpfs ? 'memory' : 'failed');
    } finally {
      this.writing = false;
      if (this.writeAgain) {
        this.writeAgain = false;
        await this.write();
      }
    }
  }

  private async open(): Promise<SyncAccessHandle> {
    const storage = (navigator as unknown as { storage?: { getDirectory?: () => Promise<OpfsDirectory> } }).storage;
    if (storage?.getDirectory === undefined) {
      throw new NoOpfs('This environment has no Origin Private File System.');
    }
    const root = await storage.getDirectory();
    const file = await root.getFileHandle(this.fileName, { create: true });
    return file.createSyncAccessHandle();
  }
}
