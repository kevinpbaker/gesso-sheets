import { parseSnapshot, type SheetSnapshot } from './SheetFile';
import type { SheetRepository } from './SheetRepository';

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

export class OpfsSheetRepository implements SheetRepository {
  private readonly encoder = new TextEncoder();
  private readonly decoder = new TextDecoder();
  private pending: SheetSnapshot | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private writing = false;
  private writeAgain = false;

  constructor(
    private readonly fileName = 'gessosheet.json',
    private readonly columnCount = 100
  ) {}

  async load(): Promise<SheetSnapshot | null> {
    try {
      const handle = await this.open();
      const size = handle.getSize();
      if (size === 0) {
        handle.close();
        return null;
      }
      const buffer = new Uint8Array(size);
      handle.read(buffer, { at: 0 });
      handle.close();
      return parseSnapshot(this.decoder.decode(buffer), this.columnCount);
    } catch (error) {
      // A browser without OPFS, a denied quota, a file this build
      // cannot read. The sheet still works for this session; only the
      // persistence is lost, and saying so is better than an empty
      // grid with no explanation.
      console.warn('[sheet] could not read from OPFS; starting fresh.', error);
      return null;
    }
  }

  save(snapshot: SheetSnapshot): void {
    this.pending = snapshot;
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
    } catch (error) {
      console.warn('[sheet] could not write to OPFS; this session will not persist.', error);
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
      throw new Error('This environment has no Origin Private File System.');
    }
    const root = await storage.getDirectory();
    const file = await root.getFileHandle(this.fileName, { create: true });
    return file.createSyncAccessHandle();
  }
}
