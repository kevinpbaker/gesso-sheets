import { OpfsSheetRepository } from './OpfsSheetRepository';
import { FIRST_DOCUMENT, type DocumentEntry, type SheetLibrary } from './SheetLibrary';

/**
 * The library in the Origin Private File System: an index of the
 * documents, and a file of its own for each.
 *
 * Each document's contents go through an `OpfsSheetRepository`, so the
 * debounced, synchronous, worker-only write that Phase 6 settled on is
 * what every document gets. The first document's file is the one every
 * build before this phase wrote, `gessosheet.json`, so opening this
 * build over an old profile finds the workbook where it was.
 *
 * The index is a few hundred bytes and is written when a document is
 * opened or saved to a file, which is rarely; it goes straight to disk
 * rather than through the debounce, because an entry that has not
 * reached the index is a document a new tab cannot find.
 */

interface SyncAccessHandle {
  getSize(): number;
  read(buffer: ArrayBufferView, options?: { at?: number }): number;
  write(buffer: ArrayBufferView, options?: { at?: number }): number;
  truncate(size: number): void;
  flush(): void;
  close(): void;
}

interface OpfsDirectory {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<{ createSyncAccessHandle(): Promise<SyncAccessHandle> }>;
}

const INDEX = 'documents.json';

export class OpfsSheetLibrary implements SheetLibrary {
  private readonly repositories = new Map<string, OpfsSheetRepository>();
  /** Writes to the index, one after another: a sync handle is exclusive. */
  private writing: Promise<void> = Promise.resolve();

  constructor(private readonly columnCount: number) {}

  async entries(): Promise<DocumentEntry[]> {
    try {
      const text = await this.readIndex();
      const parsed: unknown = text === '' ? [] : JSON.parse(text);
      return Array.isArray(parsed) ? parsed.filter(isEntry) : [];
    } catch (error) {
      console.warn('[sheet] could not read the document index; starting a fresh one.', error);
      return [];
    }
  }

  put(entry: DocumentEntry): Promise<void> {
    this.writing = this.writing.then(async () => {
      try {
        const others = (await this.entries()).filter(each => each.id !== entry.id);
        await this.writeIndex(JSON.stringify([...others, entry]));
      } catch (error) {
        console.warn('[sheet] could not write the document index.', error);
      }
    });
    return this.writing;
  }

  repository(id: string): OpfsSheetRepository {
    let repository = this.repositories.get(id);
    if (repository === undefined) {
      repository = new OpfsSheetRepository(id === FIRST_DOCUMENT ? 'gessosheet.json' : `document-${id}.json`, this.columnCount);
      this.repositories.set(id, repository);
    }
    return repository;
  }

  newId(): string {
    // Time first so that ids sort by age, which is what a person
    // listing the directory would expect; the random tail is what
    // keeps two tabs created in the same millisecond apart.
    return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  }

  private async readIndex(): Promise<string> {
    const handle = await this.open();
    try {
      const size = handle.getSize();
      const buffer = new Uint8Array(size);
      handle.read(buffer, { at: 0 });
      return new TextDecoder().decode(buffer);
    } finally {
      handle.close();
    }
  }

  private async writeIndex(text: string): Promise<void> {
    const handle = await this.open();
    try {
      const bytes = new TextEncoder().encode(text);
      handle.truncate(0);
      handle.write(bytes, { at: 0 });
      handle.flush();
    } finally {
      handle.close();
    }
  }

  private async open(): Promise<SyncAccessHandle> {
    const storage = (navigator as unknown as { storage?: { getDirectory?: () => Promise<OpfsDirectory> } }).storage;
    if (storage?.getDirectory === undefined) {
      throw new Error('This environment has no Origin Private File System.');
    }
    const root = await storage.getDirectory();
    return (await root.getFileHandle(INDEX, { create: true })).createSyncAccessHandle();
  }
}

function isEntry(value: unknown): value is DocumentEntry {
  if (typeof value !== 'object' || value === null) {
    return false;
  }
  const entry = value as Record<string, unknown>;
  return typeof entry.id === 'string' && typeof entry.name === 'string' && typeof entry.used === 'number';
}
