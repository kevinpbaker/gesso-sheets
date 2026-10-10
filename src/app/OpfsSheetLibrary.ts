import { OpfsSheetRepository, SheetReadError } from './OpfsSheetRepository';
import { parseSnapshot, type SheetSnapshot } from './SheetFile';
import { FIRST_DOCUMENT, type DocumentEntry, type SheetLibrary } from './SheetLibrary';
import { VERSION_LIMIT, type SheetVersion, type VersionReason, type VersionStore } from './SheetVersions';

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
  removeEntry(name: string): Promise<void>;
}

const INDEX = 'documents.json';

export class OpfsSheetLibrary implements SheetLibrary {
  private readonly repositories = new Map<string, OpfsSheetRepository>();
  /** Writes to the index, one after another: a sync handle is exclusive. */
  private writing: Promise<void> = Promise.resolve();

  constructor(private readonly columnCount: number) {}

  /**
   * Every document, or a rejection when the index is there and cannot
   * be read.
   *
   * Not an empty list on a failure, which is what it used to answer:
   * an empty list is "this is the first document ever", and a put
   * after it writes an index holding one entry, and every other
   * document in the browser drops out of it. See
   * `OpfsSheetRepository.load` for the same mistake with a workbook.
   */
  async entries(): Promise<DocumentEntry[]> {
    const text = await this.readIndex();
    if (text === '') {
      return [];
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(text);
    } catch (error) {
      throw new SheetReadError('The document index could not be read', { cause: error });
    }
    return Array.isArray(parsed) ? parsed.filter(isEntry) : [];
  }

  /**
   * Adds or replaces an entry: read the index, change it, write it.
   *
   * Under a Web Lock, because every tab has an application worker of
   * its own and all of them share this one file. Two tabs doing the
   * read and the write at once would each write a list without the
   * other's entry in it — a document that exists and that no new tab
   * can find — or would fail outright, since a sync access handle is
   * exclusive and the second to ask is refused. The lock makes the
   * three steps one, across every tab of the origin; the chain in front
   * of it keeps this worker's own writes in order.
   */
  put(entry: DocumentEntry): Promise<void> {
    this.writing = this.writing.then(() =>
      exclusively(async () => {
        try {
          // Read first, and not written at all if the read fails: an
          // index written from a list that could not be read is an
          // index with every other document missing from it.
          const others = (await this.entries()).filter(each => each.id !== entry.id);
          await this.writeIndex(JSON.stringify([...others, entry]));
        } catch (error) {
          console.warn('[sheet] could not update the document index; it was left as it was.', error);
        }
      })
    );
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

  /**
   * Forgets a document. The entry goes first, under the index's lock,
   * so a tab listing the library never finds an entry whose contents
   * have already gone; the files after, and a file that will not go is
   * left behind rather than holding the delete up.
   */
  async remove(id: string): Promise<void> {
    this.writing = this.writing.then(() =>
      exclusively(async () => {
        const others = (await this.entries()).filter(each => each.id !== id);
        await this.writeIndex(JSON.stringify(others));
      })
    );
    await this.writing;
    this.repositories.delete(id);
    await this.versions(id).clear();
    if (id !== FIRST_DOCUMENT) {
      await removeFile(`document-${id}.json`);
    } else {
      // The first document's file is also what a profile with no index
      // is read from, so it is emptied rather than left to come back.
      await writeFile('gessosheet.json', '');
    }
  }

  versions(id: string): VersionStore {
    return new OpfsVersionStore(id, this.columnCount);
  }

  newId(): string {
    // Time first so that ids sort by age, which is what a person
    // listing the directory would expect; the random tail is what
    // keeps two tabs created in the same millisecond apart.
    return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  }

  /**
   * A Web Lock named for the document, held until it is let go —
   * across every tab of the origin, which is the point.
   *
   * Waited for, briefly, rather than asked for only if free. A reload
   * starts the new application worker while the old one is still being
   * torn down, and the old one holds the lock until it is gone; asking
   * with `ifAvailable` made a tab that was reloaded see its own
   * document as open somewhere else. A second and a half is longer
   * than a teardown and shorter than anybody waits for a tab.
   */
  claim(id: string): Promise<(() => void) | null> {
    const locks = (
      navigator as unknown as {
        locks?: {
          request(name: string, options: { signal: AbortSignal }, run: () => Promise<void>): Promise<void>;
        };
      }
    ).locks;
    if (locks === undefined) {
      return Promise.resolve(() => {});
    }
    return new Promise(answer => {
      locks
        .request(`gessosheet-document-${id}`, { signal: AbortSignal.timeout(1500) }, () => {
          // The lock lasts as long as this promise does.
          return new Promise<void>(release => answer(release));
        })
        .catch(() => answer(null));
    });
  }

  private async readIndex(): Promise<string> {
    if (!hasOpfs()) {
      return '';
    }
    const handle = await this.openWaiting();
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
    const handle = await this.openWaiting();
    try {
      const bytes = new TextEncoder().encode(text);
      handle.truncate(0);
      handle.write(bytes, { at: 0 });
      handle.flush();
    } finally {
      handle.close();
    }
  }

  /**
   * The index's handle, waited for while another tab holds it — which
   * every tab does for a moment whenever it opens or saves a document.
   * A browser with no OPFS answers an empty index, having nothing to
   * lose; one whose file stays held throws.
   */
  private async openWaiting(): Promise<SyncAccessHandle> {
    for (let attempt = 0; ; attempt++) {
      try {
        return await this.open();
      } catch (error) {
        const held = typeof error === 'object' && error !== null && 'name' in error && error.name === 'NoModificationAllowedError';
        if (held && attempt < 9) {
          await new Promise(resolve => setTimeout(resolve, 40 * (attempt + 1)));
          continue;
        }
        throw new SheetReadError('The document index could not be opened', { cause: error });
      }
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

/**
 * A document's versions in OPFS: a small list, `versions-<id>.json`,
 * and a file per version beside it, so the list can be read without
 * reading thirty workbooks.
 */
class OpfsVersionStore implements VersionStore {
  constructor(
    private readonly id: string,
    private readonly columnCount: number
  ) {}

  private get listName(): string {
    return `versions-${this.id}.json`;
  }

  private fileOf(at: number): string {
    return `version-${this.id}-${at}.json`;
  }

  async list(): Promise<SheetVersion[]> {
    if (!hasOpfs()) {
      return [];
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse((await readFile(this.listName)) || '[]');
    } catch {
      return [];
    }
    return Array.isArray(parsed) ? parsed.filter(isVersion).sort((a, b) => b.at - a.at) : [];
  }

  async add(version: SheetVersion, snapshot: SheetSnapshot): Promise<void> {
    if (!hasOpfs()) {
      return;
    }
    await exclusivelyNamed(`gessosheet-versions-${this.id}`, async () => {
      await writeFile(this.fileOf(version.at), JSON.stringify(snapshot));
      const all = [version, ...(await this.list()).filter(each => each.at !== version.at)].sort((a, b) => b.at - a.at);
      await writeFile(this.listName, JSON.stringify(all.slice(0, VERSION_LIMIT)));
      for (const old of all.slice(VERSION_LIMIT)) {
        await removeFile(this.fileOf(old.at));
      }
    });
  }

  async get(at: number): Promise<SheetSnapshot | null> {
    if (!hasOpfs()) {
      return null;
    }
    const text = await readFile(this.fileOf(at));
    return text === '' ? null : parseSnapshot(text, this.columnCount);
  }

  async clear(): Promise<void> {
    if (!hasOpfs()) {
      return;
    }
    for (const version of await this.list()) {
      await removeFile(this.fileOf(version.at));
    }
    await removeFile(this.listName);
  }
}

function isVersion(value: unknown): value is SheetVersion {
  if (typeof value !== 'object' || value === null) {
    return false;
  }
  const version = value as Record<string, unknown>;
  const reasons: readonly VersionReason[] = ['opened', 'auto', 'saved', 'restored'];
  return typeof version.at === 'number' && reasons.includes(version.reason as VersionReason);
}

async function root(): Promise<OpfsDirectory> {
  const storage = (navigator as unknown as { storage?: { getDirectory?: () => Promise<OpfsDirectory> } }).storage;
  if (storage?.getDirectory === undefined) {
    throw new Error('This environment has no Origin Private File System.');
  }
  return storage.getDirectory();
}

/** A whole file as text, or '' when it is empty or not there. */
async function readFile(name: string): Promise<string> {
  const handle = await (await (await root()).getFileHandle(name, { create: true })).createSyncAccessHandle();
  try {
    const buffer = new Uint8Array(handle.getSize());
    handle.read(buffer, { at: 0 });
    return new TextDecoder().decode(buffer);
  } finally {
    handle.close();
  }
}

async function writeFile(name: string, text: string): Promise<void> {
  const handle = await (await (await root()).getFileHandle(name, { create: true })).createSyncAccessHandle();
  try {
    handle.truncate(0);
    handle.write(new TextEncoder().encode(text), { at: 0 });
    handle.flush();
  } finally {
    handle.close();
  }
}

async function removeFile(name: string): Promise<void> {
  try {
    await (await root()).removeEntry(name);
  } catch {
    // Not there, or held: nothing to do either way.
  }
}

function exclusivelyNamed(name: string, run: () => Promise<void>): Promise<void> {
  const locks = (navigator as unknown as { locks?: { request(name: string, run: () => Promise<void>): Promise<void> } })
    .locks;
  return locks === undefined ? run() : locks.request(name, run);
}

/**
 * Runs under the index's lock, where there is a lock manager to ask.
 *
 * Every browser with OPFS has one; a worker that somehow lacks it runs
 * the write anyway, which is what it did before there was a lock.
 */
function exclusively(run: () => Promise<void>): Promise<void> {
  const locks = (navigator as unknown as { locks?: { request(name: string, run: () => Promise<void>): Promise<void> } })
    .locks;
  return locks === undefined ? run() : locks.request('gessosheet-document-index', run);
}

/** Whether this environment has an Origin Private File System at all. */
function hasOpfs(): boolean {
  const storage = (navigator as unknown as { storage?: { getDirectory?: unknown } }).storage;
  return storage?.getDirectory !== undefined;
}

function isEntry(value: unknown): value is DocumentEntry {
  if (typeof value !== 'object' || value === null) {
    return false;
  }
  const entry = value as Record<string, unknown>;
  return typeof entry.id === 'string' && typeof entry.name === 'string' && typeof entry.used === 'number';
}
