import { InMemorySheetRepository, type SheetRepository } from './SheetRepository';
import { InMemoryVersionStore, type VersionStore } from './SheetVersions';

/**
 * Every workbook this browser keeps, and where each one is kept.
 *
 * One level above `SheetRepository`, which is where *a* workbook is
 * kept: a library is the list of them, so that a tab can say which one
 * it is showing and a new tab can start another. The same seam, for
 * the same reason — nothing above it knows whether there is a disk —
 * and the one implementation that needs a browser is the only one that
 * needs one.
 *
 * A document is not a file. It is this application's own copy, saved
 * on every edit whether or not anybody has chosen a file for it; a
 * file is where it was last saved *to*, which is what `file` records.
 * That split is the whole of how a tab can be closed without a Save
 * and lose nothing, while a Save still means something.
 */

export interface DocumentEntry {
  readonly id: string;
  /** What to call it: the file's name without its extension, or "Untitled". */
  readonly name: string;
  /** When it was last opened, in epoch milliseconds. */
  readonly used: number;
  /**
   * The file it was last saved to or opened from, or null.
   *
   * `handle` is the shell's number for the file — see
   * `ShellService.saveFile` — or null when it was saved as a download
   * and there is nothing to save back to.
   */
  readonly file: DocumentFile | null;
  /**
   * The function scripts somebody turned on for this document, as
   * fingerprints of what each said when they did: SHA-256 of its name
   * and its source. Kept here, in this browser's library, and never in
   * the file, so a file cannot say it has been trusted; and kept as
   * what the code *was*, so a file that comes back changed is off
   * again. Absent is none.
   */
  readonly trustedFunctions?: readonly string[];
  /**
   * When its contents last changed, in epoch milliseconds, to within
   * half a minute; absent for a document from before Phase 36, whose
   * `used` is the best there is. The home screen's *Last edited*.
   */
  readonly edited?: number;
  /**
   * Whether every change is written to `file` as well as to this
   * browser — *Keep saving to this file*. Only means anything while
   * `file` has a handle.
   */
  readonly autosave?: boolean;
}

export interface DocumentFile {
  readonly handle: number | null;
  readonly name: string;
}

export interface SheetLibrary {
  /** Every document, in no particular order. */
  entries(): Promise<DocumentEntry[]>;
  /** Adds or replaces a document's entry. */
  put(entry: DocumentEntry): Promise<void>;
  /** Where a document's contents are kept. */
  repository(id: string): SheetRepository;
  /** An id no document has. */
  newId(): string;
  /**
   * Holds a document for this tab, or answers null when another tab
   * already holds it. The function returned lets it go.
   *
   * What stops two tabs over one document writing over each other: the
   * copy each keeps in memory is its own, and whichever saved last
   * would win, silently, with the other's edits gone.
   */
  claim(id: string): Promise<(() => void) | null>;
  /** Forgets a document: its entry, its contents and its versions. */
  remove(id: string): Promise<void>;
  /** Where a document's restore points are kept; see `SheetVersions.ts`. */
  versions(id: string): VersionStore;
}

/**
 * The id of the first document, whose contents are the file every
 * build before this phase wrote. Kept rather than migrated: a workbook
 * somebody has been typing into since Phase 6 is theirs, and copying
 * it somewhere new is one more thing that could lose it.
 */
export const FIRST_DOCUMENT = 'main';

/** A library that forgets, for specs. */
export class InMemorySheetLibrary implements SheetLibrary {
  private readonly stored = new Map<string, DocumentEntry>();
  private readonly repositories = new Map<string, InMemorySheetRepository>();
  private next = 1;

  entries(): Promise<DocumentEntry[]> {
    return Promise.resolve([...this.stored.values()]);
  }

  put(entry: DocumentEntry): Promise<void> {
    this.stored.set(entry.id, entry);
    return Promise.resolve();
  }

  repository(id: string): InMemorySheetRepository {
    let repository = this.repositories.get(id);
    if (repository === undefined) {
      repository = new InMemorySheetRepository();
      this.repositories.set(id, repository);
    }
    return repository;
  }

  newId(): string {
    return `d${this.next++}`;
  }

  private readonly held = new Set<string>();
  private readonly versionStores = new Map<string, InMemoryVersionStore>();

  remove(id: string): Promise<void> {
    this.stored.delete(id);
    this.repositories.delete(id);
    this.versionStores.delete(id);
    return Promise.resolve();
  }

  versions(id: string): InMemoryVersionStore {
    let store = this.versionStores.get(id);
    if (store === undefined) {
      store = new InMemoryVersionStore();
      this.versionStores.set(id, store);
    }
    return store;
  }

  claim(id: string): Promise<(() => void) | null> {
    if (this.held.has(id)) {
      return Promise.resolve(null);
    }
    this.held.add(id);
    return Promise.resolve(() => this.held.delete(id));
  }
}
