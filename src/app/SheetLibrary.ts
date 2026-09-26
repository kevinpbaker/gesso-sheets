import { InMemorySheetRepository, type SheetRepository } from './SheetRepository';

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
}
