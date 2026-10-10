import type { SheetSnapshot } from './SheetFile';

/**
 * Restore points: copies of a workbook as it was, kept beside it.
 *
 * The browser's copy of a document is written over on every edit, which
 * is what makes closing a tab lose nothing — and also what makes an
 * afternoon of edits impossible to take back once the undo stack has
 * gone with the tab. A version is the whole snapshot, the same one the
 * repository writes, taken now and then and kept under the time it was
 * taken, so the list reads as a history and any line of it can be put
 * back.
 *
 * When one is taken is the service's business (`SheetService.keepVersion`);
 * this is only where they are kept.
 */

/** Why a version was taken, which is what the list says beside its time. */
export type VersionReason = 'opened' | 'auto' | 'saved' | 'restored';

export interface SheetVersion {
  /** When it was taken, in epoch milliseconds. Also its key. */
  readonly at: number;
  readonly reason: VersionReason;
}

export interface VersionStore {
  /** Every version kept, newest first. */
  list(): Promise<SheetVersion[]>;
  /** Keeps a version, letting the oldest go past `VERSION_LIMIT`. */
  add(version: SheetVersion, snapshot: SheetSnapshot): Promise<void>;
  /** The snapshot a version holds, or null when it is not there. */
  get(at: number): Promise<SheetSnapshot | null>;
  /** Forgets every version, for a document being deleted. */
  clear(): Promise<void>;
}

/**
 * How many versions a document keeps.
 *
 * Thirty at ten minutes apart is five hours of steady editing, or a
 * month of a budget opened on Mondays; each is the size of the
 * workbook, so the limit is also what a document can cost on disk.
 */
export const VERSION_LIMIT = 30;

/** How long editing goes on before another version is taken. */
export const VERSION_EVERY_MS = 10 * 60 * 1000;

/** A version store that forgets, for specs. */
export class InMemoryVersionStore implements VersionStore {
  private readonly kept = new Map<number, { version: SheetVersion; snapshot: SheetSnapshot }>();

  list(): Promise<SheetVersion[]> {
    return Promise.resolve([...this.kept.values()].map(each => each.version).sort((a, b) => b.at - a.at));
  }

  add(version: SheetVersion, snapshot: SheetSnapshot): Promise<void> {
    this.kept.set(version.at, { version, snapshot });
    const oldest = [...this.kept.keys()].sort((a, b) => b - a).slice(VERSION_LIMIT);
    for (const at of oldest) {
      this.kept.delete(at);
    }
    return Promise.resolve();
  }

  get(at: number): Promise<SheetSnapshot | null> {
    return Promise.resolve(this.kept.get(at)?.snapshot ?? null);
  }

  clear(): Promise<void> {
    this.kept.clear();
    return Promise.resolve();
  }
}

/** What a version's reason reads as in the list. */
export function describeVersion(reason: VersionReason): string {
  switch (reason) {
    case 'opened':
      return 'As it was opened';
    case 'saved':
      return 'Saved to a file';
    case 'restored':
      return 'Before a restore';
    default:
      return 'While editing';
  }
}
