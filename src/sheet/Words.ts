/**
 * The distinct words a column holds, for AutoComplete.
 *
 * Kept sorted, case folded, so that "what does `No` complete to" is a
 * binary search and not a walk down the column: the question is asked
 * on every keystroke, and a column of a hundred thousand entries must
 * answer it as quickly as a column of ten. Each word carries a count,
 * because the same region is typed on a thousand rows and deleting one
 * of them must not forget the other nine hundred and ninety-nine.
 *
 * Plain data and no imports, like everything else in `src/sheet`.
 */
export class Words {
  /** Folded words, sorted, one each. */
  private readonly folded: string[] = [];
  /** A folded word's spelling as first typed, and how many cells hold it. */
  private readonly entries = new Map<string, { text: string; count: number }>();

  get size(): number {
    return this.folded.length;
  }

  add(text: string): void {
    const key = fold(text);
    const entry = this.entries.get(key);
    if (entry !== undefined) {
      entry.count++;
      return;
    }
    this.entries.set(key, { text, count: 1 });
    this.folded.splice(lowerBound(this.folded, key), 0, key);
  }

  remove(text: string): void {
    const key = fold(text);
    const entry = this.entries.get(key);
    if (entry === undefined) {
      return;
    }
    if (--entry.count > 0) {
      return;
    }
    this.entries.delete(key);
    this.folded.splice(lowerBound(this.folded, key), 1);
  }

  /**
   * The one word `prefix` starts, as it was spelled, or null.
   *
   * Null for nothing, for more than one — `No` could be North or
   * Northwest, and guessing is how AutoComplete puts the wrong region
   * on a row — and for a prefix that is already the whole word, which
   * leaves nothing to offer.
   */
  complete(prefix: string): string | null {
    const key = fold(prefix);
    if (key === '') {
      return null;
    }
    const at = lowerBound(this.folded, key);
    const found = this.folded[at];
    if (found === undefined || !found.startsWith(key) || found === key) {
      return null;
    }
    const next = this.folded[at + 1];
    if (next !== undefined && next.startsWith(key)) {
      return null;
    }
    return this.entries.get(found)?.text ?? null;
  }
}

function fold(text: string): string {
  return text.toLocaleLowerCase();
}

/** The first index whose word is not less than `key`. */
function lowerBound(sorted: readonly string[], key: string): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (sorted[middle] < key) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
}
