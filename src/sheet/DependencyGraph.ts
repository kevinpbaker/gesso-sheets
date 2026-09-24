/**
 * Who reads whom.
 *
 * Two adjacency maps over packed cell keys. `precedents` is what a
 * formula reads and is rewritten whole whenever that formula changes;
 * `dependents` is its inverse and is what a recalc walks, since the
 * question an edit asks is "what has to be redone", which is the
 * inverse direction.
 *
 * Ranges are expanded to their cells. `=SUM(A1:A100)` stores a hundred
 * edges, which is the honest thing at this phase and wrong at scale:
 * a sheet with a column of running totals stores a triangle of them.
 * The fix is an interval index that answers "which formulas cover this
 * cell" without an edge per cell, and it is a Phase 1 follow-up rather
 * than part of it — the recalc budget this phase is judged on counts
 * *evaluations*, which an interval index would not change.
 */
export class DependencyGraph {
  private readonly precedents = new Map<number, Set<number>>();
  private readonly dependents = new Map<number, Set<number>>();

  /**
   * Replaces everything `key` reads.
   *
   * Whole rather than incremental because a formula's references are
   * rewritten as a unit when the formula is retyped, and the diff of
   * two reference sets costs more than rebuilding the smaller one.
   */
  setPrecedents(key: number, precedents: Iterable<number>): void {
    this.clearPrecedents(key);
    let owned: Set<number> | undefined;
    for (const precedent of precedents) {
      // A cell that reads itself is recorded like any other edge, so
      // that the topological pass reports the cycle of one the same
      // way it reports a cycle of five rather than by a special case.
      owned ??= new Set();
      owned.add(precedent);
      this.dependentsOfMutable(precedent).add(key);
    }
    if (owned !== undefined) {
      this.precedents.set(key, owned);
    }
  }

  /** Forgets that `key` reads anything. Its readers are untouched. */
  clearPrecedents(key: number): void {
    const existing = this.precedents.get(key);
    if (existing === undefined) {
      return;
    }
    for (const precedent of existing) {
      const readers = this.dependents.get(precedent);
      if (readers === undefined) {
        continue;
      }
      readers.delete(key);
      if (readers.size === 0) {
        this.dependents.delete(precedent);
      }
    }
    this.precedents.delete(key);
  }

  precedentsOf(key: number): ReadonlySet<number> {
    return this.precedents.get(key) ?? EMPTY;
  }

  dependentsOf(key: number): ReadonlySet<number> {
    return this.dependents.get(key) ?? EMPTY;
  }

  /**
   * Every cell that has to be recomputed when `roots` change,
   * transitively, not including the roots themselves.
   *
   * A breadth-first walk with a visited set, so a diamond — two
   * formulas reading one cell and a third reading both — visits the
   * third once. The visited set is also what makes a cycle terminate
   * here; reporting it is the topological pass's job.
   */
  closureOf(roots: Iterable<number>): Set<number> {
    const closure = new Set<number>();
    const queue: number[] = [];
    for (const root of roots) {
      for (const dependent of this.dependentsOf(root)) {
        if (!closure.has(dependent)) {
          closure.add(dependent);
          queue.push(dependent);
        }
      }
    }
    for (let at = 0; at < queue.length; at++) {
      for (const dependent of this.dependentsOf(queue[at])) {
        if (!closure.has(dependent)) {
          closure.add(dependent);
          queue.push(dependent);
        }
      }
    }
    return closure;
  }

  /**
   * `cells` in an order where every cell comes after the cells it
   * reads, and separately those that cannot be ordered at all.
   *
   * Kahn's algorithm over the subgraph induced by `cells`: edges to
   * cells outside the set are ignored, because those are already
   * correct and are not being recomputed. What is left with a
   * non-zero in-degree at the end is in a cycle or downstream of one,
   * and the two are not distinguished — a cell reading a circular cell
   * has no value either, and saying so in the same breath is what
   * every spreadsheet does.
   */
  topological(cells: ReadonlySet<number>): { order: number[]; circular: number[] } {
    const remaining = new Map<number, number>();
    for (const cell of cells) {
      let indegree = 0;
      for (const precedent of this.precedentsOf(cell)) {
        if (cells.has(precedent)) {
          indegree++;
        }
      }
      remaining.set(cell, indegree);
    }

    const order: number[] = [];
    const ready: number[] = [];
    for (const [cell, indegree] of remaining) {
      if (indegree === 0) {
        ready.push(cell);
      }
    }
    for (let at = 0; at < ready.length; at++) {
      const cell = ready[at];
      order.push(cell);
      remaining.delete(cell);
      for (const dependent of this.dependentsOf(cell)) {
        const indegree = remaining.get(dependent);
        if (indegree === undefined) {
          continue;
        }
        if (indegree === 1) {
          remaining.set(dependent, 0);
          ready.push(dependent);
        } else {
          remaining.set(dependent, indegree - 1);
        }
      }
    }
    return { order, circular: [...remaining.keys()] };
  }

  private dependentsOfMutable(key: number): Set<number> {
    let set = this.dependents.get(key);
    if (set === undefined) {
      set = new Set();
      this.dependents.set(key, set);
    }
    return set;
  }
}

const EMPTY: ReadonlySet<number> = new Set();
