import { columnKeyOf } from './A1';

/**
 * Who reads whom.
 *
 * Two adjacency maps over packed cell keys. `precedents` is what a
 * formula reads and is rewritten whole whenever that formula changes;
 * `dependents` is its inverse and is what a recalc walks, since the
 * question an edit asks is "what has to be redone", which is the
 * inverse direction.
 *
 * Ranges written as rectangles are expanded to their cells.
 * `=SUM(A1:A100)` stores a hundred edges, which is honest at this
 * scale and wrong at a larger one: a sheet with a column of running
 * totals stores a triangle of them. The recalc budget the engine is
 * judged on counts *evaluations*, which an index would not change, so
 * the expansion has stayed.
 *
 * **A whole-column reference cannot be expanded at all.** `=SUM(A:A)`
 * covers 1,048,576 cells and Phase 11 made it writable, so columns are
 * *watched* instead: one entry per column per formula, and
 * `dependentsOf` unions the watchers in. That is the interval index
 * this file has wanted since Phase 1, built for the one case that
 * cannot live without it rather than for all of them — because the
 * case that cannot live without it is also the case where the index is
 * trivially exact, a column being a whole coordinate rather than a
 * span.
 */
export class DependencyGraph {
  private readonly precedents = new Map<number, Set<number>>();
  private readonly dependents = new Map<number, Set<number>>();
  /** Formulas reading a whole column, by the column they read. */
  private readonly columnWatchers = new Map<number, Set<number>>();
  /** And the inverse, so a formula can take its watches back. */
  private readonly watchedColumns = new Map<number, number[]>();

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

  /**
   * Forgets the whole graph.
   *
   * For a structural change — an insert or a delete — where every
   * edge on the moved side of the line points at the wrong key.
   * Rebuilding from the formulas is the only version that cannot be
   * half-right, and patching a key at a time would leave the two
   * adjacency maps disagreeing the moment one of them was half done.
   */
  clear(): void {
    this.precedents.clear();
    this.dependents.clear();
    this.columnWatchers.clear();
    this.watchedColumns.clear();
  }

  /**
   * Records that `key` reads every cell of these columns.
   *
   * Replaces whatever it watched before, on the same rule as
   * `setPrecedents`: a formula's references are rewritten as a unit.
   */
  setWatchedColumns(key: number, columns: readonly number[]): void {
    this.clearWatchedColumns(key);
    if (columns.length === 0) {
      return;
    }
    this.watchedColumns.set(key, [...columns]);
    for (const column of columns) {
      let watchers = this.columnWatchers.get(column);
      if (watchers === undefined) {
        watchers = new Set();
        this.columnWatchers.set(column, watchers);
      }
      watchers.add(key);
    }
  }

  private clearWatchedColumns(key: number): void {
    const columns = this.watchedColumns.get(key);
    if (columns === undefined) {
      return;
    }
    for (const column of columns) {
      const watchers = this.columnWatchers.get(column);
      if (watchers === undefined) {
        continue;
      }
      watchers.delete(key);
      if (watchers.size === 0) {
        this.columnWatchers.delete(column);
      }
    }
    this.watchedColumns.delete(key);
  }

  /** Forgets that `key` reads anything. Its readers are untouched. */
  clearPrecedents(key: number): void {
    this.clearWatchedColumns(key);
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

  /**
   * What reads a cell: the edges to it, plus anything watching its
   * column.
   *
   * Unioned here rather than at each call site because every walk in
   * this file asks the same question, and a walk that forgot the
   * watchers would leave a `=SUM(A:A)` stale after a write to A900 —
   * silently, which is the failure mode a spreadsheet cannot have.
   */
  dependentsOf(key: number): ReadonlySet<number> {
    const direct = this.dependents.get(key) ?? EMPTY;
    const watchers = this.columnWatchers.get(columnKeyOf(key)) ?? EMPTY;
    if (watchers.size === 0) {
      return direct;
    }
    if (direct.size === 0) {
      return watchers;
    }
    const both = new Set(direct);
    for (const watcher of watchers) {
      both.add(watcher);
    }
    return both;
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
    /**
     * In-degrees, counted forwards rather than backwards.
     *
     * The obvious spelling asks each cell for its precedents and
     * counts the ones in the set. That cannot see a column watch —
     * `precedentsOf` has no entry for the million cells of `A:A` and
     * must not — so the count is taken the other way instead: walk
     * the set once and add one to each dependent, which goes through
     * `dependentsOf` and therefore through the watchers. Same number,
     * same cost, and it works for both kinds of edge.
     */
    const remaining = new Map<number, number>();
    for (const cell of cells) {
      remaining.set(cell, 0);
    }
    for (const cell of cells) {
      for (const dependent of this.dependentsOf(cell)) {
        const indegree = remaining.get(dependent);
        if (indegree !== undefined) {
          remaining.set(dependent, indegree + 1);
        }
      }
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
