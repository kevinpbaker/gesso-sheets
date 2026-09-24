import { describe, expect, it } from 'vitest';

import { DependencyGraph } from './DependencyGraph';

describe('DependencyGraph', () => {
  it('records both directions of an edge', () => {
    const graph = new DependencyGraph();
    graph.setPrecedents(2, [1]);
    expect([...graph.precedentsOf(2)]).toEqual([1]);
    expect([...graph.dependentsOf(1)]).toEqual([2]);
  });

  /**
   * A formula's references are rewritten as a unit when the formula is
   * retyped, so the old ones have to go with them. A graph that only
   * added edges would keep recomputing a cell because of a reference
   * it no longer contains.
   */
  it('forgets the edges a rewritten formula no longer has', () => {
    const graph = new DependencyGraph();
    graph.setPrecedents(3, [1, 2]);
    graph.setPrecedents(3, [2]);
    expect([...graph.precedentsOf(3)]).toEqual([2]);
    expect([...graph.dependentsOf(1)]).toEqual([]);
    expect([...graph.dependentsOf(2)]).toEqual([3]);
  });

  it('walks the closure transitively', () => {
    const graph = new DependencyGraph();
    graph.setPrecedents(2, [1]);
    graph.setPrecedents(3, [2]);
    graph.setPrecedents(4, [3]);
    expect([...graph.closureOf([1])].sort()).toEqual([2, 3, 4]);
    expect([...graph.closureOf([3])].sort()).toEqual([4]);
    expect([...graph.closureOf([4])]).toEqual([]);
  });

  it('visits a shared dependent once, not once per path', () => {
    const graph = new DependencyGraph();
    graph.setPrecedents(2, [1]);
    graph.setPrecedents(3, [1]);
    graph.setPrecedents(4, [2, 3]);
    expect([...graph.closureOf([1])].sort()).toEqual([2, 3, 4]);
  });

  it('terminates on a cycle rather than walking it forever', () => {
    const graph = new DependencyGraph();
    graph.setPrecedents(1, [2]);
    graph.setPrecedents(2, [1]);
    expect([...graph.closureOf([1])].sort()).toEqual([1, 2]);
  });

  describe('ordering', () => {
    it('puts a cell after everything it reads', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(2, [1]);
      graph.setPrecedents(3, [2]);
      const { order, circular } = graph.topological(new Set([2, 3]));
      expect(order).toEqual([2, 3]);
      expect(circular).toEqual([]);
    });

    /**
     * Edges leaving the set are ignored: those cells are already
     * correct and are not being recomputed, so waiting on them would
     * deadlock every recalc that did not start from the whole sheet.
     */
    it('ignores precedents outside the set being ordered', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(2, [1]);
      const { order } = graph.topological(new Set([2]));
      expect(order).toEqual([2]);
    });

    it('reports a cycle instead of ordering it', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(1, [2]);
      graph.setPrecedents(2, [1]);
      const { order, circular } = graph.topological(new Set([1, 2]));
      expect(order).toEqual([]);
      expect(circular.sort()).toEqual([1, 2]);
    });

    it('reports a cell that reads itself', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(1, [1]);
      const { order, circular } = graph.topological(new Set([1]));
      expect(order).toEqual([]);
      expect(circular).toEqual([1]);
    });

    /**
     * A cell downstream of a cycle is reported with it. It has no
     * value either — the cycle has no answer to give it — and every
     * spreadsheet says so in the same breath.
     */
    it('reports what is downstream of a cycle along with it', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(1, [2]);
      graph.setPrecedents(2, [1]);
      graph.setPrecedents(3, [1]);
      const { order, circular } = graph.topological(new Set([1, 2, 3]));
      expect(order).toEqual([]);
      expect(circular.sort()).toEqual([1, 2, 3]);
    });

    it('orders what it can and reports the rest', () => {
      const graph = new DependencyGraph();
      graph.setPrecedents(10, [9]);
      graph.setPrecedents(1, [2]);
      graph.setPrecedents(2, [1]);
      const { order, circular } = graph.topological(new Set([10, 1, 2]));
      expect(order).toEqual([10]);
      expect(circular.sort()).toEqual([1, 2]);
    });
  });
});
