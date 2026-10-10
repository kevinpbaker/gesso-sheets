import { describe, expect, it } from 'vitest';

import { type SheetEditor, type SheetSimulation, type SheetUncertainty } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { SheetService, type Schedule } from './SheetService';

/**
 * A simulation through the service: run in slices on the application
 * worker, read out as histograms for the visible cells and as
 * percentiles for the active one, and put away by the first edit.
 *
 * B1 is a growth with a spread, B2 a start, B3 a year of it; D1 is a
 * number nothing guesses at.
 */
function harness(budget = 2_000) {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule, budget });
  const drain = () => {
    let guard = 0;
    while (queue.length > 0) {
      queue.shift()!();
      if (guard++ > 100_000) {
        throw new Error('the pump never finished');
      }
    }
  };
  service.setViewport(0, 0, 9, 0, 4);
  service.setCell(0, 1, '=NORMAL(0.05, 0.02)');
  service.setCell(1, 1, '1000');
  service.setCell(2, 1, '=B2*(1+B1)^12');
  service.setCell(0, 3, '7');
  drain();
  return { document, service, drain, queue };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

describe('a simulation, through the service', () => {
  it('knows there is a guess to run before anything has run', () => {
    const { service } = harness();
    expect(latest<SheetSimulation>(service.simulation)).toMatchObject({ state: 'idle', guesses: 1, on: 'Base' });
  });

  it('runs in slices and finishes, reaching the guess and what reads it', () => {
    const { service, drain, queue } = harness();
    service.runSimulation(2_000);
    expect(latest<SheetSimulation>(service.simulation).state).toBe('running');
    expect(queue.length).toBeGreaterThan(0);
    drain();
    expect(latest<SheetSimulation>(service.simulation)).toMatchObject({ state: 'done', done: 2_000, trials: 2_000, cells: 2 });
  });

  it('draws a histogram for each cell the guess reaches, and none for the rest', () => {
    const { service, drain } = harness();
    service.runSimulation(2_000);
    drain();
    const cells = latest<SheetUncertainty>(service.uncertainty).cells;
    expect(cells[0]?.[1]).toHaveLength(16);
    expect(Math.max(...cells[2][1])).toBe(8);
    expect(cells[1]?.[1]).toBeUndefined();
    expect(cells[0]?.[3]).toBeUndefined();
  });

  it('reads out the active cell’s percentiles, in its own format', () => {
    const { service, drain } = harness();
    service.runSimulation(2_000);
    drain();
    service.setSelection(2, 1, 2, 1);
    const spread = latest<SheetEditor>(service.editor).spread!;
    expect(spread.count).toBe(2_000);
    expect(Number(spread.p10)).toBeLessThan(Number(spread.p50));
    expect(Number(spread.p50)).toBeLessThan(Number(spread.p90));
    service.setSelection(0, 3, 0, 3);
    expect(latest<SheetEditor>(service.editor).spread).toBeNull();
  });

  it('is put away by the first edit, and the guesses keep their likeliest value', () => {
    const { document, service, drain } = harness();
    service.runSimulation(500);
    drain();
    service.setCell(1, 1, '2000');
    drain();
    expect(latest<SheetSimulation>(service.simulation).state).toBe('idle');
    expect(latest<SheetUncertainty>(service.uncertainty).cells).toEqual({});
    expect(document.book.value(0, 0, 1)).toBe(0.05);
  });

  it('stops where it is, keeping the trials it ran, and clears when stopped again', () => {
    const { service, queue } = harness();
    service.runSimulation(100_000);
    queue.shift()!();
    service.stopSimulation();
    const stopped = latest<SheetSimulation>(service.simulation);
    expect(stopped.state).toBe('done');
    expect(stopped.done).toBeGreaterThan(0);
    expect(stopped.done).toBeLessThan(100_000);
    service.stopSimulation();
    expect(latest<SheetSimulation>(service.simulation).state).toBe('idle');
  });

  it('waits for a recalculation to settle before it starts', () => {
    // A cell a slice, so the recalculation an edit starts is still going.
    const { service, drain } = harness(1);
    service.setCell(0, 4, '=B3+1');
    service.setCell(1, 1, '3000');
    service.runSimulation(300);
    expect(latest<SheetSimulation>(service.simulation).state).toBe('idle');
    drain();
    expect(latest<SheetSimulation>(service.simulation)).toMatchObject({ state: 'done', done: 300 });
  });

  it('runs on the scenario showing', () => {
    const { service, drain } = harness();
    service.addScenario('Bigger start', false);
    service.setCell(1, 1, '5000');
    drain();
    service.runSimulation(1_000);
    drain();
    expect(latest<SheetSimulation>(service.simulation).on).toBe('Bigger start');
    service.setSelection(2, 1, 2, 1);
    expect(Number(latest<SheetEditor>(service.editor).spread!.p50)).toBeGreaterThan(5_000);
  });
});
