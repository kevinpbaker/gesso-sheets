import { describe, expect, it } from 'vitest';

import { keyOn } from './A1';
import { NUM } from './Values';
import { SAMPLE_LIMIT, Simulation, seeded, spreadOf } from './Simulation';
import { Workbook } from './Workbook';

/**
 * Guesses and what they come to.
 *
 * A guess is its likeliest value until a simulation draws it, and a
 * simulation draws every guess in every trial, on a fork, so the
 * workbook somebody is editing never moves. The model is a price with a
 * spread, units with a range, and their product.
 */
function model(): Workbook {
  const w = new Workbook(['Plan']);
  w.setCell(0, 0, 0, '=NORMAL(100, 10)'); // A1, the price
  w.setCell(0, 1, 0, '=UNIFORM(40, 60)'); // A2, the units
  w.setCell(0, 2, 0, '=A1*A2'); // A3, the revenue
  w.setCell(0, 3, 0, '7'); // A4, a number nothing guesses at
  w.setCell(0, 4, 0, '=A4*2'); // A5, nor at this
  w.recalculate();
  return w;
}

const at = (row: number) => keyOn(0, row, 0);

function run(w: Workbook, trials: number, seed = 1): Simulation {
  const simulation = new Simulation(w, trials, seed);
  while (!simulation.finished) {
    simulation.step(1_000);
  }
  return simulation;
}

describe('a guess', () => {
  it('is its likeliest value outside a simulation', () => {
    const w = model();
    expect(w.value(0, 0, 0)).toBe(100);
    expect(w.value(0, 1, 0)).toBe(50);
    expect(w.value(0, 2, 0)).toBe(5000);
    const t = new Workbook(['Plan']);
    t.setCell(0, 0, 0, '=TRIANGULAR(0.05, 0.061, 0.08)');
    t.recalculate();
    expect(t.value(0, 0, 0)).toBe(0.061);
  });

  it('refuses a spread that cannot be', () => {
    const w = new Workbook(['Plan']);
    w.setCell(0, 0, 0, '=NORMAL(1, -1)');
    w.setCell(0, 1, 0, '=UNIFORM(9, 1)');
    w.setCell(0, 2, 0, '=TRIANGULAR(1, 5, 3)');
    w.recalculate();
    expect(w.value(0, 0, 0)).toBe(NUM);
    expect(w.value(0, 1, 0)).toBe(NUM);
    expect(w.value(0, 2, 0)).toBe(NUM);
  });

  it('is not volatile: an edit elsewhere leaves it as it was', () => {
    const w = model();
    w.setCell(0, 3, 0, '8');
    expect(w.pending).toBe(1);
  });
});

describe('a simulation', () => {
  it('reaches the guesses and what reads them, and nothing else', () => {
    const w = model();
    const simulation = run(w, 200);
    expect(simulation.cells).toBe(3);
    expect(simulation.covers(at(2))).toBe(true);
    expect(simulation.spreadOf(at(4))).toBeNull();
  });

  it('draws around the guess: a bell centred on its mean, a range inside its ends', () => {
    const simulation = run(model(), 4_000);
    const price = simulation.spreadOf(at(0))!;
    expect(price.count).toBe(4_000);
    expect(price.mean).toBeGreaterThan(99);
    expect(price.mean).toBeLessThan(101);
    // Ten per cent of a bell lies beyond 1.28 deviations either side.
    expect(price.p10).toBeGreaterThan(85);
    expect(price.p10).toBeLessThan(89);
    expect(price.p90).toBeGreaterThan(111);
    expect(price.p90).toBeLessThan(115);
    const units = simulation.spreadOf(at(1))!;
    expect(units.min).toBeGreaterThanOrEqual(40);
    expect(units.max).toBeLessThanOrEqual(60);
    const revenue = simulation.spreadOf(at(2))!;
    expect(revenue.p10).toBeLessThan(revenue.p50);
    expect(revenue.p50).toBeLessThan(revenue.p90);
    expect(revenue.bins.reduce((sum, bin) => sum + bin, 0)).toBe(4_000);
  });

  it('leaves the workbook being edited where it was', () => {
    const w = model();
    run(w, 500);
    expect(w.value(0, 2, 0)).toBe(5000);
    expect(w.pending).toBe(0);
    expect(w.sampler).toBeNull();
  });

  it('says the same thing twice with the same seed, and something else with another', () => {
    const one = run(model(), 300, 42).spreadOf(at(2))!;
    const two = run(model(), 300, 42).spreadOf(at(2))!;
    const three = run(model(), 300, 43).spreadOf(at(2))!;
    expect(two).toEqual(one);
    expect(three.mean).not.toBe(one.mean);
  });

  it('runs in pieces, and its spread so far is the trials so far', () => {
    let clock = 0;
    const simulation = new Simulation(model(), 1_000, 1);
    // A clock that moves a millisecond a reading: two readings a trial.
    const ran = simulation.step(5, () => clock++);
    expect(ran).toBeGreaterThan(0);
    expect(ran).toBeLessThan(1_000);
    expect(simulation.spreadOf(at(0))!.count).toBe(ran);
  });

  it('cuts its trials down to stay inside the memory it is allowed', () => {
    const w = model();
    expect(new Simulation(w, SAMPLE_LIMIT, 1).trials).toBe(Math.floor(SAMPLE_LIMIT / 3));
  });

  it('draws the same seed the same way', () => {
    const a = seeded(9);
    const b = seeded(9);
    expect([a(), a(), a()]).toEqual([b(), b(), b()]);
  });

  it('bins a spread that never moved into one bar', () => {
    const spread = spreadOf(Float64Array.from([3, 3, 3]), 3)!;
    expect(spread.p50).toBe(3);
    expect(spread.bins.filter(bin => bin > 0)).toEqual([3]);
  });
});
