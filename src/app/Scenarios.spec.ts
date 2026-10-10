import { describe, expect, it } from 'vitest';

import {
  cellIn,
  type SheetClipboard,
  type SheetCompare,
  type SheetEditor,
  type SheetFormatWindow,
  type SheetPalette,
  type SheetScenarios,
  type SheetTransfer,
  type SheetWindow
} from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { SheetService, type Schedule } from './SheetService';
import { SCENARIO_CHANGED, SCENARIO_TYPED } from './scenarioPaint';

/**
 * Scenarios, end to end through the service: a named set of inputs the
 * workbook is forked with, shown in place of the base, with what it
 * typed and what it moved tinted.
 *
 * The model is three cells. B1 is a rate, B2 a price, B3 what they come
 * to; A5 holds a note nothing reads.
 */
function harness() {
  const queue: (() => void)[] = [];
  const schedule: Schedule = run => queue.push(run);
  const document = new SheetDocument();
  const service = new SheetService(document, { schedule });
  const drain = () => {
    let guard = 0;
    while (queue.length > 0) {
      queue.shift()!();
      if (guard++ > 10_000) {
        throw new Error('the pump never finished');
      }
    }
  };
  service.setViewport(0, 0, 9, 0, 3);
  service.setCell(0, 1, '0.1');
  service.setCell(1, 1, '20');
  service.setCell(2, 1, '=B2*(1+B1)');
  service.setCell(4, 0, 'Prices are ex works');
  drain();
  return { document, service, drain };
}

function latest<T>(source: { subscribe(next: (value: T) => void): { unsubscribe(): void } }): T {
  let value!: T;
  source.subscribe(next => (value = next)).unsubscribe();
  return value;
}

const shown = (service: SheetService, row: number, column: number) => cellIn(latest<SheetWindow>(service.window), row, column);

/** The fill a cell is painted with, or '' for none. */
function fillOf(service: SheetService, row: number, column: number): string {
  const id = latest<SheetFormatWindow>(service.formats).cells[row]?.[column] ?? 0;
  return latest<SheetPalette>(service.palette).entries[id]?.fill ?? '';
}

describe('a scenario', () => {
  it('is shown once added, and typing into it leaves the base alone', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    drain();
    expect(latest<SheetScenarios>(service.scenarios).shown).not.toBeNull();

    service.setCell(0, 1, '0.5');
    drain();
    expect(shown(service, 0, 1)).toBe('0.5');
    expect(shown(service, 2, 1)).toBe('30');
    expect(document.book.value(0, 0, 1)).toBe(0.1);
    expect(document.book.value(0, 2, 1)).toBe(22);

    service.showScenario(null);
    drain();
    expect(shown(service, 2, 1)).toBe('22');
  });

  it('tints what it typed and what that moved, and nothing else', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    drain();
    expect(fillOf(service, 0, 1)).toBe(SCENARIO_TYPED);
    expect(fillOf(service, 2, 1)).toBe(SCENARIO_CHANGED);
    expect(fillOf(service, 1, 1)).toBe('');
    expect(fillOf(service, 4, 0)).toBe('');

    service.showScenario(null);
    drain();
    expect(fillOf(service, 0, 1)).toBe('');
    expect(fillOf(service, 2, 1)).toBe('');
  });

  it('tells the editor what the base says about the active cell', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    drain();
    service.setSelection(2, 1, 2, 1);
    expect(latest<SheetEditor>(service.editor).scenario).toEqual({ name: 'Optimistic', base: '22', typed: false });
    service.setSelection(0, 1, 0, 1);
    expect(latest<SheetEditor>(service.editor).scenario).toEqual({ name: 'Optimistic', base: '0.1', typed: true });
    expect(latest<SheetEditor>(service.editor).input).toBe('0.5');
    service.setSelection(1, 1, 1, 1);
    expect(latest<SheetEditor>(service.editor).scenario).toBeNull();
  });

  it('takes typing back with undo, and gives cells back to the base', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.setCell(1, 1, '40');
    drain();
    service.undo();
    drain();
    expect(shown(service, 1, 1)).toBe('20');
    expect(shown(service, 2, 1)).toBe('30');

    service.setSelection(0, 1, 0, 1);
    service.resetScenarioCells();
    drain();
    expect(shown(service, 2, 1)).toBe('22');
    expect(latest<SheetScenarios>(service.scenarios).entries[0].inputs).toBe(0);
  });

  it('typing the base’s own input gives the cell back to it', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.setCell(0, 1, '0.1');
    drain();
    expect(document.overridesOf(document.scenario!)).toEqual([]);
  });

  it('follows a change to the base it was forked from', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.showScenario(null);
    service.setCell(1, 1, '100');
    drain();
    const id = latest<SheetScenarios>(service.scenarios).entries[0].id;
    service.showScenario(id);
    drain();
    expect(shown(service, 2, 1)).toBe('150');
  });

  it('starts from the scenario showing when it is made as a copy', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.addScenario('Optimistic', true);
    drain();
    const scenarios = latest<SheetScenarios>(service.scenarios);
    expect(scenarios.entries.map(entry => entry.name)).toEqual(['Optimistic', 'Optimistic 2']);
    expect(scenarios.shown).toBe(scenarios.entries[1].id);
    expect(shown(service, 2, 1)).toBe('30');
  });

  it('is never called Base, and two are never called the same', () => {
    const { service, drain } = harness();
    service.addScenario('Base', false);
    service.addScenario('Low', false);
    const [first, second] = latest<SheetScenarios>(service.scenarios).entries;
    service.renameScenario(second.id, 'base');
    drain();
    const names = latest<SheetScenarios>(service.scenarios).entries.map(entry => entry.name);
    expect(names[0]).toBe('Base 2');
    // Base 2 is taken, whatever its capitals.
    expect(names[1]).toBe('base 3');
    expect(first.id).not.toBe(second.id);
  });

  it('moves its inputs with a row inserted above them', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.setSelection(0, 0, 0, 0);
    service.insertRows(0, 1);
    drain();
    expect(document.overridesOf(document.scenario!)).toMatchObject([{ row: 1, column: 1, input: '0.5' }]);
    expect(shown(service, 3, 1)).toBe('30');
  });

  it('goes, with what it typed, when it is deleted', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    const id = document.scenario!;
    service.deleteScenario(id);
    drain();
    expect(latest<SheetScenarios>(service.scenarios)).toEqual({ entries: [], shown: null });
    expect(shown(service, 2, 1)).toBe('22');
    expect(document.overridesOf(id)).toEqual([]);
  });

  it('copies what it shows, and pastes into itself', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    drain();
    service.setSelection(0, 1, 0, 1);
    service.copy(false);
    const text = latest<SheetClipboard>(service.clipboard).text;
    expect(text).toBe('0.5');
    service.setSelection(0, 2, 0, 2);
    service.paste(text);
    drain();
    expect(shown(service, 0, 2)).toBe('0.5');
    expect(document.book.input(0, 0, 2)).toBe('');
  });

  it('will not sort, because the order of rows is the base’s', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setSelection(0, 1, 2, 1);
    service.sortRange(1, false, false);
    drain();
    expect(latest<SheetTransfer>(service.transfer).report).toBe('Sort with Base showing: a scenario changes values, not the order of rows.');
    expect(document.book.input(0, 0, 1)).toBe('0.1');
  });

  it('is kept by the file, and opens with the base showing', () => {
    const { document, service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    drain();
    const text = JSON.stringify(snapshotOf(document, 1000));
    const parsed = parseSnapshot(text, 26);
    expect(parsed).not.toBeNull();
    const reopened = new SheetDocument();
    applySnapshot(reopened, parsed!);
    expect(reopened.scenarios).toEqual(document.scenarios);
    expect(reopened.scenario).toBeNull();
    expect(reopened.overridesOf(reopened.scenarios[0].id)).toMatchObject([{ sheet: 0, row: 0, column: 1, input: '0.5' }]);
  });
});

describe('the second pane', () => {
  const beside = (service: SheetService, row: number, column: number) =>
    cellIn(latest<SheetWindow>(service.compareWindow), row, column);
  function besideFill(service: SheetService, row: number, column: number): string {
    const id = latest<SheetFormatWindow>(service.compareFormats).cells[row]?.[column] ?? 0;
    return latest<SheetPalette>(service.palette).entries[id]?.fill ?? '';
  }

  it('shows the base beside a scenario, tinted where the two differ', () => {
    const { service, drain } = harness();
    service.addScenario('Optimistic', false);
    service.setCell(0, 1, '0.5');
    service.setCompare(true, null);
    drain();
    expect(latest<SheetCompare>(service.compare)).toEqual({ open: true, against: null, name: 'Base' });
    expect(shown(service, 2, 1)).toBe('30');
    expect(beside(service, 2, 1)).toBe('22');
    expect(beside(service, 0, 1)).toBe('0.1');
    expect(besideFill(service, 2, 1)).toBe(SCENARIO_CHANGED);
    expect(besideFill(service, 1, 1)).toBe('');
    expect(beside(service, 4, 0)).toBe('Prices are ex works');
  });

  it('shows another scenario, and follows an edit to the base in both', () => {
    const { service, drain } = harness();
    service.addScenario('Low', false);
    service.setCell(0, 1, '0');
    const low = latest<SheetScenarios>(service.scenarios).shown;
    service.addScenario('High', false);
    service.setCell(0, 1, '1');
    service.setCompare(true, low);
    drain();
    expect(shown(service, 2, 1)).toBe('40');
    expect(beside(service, 2, 1)).toBe('20');
    service.showScenario(null);
    service.setCell(1, 1, '10');
    drain();
    expect(shown(service, 2, 1)).toBe('11');
    expect(beside(service, 2, 1)).toBe('10');
  });

  it('goes back to the base when the scenario it shows is deleted, and empties when closed', () => {
    const { service, drain } = harness();
    service.addScenario('Low', false);
    service.setCell(0, 1, '0');
    const low = latest<SheetScenarios>(service.scenarios).shown!;
    service.showScenario(null);
    service.setCompare(true, low);
    drain();
    expect(beside(service, 2, 1)).toBe('20');
    service.deleteScenario(low);
    drain();
    expect(latest<SheetCompare>(service.compare).against).toBeNull();
    expect(beside(service, 2, 1)).toBe('22');
    service.setCompare(false, null);
    drain();
    expect(latest<SheetWindow>(service.compareWindow).lastRow).toBe(-1);
  });
});
