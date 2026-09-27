import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';

import { relativeRef } from '../sheet/A1';
import { SheetApp } from './SheetApp';
import type { SheetNames } from './SheetContract';
import { SheetDocument } from './SheetDocument';
import { applySnapshot, parseSnapshot, snapshotOf } from './SheetFile';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Naming a range, from the name box to the file and back.
 *
 * The gesture is the whole of `Insert ▸ Name`: pick a range, type
 * what it is into the box that already says where you are, press
 * Enter. There is no dialog, because there is nothing a dialog would
 * ask that the selection and the box do not already answer.
 */

const area = (row: number, column: number, lastRow: number, lastColumn: number) => ({
  start: relativeRef(row, column),
  end: relativeRef(lastRow, lastColumn)
});

describe('naming a range through the service', () => {
  function harness() {
    const document = new SheetDocument();
    document.setCell(1, 1, '10');
    document.setCell(2, 1, '20');
    const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
    return { document, service };
  }

  const names = (service: SheetService) => {
    let seen: SheetNames | undefined;
    service.names.subscribe(value => (seen = value)).unsubscribe();
    return seen!;
  };

  it('names the selection', () => {
    const { service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');

    expect(names(service).entries.map(entry => entry.name)).toEqual(['Sales']);
    expect(names(service).refused).toBe('');
  });

  it('publishes the corners, so the box can go there', () => {
    const { service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');
    expect(names(service).entries[0]).toMatchObject({ firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 1 });
  });

  it('says why when the name cannot be used', () => {
    const { service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('A1');

    expect(names(service).entries).toEqual([]);
    expect(names(service).refused).toContain('already means something else');
  });

  it('says why when the name is a function', () => {
    const { service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('MEDIAN');
    expect(names(service).refused).toContain('function');
  });

  it('clears the refusal once a good name is given', () => {
    const { service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('A1');
    service.defineName('Sales');
    expect(names(service).refused).toBe('');
  });

  it('lets a formula use it', () => {
    const { document, service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');
    service.setCell(0, 5, '=SUM(Sales)');
    document.sheet.recalculate();
    expect(document.sheet.value(0, 5)).toBe(30);
  });

  it('takes a definition back on undo', () => {
    const { document, service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');
    expect(document.sheet.names.rangeOf('Sales')).not.toBeNull();

    document.undo();
    expect(document.sheet.names.rangeOf('Sales')).toBeNull();

    document.redo();
    expect(document.sheet.names.rangeOf('Sales')).not.toBeNull();
  });

  /** A formula that used it has to notice, both ways. */
  /**
   * The formula is written *first*, so the last thing on the undo
   * stack is the name. Undoing it has to reach the formula, which is
   * the claim: a name is not only what a formula answers, it is what
   * the formula reads.
   */
  it('leaves a formula saying #NAME? after the definition is undone', () => {
    const { document, service } = harness();
    service.setCell(0, 5, '=SUM(Sales)');
    document.sheet.recalculate();
    expect(document.sheet.display(0, 5)).toBe('#NAME?');

    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');
    document.sheet.recalculate();
    expect(document.sheet.value(0, 5)).toBe(30);

    document.undo();
    document.sheet.recalculate();
    expect(document.sheet.display(0, 5)).toBe('#NAME?');
  });

  it('removes one', () => {
    const { document, service } = harness();
    service.setSelection(1, 1, 2, 1);
    service.defineName('Sales');
    service.removeName('SALES');
    expect(document.sheet.names.rangeOf('Sales')).toBeNull();
  });
});

/** Phase 28: a name can hold a formula, and a LAMBDA makes it a function. */
describe('naming a formula through the service', () => {
  const names = (service: SheetService) => {
    let seen: SheetNames | undefined;
    service.names.subscribe(value => (seen = value)).unsubscribe();
    return seen!;
  };

  it('makes a function the sheet can call, and publishes it apart from the ranges', () => {
    const document = new SheetDocument();
    document.setCell(0, 0, '21');
    const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
    service.defineFormulaName('Double', '=LAMBDA(x, x*2)');
    service.setCell(0, 1, '=Double(A1)');
    document.sheet.recalculate();
    expect(document.sheet.value(0, 1)).toBe(42);
    expect(names(service).entries).toEqual([]);
    expect(names(service).formulas).toEqual([{ name: 'Double', formula: '=LAMBDA(x, x*2)' }]);
  });

  it('says why when the formula does not parse', () => {
    const service = new SheetService(new SheetDocument(), { rowCount: 200, columnCount: 20 });
    service.defineFormulaName('Double', '=LAMBDA(x, x*');
    expect(names(service).formulas).toEqual([]);
    expect(names(service).refused).toContain('does not parse');
  });

  it('takes it back on undo, and the caller notices', () => {
    const document = new SheetDocument();
    const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
    service.setCell(0, 0, '=Rate*100');
    service.defineFormulaName('Rate', '=0.2');
    document.sheet.recalculate();
    expect(document.sheet.value(0, 0)).toBe(20);
    document.undo();
    document.sheet.recalculate();
    expect(document.sheet.display(0, 0)).toBe('#NAME?');
  });

  it('comes back with the sheet from a saved file', () => {
    const document = new SheetDocument();
    document.defineFormulaName('Double', '=LAMBDA(x, x*2)');
    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(snapshotOf(document)), 20)!);
    reopened.setCell(0, 0, '=Double(4)');
    reopened.sheet.recalculate();
    expect(reopened.sheet.value(0, 0)).toBe(8);
  });
});

describe('a name in a saved file', () => {
  it('comes back with the sheet', () => {
    const document = new SheetDocument();
    document.setCell(1, 1, '10');
    document.defineName('Sales', area(1, 1, 9, 1));

    const reopened = new SheetDocument();
    applySnapshot(reopened, parseSnapshot(JSON.stringify(snapshotOf(document)), 20)!);

    expect(reopened.sheet.names.rangeOf('Sales')).toMatchObject({
      start: { row: 1, column: 1 },
      end: { row: 9, column: 1 }
    });
  });

  it('is absent from a file written before names existed', () => {
    const document = new SheetDocument();
    const stored = JSON.parse(JSON.stringify(snapshotOf(document))) as Record<string, unknown>;
    delete stored.names;
    expect(parseSnapshot(JSON.stringify(stored), 20)?.names).toEqual([]);
  });

  /** A file is untrusted input in exactly the way a keystroke is. */
  it('drops a name a file should not have held', () => {
    const document = new SheetDocument();
    const stored = JSON.parse(JSON.stringify(snapshotOf(document))) as Record<string, unknown>;
    stored.names = [
      { name: 'A1', firstRow: 0, firstColumn: 0, lastRow: 0, lastColumn: 0 },
      { name: 'Fine', firstRow: 0, firstColumn: 0, lastRow: 1, lastColumn: 1 },
      { name: 'Broken', firstRow: -1, firstColumn: 0, lastRow: 1, lastColumn: 1 }
    ];
    expect(parseSnapshot(JSON.stringify(stored), 20)?.names.map(entry => entry.name)).toEqual(['Fine']);
  });
});

describe('the name box', () => {
  let ui: Rendered;
  let served: ServedForTest;

  afterEach(() => {
    ui?.unmount();
    served?.dispose();
  });

  async function mount(fill?: (document: SheetDocument) => void) {
    const document = new SheetDocument();
    fill?.(document);
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
    served = serveForTest([sheetChannel(service)]);
    ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 900, height: 400 });
    await ui.settle();
    await served.settle();
    await ui.settle();
    return { document, service };
  }

  async function typeIntoBox(text: string): Promise<void> {
    // Tabbed to rather than focused directly, because focus is the
    // input stack's and a field that was never entered has no caret
    // for the typing to land in.
    const box = ui.getByRole('textbox', { name: 'Name box' });
    let stops = 0;
    while (ui.runtime.input.focus.focusedNode !== box) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the name box');
      }
      ui.fireEvent.tab();
      await ui.settle();
    }
    ui.fireEvent.type(text);
    await ui.settle();
    ui.fireEvent.press('Enter');
    await ui.settle();
    await served.settle();
    await ui.settle();
  }

  it('names the selected range', async () => {
    const { document, service } = await mount();
    service.setSelection(1, 1, 4, 2);
    await served.settle();
    await ui.settle();

    await typeIntoBox('Sales');
    expect(document.sheet.names.rangeOf('Sales')).toMatchObject({
      start: { row: 1, column: 1 },
      end: { row: 4, column: 2 }
    });
  });

  /** Almost always a slip: the selection was a range a moment ago. */
  it('does not name a single cell', async () => {
    const { document, service } = await mount();
    service.setSelection(1, 1, 1, 1);
    await served.settle();
    await ui.settle();

    await typeIntoBox('Sales');
    expect(document.sheet.names.rangeOf('Sales')).toBeNull();
  });

  /** `Name=formula`: the same gesture, with the formula in place of the selection. */
  it('names a formula typed after an =', async () => {
    const { document, service } = await mount(d => d.setCell(0, 0, '5'));
    service.setSelection(1, 1, 1, 1);
    await served.settle();
    await ui.settle();

    await typeIntoBox('Triple=LAMBDA(x, x*3)');
    expect(document.sheet.names.formulaOf('Triple')).not.toBeNull();
    service.setCell(2, 2, '=Triple(A1)');
    document.sheet.recalculate();
    expect(document.sheet.value(2, 2)).toBe(15);
  });

  it('goes to a name it already knows', async () => {
    const { document } = await mount(d => d.defineName('Sales', area(3, 2, 6, 4)));
    await typeIntoBox('Sales');

    expect(document.selection.row).toBe(3);
    expect(document.selection.column).toBe(2);
    expect(document.selection.anchorRow).toBe(6);
  });

  it('finds it whatever case it was typed in', async () => {
    const { document } = await mount(d => d.defineName('Sales', area(3, 2, 6, 4)));
    await typeIntoBox('SALES');
    expect(document.selection.row).toBe(3);
  });

  it('still goes to a plain address', async () => {
    const { document } = await mount();
    await typeIntoBox('C9');
    expect(document.selection.row).toBe(8);
    expect(document.selection.column).toBe(2);
  });

  /**
   * Tabbing in selects what is there, so typing replaces it.
   *
   * Without this the caret sat at the start and typing `C9` over
   * `A1` gave `C9A1`, which is not an address and not a name and
   * therefore did nothing whatsoever.
   */
  it('replaces what it was showing rather than typing into it', async () => {
    await mount();
    const box = ui.getByRole('textbox', { name: 'Name box' });
    let stops = 0;
    while (ui.runtime.input.focus.focusedNode !== box) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the name box');
      }
      ui.fireEvent.tab();
      await ui.settle();
    }
    ui.fireEvent.type('C9');
    await ui.settle();
    expect(box.properties.get('value')).toBe('C9');
  });
});
