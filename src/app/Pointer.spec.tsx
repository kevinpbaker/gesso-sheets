import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';

import { COLUMN_WIDTH } from './dimensions';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * The pointer — Phase 18.
 *
 * What a hand that knows Excel does with a mouse before it has decided
 * to: click a column's letter to select the column, drag across the
 * letters for several, double-click an edge to fit it, double-click
 * the fill handle to fill down. Every one of these did nothing before
 * this phase, and none of them is a feature anybody would list.
 */

const VIEWPORT = { width: 700, height: 331 };
const ROWS = 500;
const COLUMNS = 40;

interface Harness {
  ui: Rendered;
  served: ServedForTest;
  service: SheetService;
  document: SheetDocument;
}

async function mount(fill?: (document: SheetDocument) => void, viewport = VIEWPORT): Promise<Harness> {
  const document = new SheetDocument();
  fill?.(document);
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: ROWS, columnCount: COLUMNS });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(SheetApp), { channels: served.registry, ...viewport });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served, service, document };
}

let h: Harness;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
});

async function settle(): Promise<void> {
  await h.ui.settle();
  await h.served.settle();
  await h.ui.settle();
  await h.served.settle();
  await h.ui.settle();
}

const column = (name: string) => h.ui.getByRole('columnheader', { name });
const row = (name: string) => h.ui.getByRole('rowheader', { name });
const centre = (node: ReturnType<typeof column>) => {
  const box = h.ui.getVisibleBox(node);
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
};

/** A press and release at a point, through the hit tester. */
function clickAt(point: { x: number; y: number }): void {
  h.ui.fireEvent.pointerDown(point.x, point.y, { buttons: 1 });
  h.ui.fireEvent.pointerUp(point.x, point.y);
}

describe('clicking a header', () => {
  it('selects the whole column under a letter, with the cursor at its top', async () => {
    h = await mount();
    h.ui.fireEvent.click(column('C'));
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 2, anchorRow: ROWS - 1, anchorColumn: 2 });
  });

  it('selects the whole row under a number', async () => {
    h = await mount();
    h.ui.fireEvent.click(row('3'));
    await settle();

    expect(h.document.selection).toEqual({ row: 2, column: 0, anchorRow: 2, anchorColumn: COLUMNS - 1 });
  });

  it('extends from the column it was on with Shift', async () => {
    h = await mount();
    h.ui.fireEvent.click(column('C'));
    h.ui.fireEvent.click(column('E'), { modifiers: { shift: true } });
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 2, anchorRow: ROWS - 1, anchorColumn: 2, cornerRow: 0, cornerColumn: 4 });
  });

  it('extends backwards as well as forwards', async () => {
    h = await mount();
    h.ui.fireEvent.click(row('5'));
    h.ui.fireEvent.click(row('2'), { modifiers: { shift: true } });
    await settle();

    expect(h.document.selection).toEqual({ row: 4, column: 0, anchorRow: 4, anchorColumn: COLUMNS - 1, cornerRow: 1, cornerColumn: 0 });
  });

  it('selects the sheet from the corner', async () => {
    h = await mount();
    h.ui.fireEvent.click(h.ui.getByRole('button', { name: 'Select all' }));
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 0, anchorRow: ROWS - 1, anchorColumn: COLUMNS - 1 });
  });

  /**
   * Every command already written for "the selected columns" starts
   * working from the mouse the moment a letter selects one, which is
   * most of what makes this worth doing.
   */
  it('is what Hide columns then hides', async () => {
    h = await mount(d => d.setCell(0, 1, 'B'));
    h.ui.fireEvent.click(column('B'));
    await settle();
    h.service.hideColumns(1, 1);
    await settle();

    expect(h.document.columnWidths[1]).toBe(0);
  });

  it('commits a cell that was being typed into', async () => {
    h = await mount(d => d.setCell(0, 0, 'start'));
    h.ui.fireEvent.click(h.ui.getByRole('cell', { name: 'start' }));
    await settle();
    // A letter over a selected cell opens it with that letter; the rest
    // is typed into the cell it opened.
    h.ui.fireEvent.press('k');
    await settle();
    h.ui.fireEvent.type('ept');
    await settle();
    h.ui.fireEvent.click(column('C'));
    await settle();

    expect(h.document.sheet.input(0, 0)).toBe('kept');
  });

  it('takes the keyboard back to the sheet', async () => {
    h = await mount();
    h.ui.fireEvent.click(column('C'));
    await settle();

    expect(h.ui.runtime.input.focus.focusedNode).toBe(h.ui.getByRole('grid'));
  });
});

describe('dragging across the headers', () => {
  it('sweeps several columns', async () => {
    h = await mount();
    const from = centre(column('A'));
    const to = centre(column('C'));
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 10, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(to.x, to.y);
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 0, anchorRow: ROWS - 1, anchorColumn: 0, cornerRow: 0, cornerColumn: 2 });
  });

  it('sweeps several rows', async () => {
    h = await mount();
    const from = centre(row('2'));
    const to = centre(row('4'));
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x, from.y + 10, { buttons: 1 });
    h.ui.fireEvent.pointerMove(to.x, to.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(to.x, to.y);
    await settle();

    expect(h.document.selection).toEqual({ row: 1, column: 0, anchorRow: 1, anchorColumn: COLUMNS - 1, cornerRow: 3, cornerColumn: 0 });
  });

  it('keeps sweeping letters when the pointer wanders down into the cells', async () => {
    h = await mount();
    const from = centre(column('A'));
    const to = centre(column('B'));
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 10, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(to.x, to.y + 60, { buttons: 1 });
    h.ui.fireEvent.pointerUp(to.x, to.y + 60);
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 0, anchorRow: ROWS - 1, anchorColumn: 0, cornerRow: 0, cornerColumn: 1 });
  });
});

describe('the edge of a header', () => {
  /** Just inside column A's trailing edge, where its grip is. */
  const gripOfA = () => {
    const box = h.ui.getLayout(column('A'));
    return { x: box.x + box.width - 2, y: box.y + box.height / 2 };
  };

  it('does not select the column when it is clicked', async () => {
    h = await mount();
    clickAt(gripOfA());
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 0, anchorRow: 0, anchorColumn: 0 });
  });

  it('does not select anything while it is dragged', async () => {
    h = await mount();
    const from = gripOfA();
    h.ui.fireEvent.pointerDown(from.x, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 10, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerMove(from.x + 60, from.y, { buttons: 1 });
    h.ui.fireEvent.pointerUp(from.x + 60, from.y);
    await settle();

    expect(h.document.selection).toEqual({ row: 0, column: 0, anchorRow: 0, anchorColumn: 0 });
    expect(column('A')).toHaveBox({ width: COLUMN_WIDTH + 60 });
  });

  it('fits the column to what it holds on a double click', async () => {
    h = await mount(d => d.setCell(1, 0, 'a very considerably longer piece of text'));
    clickAt(gripOfA());
    clickAt(gripOfA());
    await settle();

    expect(h.ui.getLayout(column('A')).width).toBeGreaterThan(COLUMN_WIDTH);
  });

  it('fits every selected column when the edge is one of theirs', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'x');
      d.setCell(0, 1, 'y');
    });
    h.ui.fireEvent.click(column('A'));
    h.ui.fireEvent.click(column('B'), { modifiers: { shift: true } });
    await settle();
    clickAt(gripOfA());
    clickAt(gripOfA());
    await settle();

    expect(h.ui.getLayout(column('A')).width).toBeLessThan(COLUMN_WIDTH);
    expect(h.ui.getLayout(column('B')).width).toBeLessThan(COLUMN_WIDTH);
  });

  it('fits only its own column when the selection is somewhere else', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'x');
      d.setCell(0, 1, 'y');
    });
    clickAt(gripOfA());
    clickAt(gripOfA());
    await settle();

    expect(h.ui.getLayout(column('A')).width).toBeLessThan(COLUMN_WIDTH);
    expect(h.ui.getLayout(column('B')).width).toBe(COLUMN_WIDTH);
  });

  it('gives a row dragged to a height its fitted height back on a double click', async () => {
    h = await mount(d => d.setCell(0, 0, 'A1'));
    h.service.setRowHeight(0, 80);
    await settle();
    expect(row('1')).toHaveBox({ height: 80 });

    const box = h.ui.getLayout(row('1'));
    const edge = { x: box.x + box.width / 2, y: box.y + box.height - 2 };
    clickAt(edge);
    clickAt(edge);
    await settle();

    expect(h.ui.getLayout(row('1')).height).toBeLessThan(80);
  });
});

describe('a double click on the fill handle', () => {
  const handle = () => h.ui.getByRole('button', { name: 'Fill' });

  it('fills down as far as the column on the left goes', async () => {
    h = await mount(d => {
      for (let r = 0; r < 5; r++) {
        d.setCell(r, 0, String(r + 1));
      }
      d.setCell(0, 1, '=A1*2');
    });
    h.service.setSelection(0, 1, 0, 1);
    await settle();
    h.ui.fireEvent.doubleClick(handle());
    await settle();

    expect(h.document.sheet.input(4, 1)).toBe('=A5*2');
    expect(h.document.sheet.value(4, 1)).toBe(10);
    expect(h.document.sheet.input(5, 1)).toBe('');
    expect(h.document.selection).toEqual({ row: 0, column: 1, anchorRow: 4, anchorColumn: 1 });
  });

  it('follows the column on the right when the left one has nothing below', async () => {
    h = await mount(d => {
      d.setCell(0, 0, '=B1+1');
      for (let r = 0; r < 3; r++) {
        d.setCell(r, 1, String(r));
      }
    });
    h.ui.fireEvent.doubleClick(handle());
    await settle();

    expect(h.document.sheet.input(2, 0)).toBe('=B3+1');
    expect(h.document.sheet.input(3, 0)).toBe('');
  });

  it('does nothing with nothing beside it', async () => {
    h = await mount(d => d.setCell(0, 3, 'alone'));
    h.service.setSelection(0, 3, 0, 3);
    await settle();
    h.ui.fireEvent.doubleClick(handle());
    await settle();

    expect(h.document.sheet.input(1, 3)).toBe('');
    expect(h.document.selection).toEqual({ row: 0, column: 3, anchorRow: 0, anchorColumn: 3 });
  });

  it('takes it all back on one undo', async () => {
    h = await mount(d => {
      for (let r = 0; r < 4; r++) {
        d.setCell(r, 0, String(r));
      }
      d.setCell(0, 1, '=A1');
    });
    h.service.setSelection(0, 1, 0, 1);
    await settle();
    h.ui.fireEvent.doubleClick(handle());
    await settle();
    h.service.undo();
    await settle();

    expect(h.document.sheet.input(3, 1)).toBe('');
    expect(h.document.sheet.input(0, 1)).toBe('=A1');
  });
});

describe('the menu under the right button', () => {
  const items = (): string[] => {
    try {
      return h.ui.getAllByRole('menuitem').map(item => String(item.properties.get('label')));
    } catch {
      return [];
    }
  };
  const menuItem = (label: string) => h.ui.getAllByRole('menuitem').find(item => item.properties.get('label') === label)!;

  it('offers what can be done to a row, on a row number, and selects that row first', async () => {
    h = await mount();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(row('4')));
    await settle();
    expect(h.document.selection).toEqual({ row: 3, column: 0, anchorRow: 3, anchorColumn: COLUMNS - 1 });
    expect(items()).toContain('Delete rows');
    expect(items()).toContain('Hide rows');
    expect(items()).not.toContain('Delete columns');
  });

  it('offers what can be done to a column, on a column letter', async () => {
    h = await mount();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(column('C')));
    await settle();
    expect(h.document.selection).toEqual({ row: 0, column: 2, anchorRow: ROWS - 1, anchorColumn: 2 });
    expect(items()).toContain('Delete columns');
    expect(items()).not.toContain('Delete rows');
  });

  it('runs what is chosen on what it was opened over, and closes', async () => {
    h = await mount(d => {
      d.setCell(0, 0, 'one');
      d.setCell(1, 0, 'two');
    });
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(row('1')));
    await settle();
    h.ui.fireEvent.click(menuItem('Delete rows'));
    await settle();
    expect(h.document.sheet.input(0, 0)).toBe('two');
    expect(items()).toEqual([]);
  });

  it('freezes through the column it was opened on, and keeps the rows frozen', async () => {
    h = await mount();
    h.service.freeze(2, 0);
    await settle();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(column('C')));
    await settle();
    expect(items()).not.toContain('Unfreeze columns');
    h.ui.fireEvent.click(menuItem('Freeze up to these columns'));
    await settle();
    expect([h.document.frozenRows, h.document.frozenColumns]).toEqual([2, 3]);
  });

  it('unfreezes the rows from a row number, and leaves the columns frozen', async () => {
    h = await mount();
    h.service.freeze(2, 1);
    await settle();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(row('1')));
    await settle();
    h.ui.fireEvent.click(menuItem('Unfreeze rows'));
    await settle();
    expect([h.document.frozenRows, h.document.frozenColumns]).toEqual([0, 1]);
  });

  it('offers to unfreeze rows only when rows are frozen', async () => {
    h = await mount();
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), centre(row('4')));
    await settle();
    expect(items()).toContain('Freeze up to these rows');
    expect(items()).not.toContain('Unfreeze rows');
  });

  it('keeps a selection it was opened inside', async () => {
    h = await mount();
    h.service.setSelection(1, 1, 4, 3);
    await settle();
    const c3 = h.ui.getVisibleBox(column('C'));
    const r3 = h.ui.getVisibleBox(row('3'));
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), { x: c3.x + c3.width / 2, y: r3.y + r3.height / 2 });
    await settle();
    expect(h.document.selection).toEqual({ row: 1, column: 1, anchorRow: 4, anchorColumn: 3 });
    expect(items()).toContain('Cut');
  });

  it('opens from the keyboard with Shift+F10', async () => {
    h = await mount();
    const grid = h.ui.getByRole('grid');
    let stops = 0;
    while (h.ui.runtime.input.focus.focusedNode !== grid) {
      if (stops++ > 12) {
        throw new Error('Tab never reached the grid');
      }
      h.ui.fireEvent.tab();
      await h.ui.settle();
    }
    h.ui.fireEvent.press('F10', { shift: true });
    await settle();
    expect(items()).toContain('Copy');
  });
});

/**
 * The menu is on the screen, all of it.
 *
 * Eleven commands are 424 pixels of menu. Opened from the middle of a
 * window 600 tall it fits neither below the pointer nor above it, and
 * the engine put it on one side anyway with a piece hanging off the
 * window — in a browser 813 tall, its first command off the top.
 * Since Gesso 0.6.17 the engine slides it back over the pointer until
 * all of it is on the screen, as a native menu does: its bottom at the
 * window's, 658.4.
 */
describe('where the menu under the right button opens', () => {
  /** Raised with the title bar and the toolbar's own row in Phases 38 and 40, so the grid is the height it was. */
  const TALL = { width: 700, height: 658.4 };
  const menu = () => h.ui.getVisibleBox(h.ui.getByRole('menu', { name: 'Cell actions' }));
  /** A right-click in the middle of a row's height, 24 pixels each from 152.8. */
  const rightClickRow = async (rowName: string) => {
    const box = h.ui.getVisibleBox(row(rowName));
    h.ui.fireEvent.contextMenu(h.ui.getByRole('grid'), { x: 300, y: box.y + box.height / 2 });
    await settle();
  };
  const near = (value: number) => Math.round(value * 10) / 10;

  it('opens below the pointer when it fits below', async () => {
    h = await mount(undefined, TALL);
    await rightClickRow('1');
    expect([near(menu().y), menu().height]).toEqual([164.8, 424]);
  });

  it('opens above the pointer when it fits above and not below', async () => {
    h = await mount(undefined, TALL);
    await rightClickRow('18');
    expect(near(menu().y)).toBe(near(572.8 - 424));
  });

  it('is moved until all of it shows when it fits on neither side', async () => {
    h = await mount(undefined, TALL);
    await rightClickRow('8');
    expect(near(menu().y)).toBe(near(TALL.height - 424));
  });
});

describe('a header click on a big sheet', () => {
  it('does not scroll back to the top to put the cursor there', async () => {
    h = await mount();
    h.ui.fireEvent.wheel({ x: 300, y: 200, deltaY: 24 * 100 });
    await settle();
    const before = h.ui.getAllByRole('rowheader')[0].properties.get('label');
    h.ui.fireEvent.click(column('C'));
    await settle();
    expect(h.document.selection.column).toBe(2);
    expect(h.ui.getAllByRole('rowheader')[0].properties.get('label')).toBe(before);
  });

  /**
   * The exit's second half: a column of a sheet a million rows tall is
   * a range, and neither selecting it nor adding it up walks it.
   */
  it('selects and adds up a column of a million rows at the cost of what is in it', async () => {
    const document = new SheetDocument();
    for (let at = 0; at < 50; at++) {
      document.setCell(at, 1, String(at + 1));
    }
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 1_000_000, columnCount: COLUMNS });
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, ...VIEWPORT });
    h = { ui, served, service, document };
    await settle();
    const started = performance.now();
    h.ui.fireEvent.click(column('B'));
    await settle();
    expect(performance.now() - started).toBeLessThan(500);
    expect(h.document.selection).toEqual({ row: 0, column: 1, anchorRow: 999_999, anchorColumn: 1 });
    let sum: number | null = null;
    h.service.selectionStats.subscribe(stats => (sum = (stats as { sum: number | null }).sum)).unsubscribe();
    expect(sum).toBe(1275);
  });
});
